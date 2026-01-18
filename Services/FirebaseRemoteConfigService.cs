using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using XnrgyEngineeringAutomationTools.Models;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service pour la verification de la configuration Firebase Realtime Database
    /// Permet le controle a distance de l'application (kill switch, maintenance, mises a jour)
    /// OPTIMISE: Chargement parallele et cache pour performance au demarrage
    /// </summary>
    public class FirebaseRemoteConfigService
    {
        // URL de la Firebase Realtime Database
        private const string FIREBASE_DATABASE_URL = "https://xeat-remote-control-default-rtdb.firebaseio.com";
        
        // Timeout reduit pour les requetes HTTP (5 secondes au lieu de 10)
        private static readonly TimeSpan HTTP_TIMEOUT = TimeSpan.FromSeconds(5);

        // Version actuelle de l'application
        private static readonly string CURRENT_VERSION = GetCurrentVersion();
        
        // Utilisateur et machine actuels
        private static readonly string CURRENT_USER = Environment.UserName?.ToLowerInvariant() ?? "unknown";
        private static readonly string CURRENT_DEVICE = $"{Environment.MachineName}_{Environment.UserName}".Replace(".", "_");

        // Cache des donnees Firebase pour eviter les appels multiples
        private static Dictionary<string, FirebaseDevice> _cachedDevices;
        private static Dictionary<string, FirebaseUser> _cachedUsers;
        private static Dictionary<string, BroadcastMessage> _cachedBroadcasts;
        private static WelcomeMessagesConfig _cachedWelcomeMessages;
        private static DateTime _cacheExpiry = DateTime.MinValue;
        private static readonly TimeSpan CACHE_DURATION = TimeSpan.FromMinutes(5);

        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = HTTP_TIMEOUT
        };

        /// <summary>
        /// Verifie la configuration Firebase et retourne le resultat
        /// OPTIMISE: Chargement parallele de toutes les donnees pour performance
        /// </summary>
        public static async Task<FirebaseCheckResult> CheckConfigurationAsync()
        {
            var stopwatch = Stopwatch.StartNew();
            
            try
            {
                Logger.Log("[>] Verification Firebase en cours...");

                // OPTIMISATION: Charger TOUTES les donnees en parallele en un seul bloc
                var configTask = FetchFirebaseConfigAsync();
                var devicesTask = FetchJsonAsync<Dictionary<string, FirebaseDevice>>("/devices.json");
                var usersTask = FetchJsonAsync<Dictionary<string, FirebaseUser>>("/users.json");
                var broadcastsTask = FetchJsonAsync<Dictionary<string, BroadcastMessage>>("/broadcasts.json");
                var welcomeTask = FetchJsonAsync<WelcomeMessagesConfig>("/welcomeMessages.json");

                // Attendre tous les appels en parallele (max 5 secondes grace au timeout)
                await Task.WhenAll(configTask, devicesTask, usersTask, broadcastsTask, welcomeTask);
                
                // Recuperer les resultats et mettre en cache
                var config = await configTask;
                _cachedDevices = await devicesTask;
                _cachedUsers = await usersTask;
                _cachedBroadcasts = await broadcastsTask;
                _cachedWelcomeMessages = await welcomeTask;
                _cacheExpiry = DateTime.Now.Add(CACHE_DURATION);

                Logger.Log($"[+] Donnees Firebase chargees en {stopwatch.ElapsedMilliseconds}ms");

                if (config == null)
                {
                    Logger.Log("[!] Impossible de lire la configuration Firebase - Mode hors ligne", Logger.LogLevel.WARNING);
                    return FirebaseCheckResult.CreateSuccess(); // Continuer en mode hors ligne
                }

                var result = new FirebaseCheckResult { Success = true };

                // 1. Verifier le Kill Switch global (PRIORITE HAUTE)
                if (config.Commands?.Global?.KillSwitch == true)
                {
                    result.KillSwitchActive = true;
                    result.KillSwitchMessage = config.Commands.Global.KillSwitchMessage 
                        ?? "Application desactivee par l'administrateur.";
                    Logger.Log("[-] Kill Switch actif: " + result.KillSwitchMessage, Logger.LogLevel.ERROR);
                    return result;
                }

                // 2. Verifier le DEVICE et l'UTILISATEUR sur ce device (utilise le cache)
                var deviceStatus = CheckDeviceStatusFromCache();
                
                // 2a. Device entier suspendu
                if (deviceStatus.deviceDisabled)
                {
                    result.DeviceDisabled = true;
                    result.DeviceDisabledMessage = deviceStatus.deviceMessage;
                    result.DeviceDisabledReason = deviceStatus.deviceReason;
                    Logger.Log($"[-] Device suspendu: {CURRENT_DEVICE} - {deviceStatus.deviceReason}", Logger.LogLevel.ERROR);
                    return result;
                }
                
                // 2b. Utilisateur suspendu sur ce device specifique
                if (deviceStatus.userDisabled)
                {
                    result.DeviceUserDisabled = true;
                    result.DeviceUserDisabledMessage = deviceStatus.userMessage;
                    result.DeviceUserDisabledReason = deviceStatus.userReason;
                    Logger.Log($"[-] Utilisateur suspendu sur device: {CURRENT_USER}@{Environment.MachineName} - {deviceStatus.userReason}", Logger.LogLevel.ERROR);
                    return result;
                }

                // 3. Verifier si l'utilisateur est desactive globalement (utilise le cache)
                var userStatus = CheckUserStatusFromCache();
                if (userStatus.isDisabled)
                {
                    result.UserDisabled = true;
                    result.UserDisabledMessage = userStatus.message;
                    Logger.Log($"[-] Utilisateur desactive globalement: {CURRENT_USER}", Logger.LogLevel.ERROR);
                    return result;
                }

                // 4. Verifier le mode maintenance
                if (config.AppConfig?.MaintenanceMode == true)
                {
                    result.MaintenanceMode = true;;
                    result.MaintenanceMessage = config.AppConfig.MaintenanceMessage 
                        ?? "Application en maintenance. Reessayez dans quelques minutes.";
                    Logger.Log("[!] Mode maintenance actif: " + result.MaintenanceMessage, Logger.LogLevel.WARNING);
                    return result;
                }

                // 5. Verifier les mises a jour
                string latestVersion = config.VersionInfo?.Latest?.Version;
                if (!string.IsNullOrEmpty(latestVersion))
                {
                    result.LatestVersion = latestVersion;
                    result.DownloadUrl = config.VersionInfo.Latest.DownloadUrl;
                    result.Changelog = config.VersionInfo.Latest.Changelog;

                    if (IsNewerVersion(latestVersion, CURRENT_VERSION))
                    {
                        result.UpdateAvailable = true;
                        result.ForceUpdate = config.AppConfig?.ForceUpdate == true 
                            || config.Commands?.Global?.ForceUpdate == true;

                        Logger.Log($"[+] Mise a jour disponible: {CURRENT_VERSION} -> {latestVersion}");
                        
                        if (result.ForceUpdate)
                        {
                            Logger.Log("[!] Mise a jour forcee requise", Logger.LogLevel.WARNING);
                        }
                    }
                    else
                    {
                        Logger.Log($"[+] Version a jour: {CURRENT_VERSION}");
                    }
                }

                // 6. Verifier les messages broadcast (utilise le cache)
                var broadcast = CheckBroadcastMessageFromCache();
                if (broadcast.hasMessage)
                {
                    result.HasBroadcastMessage = true;
                    result.BroadcastTitle = broadcast.title;
                    result.BroadcastMessage = broadcast.message;
                    result.BroadcastType = broadcast.type;
                    Logger.Log($"[i] Message broadcast recu: {broadcast.title}");
                }

                // 7. Verifier le message de bienvenue (utilise le cache)
                var welcome = CheckWelcomeMessageFromCache();
                if (welcome.hasMessage)
                {
                    result.HasWelcomeMessage = true;
                    result.WelcomeTitle = welcome.title;
                    result.WelcomeMessage = welcome.message;
                    result.WelcomeType = welcome.type;
                    Logger.Log($"[i] Message de bienvenue actif: {welcome.title}");
                }

                Logger.Log($"[+] Verification Firebase terminee en {stopwatch.ElapsedMilliseconds}ms");
                return result;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur Firebase ({stopwatch.ElapsedMilliseconds}ms): {ex.Message}", Logger.LogLevel.WARNING);
                // En cas d'erreur, continuer normalement (mode hors ligne)
                return FirebaseCheckResult.CreateSuccess();
            }
        }

        /// <summary>
        /// Verifie le statut utilisateur depuis le cache (pas d'appel reseau)
        /// </summary>
        private static (bool isDisabled, string message) CheckUserStatusFromCache()
        {
            try
            {
                if (_cachedUsers == null) return (false, null);

                foreach (var kvp in _cachedUsers)
                {
                    var user = kvp.Value;
                    if (user == null) continue;

                    string userEmail = user.GetEmail().ToLowerInvariant();
                    string userName = user.GetDisplayName().ToLowerInvariant();

                    if (userEmail.Contains(CURRENT_USER) || userName.Contains(CURRENT_USER) ||
                        (!string.IsNullOrEmpty(userEmail) && userEmail.Contains("@") && 
                         CURRENT_USER.Contains(userEmail.Split('@')[0])))
                    {
                        bool isBlocked = user.status?.blocked == true;
                        bool isDisabled = !user.IsEnabled();
                        
                        if (isDisabled || isBlocked)
                        {
                            string message = user.GetDisabledMessage() ?? 
                                $"Votre compte ({user.GetDisplayName()}) a ete desactive par l'administrateur.";
                            return (true, message);
                        }
                        break;
                    }
                }

                return (false, null);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification utilisateur: {ex.Message}", Logger.LogLevel.DEBUG);
                return (false, null);
            }
        }

        /// <summary>
        /// Verifie le statut device depuis le cache (pas d'appel reseau)
        /// </summary>
        private static (bool deviceDisabled, string deviceMessage, string deviceReason, 
                       bool userDisabled, string userMessage, string userReason) CheckDeviceStatusFromCache()
        {
            try
            {
                if (_cachedDevices == null) return (false, null, null, false, null, null);

                string currentUserLower = CURRENT_USER.ToLowerInvariant();
                
                foreach (var kvp in _cachedDevices)
                {
                    var deviceId = kvp.Key;
                    var device = kvp.Value;
                    if (device == null) continue;

                    bool isThisDevice = 
                        deviceId.Equals(CURRENT_DEVICE, StringComparison.OrdinalIgnoreCase) ||
                        (device.MachineName?.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase) ?? false);

                    if (isThisDevice)
                    {
                        if (device.IsBlocked())
                        {
                            string message = device.GetBlockedMessage();
                            string reason = device.Status?.BlockReason ?? device.DisabledReason ?? "suspended";
                            return (true, message, reason, false, null, null);
                        }

                        if (device.Users != null && device.Users.Count > 0)
                        {
                            foreach (var userKvp in device.Users)
                            {
                                string userId = userKvp.Key?.ToLowerInvariant() ?? "";
                                var deviceUser = userKvp.Value;
                                if (deviceUser == null) continue;

                                bool isThisUser = 
                                    userId.Contains(currentUserLower) || 
                                    currentUserLower.Contains(userId) ||
                                    userId.Equals(currentUserLower);

                                if (isThisUser && deviceUser.Enabled == false)
                                {
                                    string userReason = deviceUser.DisabledReason ?? "suspended";
                                    string userMessage = deviceUser.DisabledMessage;

                                    if (string.IsNullOrEmpty(userMessage))
                                    {
                                        switch (userReason.ToLowerInvariant())
                                        {
                                            case "unauthorized":
                                                userMessage = $"Votre compte ({CURRENT_USER}) n'est pas autorise sur ce poste.";
                                                break;
                                            case "revoked":
                                                userMessage = $"L'acces de {CURRENT_USER} a ce poste a ete revoque.";
                                                break;
                                            default:
                                                userMessage = $"Votre compte ({CURRENT_USER}) a ete suspendu sur ce poste.";
                                                break;
                                        }
                                    }

                                    return (false, null, null, true, userMessage, userReason);
                                }
                            }
                        }

                        break;
                    }
                }

                return (false, null, null, false, null, null);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification device: {ex.Message}", Logger.LogLevel.DEBUG);
                return (false, null, null, false, null, null);
            }
        }

        /// <summary>
        /// Verifie les broadcasts depuis le cache (pas d'appel reseau)
        /// </summary>
        private static (bool hasMessage, string title, string message, string type) CheckBroadcastMessageFromCache()
        {
            try
            {
                if (_cachedBroadcasts == null) return (false, null, null, null);

                long currentTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();

                foreach (var kvp in _cachedBroadcasts)
                {
                    if (kvp.Key == "placeholder") continue;
                    
                    var broadcast = kvp.Value;
                    if (broadcast == null) continue;
                    
                    if (!broadcast.IsActive()) continue;
                    if (broadcast.ExpiresAt > 0 && broadcast.ExpiresAt < currentTime) continue;

                    string targetType = broadcast.GetTargetType();
                    bool isTargeted = false;

                    switch (targetType.ToLowerInvariant())
                    {
                        case "all":
                            isTargeted = true;
                            break;
                        case "user":
                            string targetUser = broadcast.Target?.UserId ?? broadcast.TargetUser ?? "";
                            isTargeted = !string.IsNullOrEmpty(targetUser) && 
                                        targetUser.ToLowerInvariant().Contains(CURRENT_USER);
                            break;
                        case "device":
                            string targetDevice = broadcast.Target?.DeviceId ?? broadcast.TargetDevice ?? "";
                            isTargeted = !string.IsNullOrEmpty(targetDevice) && 
                                        CURRENT_DEVICE.ToLowerInvariant().Contains(targetDevice.ToLowerInvariant());
                            break;
                        default:
                            isTargeted = true;
                            break;
                    }

                    if (isTargeted)
                    {
                        return (true, broadcast.Title ?? "Message", broadcast.Message, broadcast.Type ?? "info");
                    }
                }

                return (false, null, null, null);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification broadcast: {ex.Message}", Logger.LogLevel.DEBUG);
                return (false, null, null, null);
            }
        }

        /// <summary>
        /// Verifie les messages de bienvenue depuis le cache (pas d'appel reseau)
        /// </summary>
        private static (bool hasMessage, string title, string message, string type) CheckWelcomeMessageFromCache()
        {
            try
            {
                if (_cachedWelcomeMessages == null) return (false, null, null, null);
                
                var welcomeMessages = _cachedWelcomeMessages;
                var seenMessages = LoadSeenWelcomeMessages();
                string currentUserKey = $"{Environment.MachineName}_{Environment.UserName}";
                
                // 1. Premier lancement (firstInstall)
                if (welcomeMessages.FirstInstall?.Enabled == true)
                {
                    string firstInstallKey = $"firstInstall_{currentUserKey}";
                    if (!seenMessages.Contains(firstInstallKey))
                    {
                        MarkWelcomeMessageAsSeen(firstInstallKey);
                        return (true,
                            welcomeMessages.FirstInstall.Title ?? "Bienvenue!",
                            welcomeMessages.FirstInstall.Message,
                            welcomeMessages.FirstInstall.Type ?? "info");
                    }
                }
                
                // 2. Nouvel utilisateur sur ce device
                if (welcomeMessages.NewUserOnDevice?.Enabled == true)
                {
                    string newUserKey = $"newUserOnDevice_{currentUserKey}";
                    if (!seenMessages.Contains(newUserKey))
                    {
                        MarkWelcomeMessageAsSeen(newUserKey);
                        return (true,
                            welcomeMessages.NewUserOnDevice.Title ?? "Bienvenue sur ce poste!",
                            welcomeMessages.NewUserOnDevice.Message,
                            welcomeMessages.NewUserOnDevice.Type ?? "info");
                    }
                }
                
                // 3. Message global
                if (welcomeMessages.Global?.Enabled == true)
                {
                    string globalKey = $"global_{welcomeMessages.Global.Title?.GetHashCode() ?? 0}";
                    if (!seenMessages.Contains(globalKey))
                    {
                        MarkWelcomeMessageAsSeen(globalKey);
                        return (true,
                            welcomeMessages.Global.Title ?? "Information",
                            welcomeMessages.Global.Message,
                            welcomeMessages.Global.Type ?? "info");
                    }
                }

                return (false, null, null, null);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification welcome: {ex.Message}", Logger.LogLevel.DEBUG);
                return (false, null, null, null);
            }
        }

        /// <summary>
        /// Verifie s'il y a un message broadcast actif (version async pour compatibilite)
        /// </summary>
        private static async Task<(bool hasMessage, string title, string message, string type)> CheckBroadcastMessageAsync()
        {
            // Utiliser le cache si disponible
            if (_cachedBroadcasts != null && DateTime.Now < _cacheExpiry)
            {
                return CheckBroadcastMessageFromCache();
            }
            
            // Sinon charger depuis Firebase
            _cachedBroadcasts = await FetchJsonAsync<Dictionary<string, BroadcastMessage>>("/broadcasts.json");
            return CheckBroadcastMessageFromCache();
        }

        /// <summary>
        /// Verifie le message de bienvenue (version async pour compatibilite)
        /// </summary>
        private static async Task<(bool hasMessage, string title, string message, string type)> CheckWelcomeMessageAsync()
        {
            // Utiliser le cache si disponible
            if (_cachedWelcomeMessages != null && DateTime.Now < _cacheExpiry)
            {
                return CheckWelcomeMessageFromCache();
            }
            
            // Sinon charger depuis Firebase
            _cachedWelcomeMessages = await FetchJsonAsync<WelcomeMessagesConfig>("/welcomeMessages.json");
            return CheckWelcomeMessageFromCache();
        }

        // Chemin du fichier de suivi des messages de bienvenue vus
        private static readonly string WelcomeSeenFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "XNRGY", "XEAT", "welcome_seen.json");

        /// <summary>
        /// Charge la liste des messages de bienvenue deja vus
        /// </summary>
        private static List<string> LoadSeenWelcomeMessages()
        {
            try
            {
                if (File.Exists(WelcomeSeenFilePath))
                {
                    string json = File.ReadAllText(WelcomeSeenFilePath);
                    return JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur lecture welcome_seen: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return new List<string>();
        }

        /// <summary>
        /// Marque un message de bienvenue comme vu
        /// </summary>
        private static void MarkWelcomeMessageAsSeen(string messageKey)
        {
            try
            {
                var seenMessages = LoadSeenWelcomeMessages();
                if (!seenMessages.Contains(messageKey))
                {
                    seenMessages.Add(messageKey);
                    
                    string dir = Path.GetDirectoryName(WelcomeSeenFilePath);
                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }
                    
                    File.WriteAllText(WelcomeSeenFilePath, JsonSerializer.Serialize(seenMessages));
                    Logger.Log($"[i] Message de bienvenue marque comme vu: {messageKey}");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur sauvegarde welcome_seen: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }

        /// <summary>
        /// Recupere la configuration depuis Firebase
        /// OPTIMISE: Charge plusieurs endpoints en parallele
        /// </summary>
        private static async Task<FirebaseConfig> FetchFirebaseConfigAsync()
        {
            try
            {
                // Lire les endpoints necessaires en parallele
                var appConfigTask = FetchJsonAsync<AppConfig>("/appConfig.json");
                var killSwitchTask = FetchJsonAsync<KillSwitchConfig>("/killSwitch.json");
                var maintenanceTask = FetchJsonAsync<MaintenanceConfig>("/maintenance.json");
                var forceUpdateTask = FetchJsonAsync<ForceUpdateConfig>("/forceUpdate.json");
                var updatesTask = FetchJsonAsync<UpdatesConfig>("/updates.json");

                await Task.WhenAll(appConfigTask, killSwitchTask, maintenanceTask, forceUpdateTask, updatesTask);

                // Construire le resultat avec la bonne structure
                var appConfig = await appConfigTask ?? new AppConfig();
                var killSwitch = await killSwitchTask;
                var maintenance = await maintenanceTask;
                var forceUpdate = await forceUpdateTask;
                var updates = await updatesTask;
                
                // Mapper vers les anciennes structures pour compatibilite
                // Kill Switch: /killSwitch/global/enabled -> Commands.Global.KillSwitch
                if (killSwitch?.Global != null && killSwitch.Global.Enabled)
                {
                    appConfig.MaintenanceMode = false; // Ne pas confondre avec maintenance
                }
                
                // Maintenance: /maintenance/enabled -> AppConfig.MaintenanceMode
                if (maintenance?.Enabled == true)
                {
                    appConfig.MaintenanceMode = true;
                    appConfig.MaintenanceMessage = maintenance.Message;
                }
                
                // Force Update: /forceUpdate/enabled -> AppConfig.ForceUpdate
                if (forceUpdate?.Enabled == true)
                {
                    appConfig.ForceUpdate = true;
                    appConfig.MinVersion = forceUpdate.MinimumVersion;
                }
                
                // Version: /updates/latest/version -> VersionInfo.Latest.Version
                var versionInfo = new VersionInfo();
                if (updates?.Latest != null)
                {
                    versionInfo.Latest = new LatestVersion
                    {
                        Version = updates.Latest.Version,
                        DownloadUrl = updates.Latest.DownloadUrl,
                        Changelog = updates.Latest.ReleaseNotes,
                        ReleaseDate = updates.Latest.PublishedAt
                    };
                }
                
                // Construire la config avec le kill switch
                var commands = new GlobalCommands
                {
                    Global = new GlobalCommandSettings
                    {
                        KillSwitch = killSwitch?.Global?.Enabled ?? false,
                        KillSwitchMessage = killSwitch?.Global?.Message ?? "Application desactivee par administrateur",
                        ForceUpdate = forceUpdate?.Enabled ?? false
                    }
                };

                return new FirebaseConfig
                {
                    AppConfig = appConfig,
                    Commands = commands,
                    VersionInfo = versionInfo
                };
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur lecture Firebase: {ex.Message}", Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// Effectue une requete GET vers Firebase et deserialise le JSON
        /// </summary>
        private static async Task<T> FetchJsonAsync<T>(string endpoint) where T : class
        {
            try
            {
                string url = FIREBASE_DATABASE_URL + endpoint;
                string json = await _httpClient.GetStringAsync(url);

                if (string.IsNullOrEmpty(json) || json == "null")
                    return null;

                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                return JsonSerializer.Deserialize<T>(json, options);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Compare deux versions semver (ex: "1.0.0" vs "1.1.0")
        /// </summary>
        private static bool IsNewerVersion(string latestVersion, string currentVersion)
        {
            try
            {
                // Nettoyer les prefixes "v" si presents
                latestVersion = latestVersion?.TrimStart('v', 'V') ?? "0.0.0";
                currentVersion = currentVersion?.TrimStart('v', 'V') ?? "0.0.0";

                var latest = Version.Parse(latestVersion);
                var current = Version.Parse(currentVersion);

                return latest > current;
            }
            catch
            {
                // En cas d'erreur de parsing, comparer comme strings
                return string.Compare(latestVersion, currentVersion, StringComparison.OrdinalIgnoreCase) > 0;
            }
        }

        /// <summary>
        /// Obtient la version actuelle de l'application depuis l'assembly
        /// </summary>
        private static string GetCurrentVersion()
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                return $"{version.Major}.{version.Minor}.{version.Build}";
            }
            catch
            {
                return "1.0.0";
            }
        }

        /// <summary>
        /// Ouvre l'URL de telechargement dans le navigateur
        /// </summary>
        public static void OpenDownloadUrl(string url)
        {
            try
            {
                if (!string.IsNullOrEmpty(url))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur ouverture URL: {ex.Message}", Logger.LogLevel.ERROR);
            }
        }

        /// <summary>
        /// Ouvre la page des releases GitHub
        /// </summary>
        public static void OpenReleasesPage()
        {
            OpenDownloadUrl("https://github.com/mohammedamineelgalai/XnrgyEngineeringAutomationTools/releases/latest");
        }
    }
}
