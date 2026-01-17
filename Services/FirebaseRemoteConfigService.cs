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
    /// </summary>
    public class FirebaseRemoteConfigService
    {
        // URL de la Firebase Realtime Database
        private const string FIREBASE_DATABASE_URL = "https://xeat-remote-control-default-rtdb.firebaseio.com";
        
        // Timeout pour les requetes HTTP
        private static readonly TimeSpan HTTP_TIMEOUT = TimeSpan.FromSeconds(10);

        // Version actuelle de l'application
        private static readonly string CURRENT_VERSION = GetCurrentVersion();
        
        // Utilisateur et machine actuels
        private static readonly string CURRENT_USER = Environment.UserName?.ToLowerInvariant() ?? "unknown";
        private static readonly string CURRENT_DEVICE = $"{Environment.MachineName}_{Environment.UserName}".Replace(".", "_");

        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = HTTP_TIMEOUT
        };

        /// <summary>
        /// Verifie la configuration Firebase et retourne le resultat
        /// </summary>
        public static async Task<FirebaseCheckResult> CheckConfigurationAsync()
        {
            try
            {
                Logger.Log("[>] Verification Firebase en cours...");

                // Lire les donnees Firebase
                var config = await FetchFirebaseConfigAsync();
                if (config == null)
                {
                    Logger.Log("[!] Impossible de lire la configuration Firebase - Mode hors ligne", Logger.LogLevel.WARNING);
                    return FirebaseCheckResult.CreateSuccess(); // Continuer en mode hors ligne
                }

                var result = new FirebaseCheckResult { Success = true };

                // 1. Verifier le Kill Switch global
                if (config.Commands?.Global?.KillSwitch == true)
                {
                    result.KillSwitchActive = true;
                    result.KillSwitchMessage = config.Commands.Global.KillSwitchMessage 
                        ?? "Application desactivee par l'administrateur.";
                    Logger.Log("[-] Kill Switch actif: " + result.KillSwitchMessage, Logger.LogLevel.ERROR);
                    return result;
                }

                // 2. Verifier si l'utilisateur est desactive
                var userStatus = await CheckUserStatusAsync();
                if (userStatus.isDisabled)
                {
                    result.UserDisabled = true;
                    result.UserDisabledMessage = userStatus.message;
                    Logger.Log($"[-] Utilisateur desactive: {CURRENT_USER}", Logger.LogLevel.ERROR);
                    return result;
                }

                // 3. Verifier le mode maintenance
                if (config.AppConfig?.MaintenanceMode == true)
                {
                    result.MaintenanceMode = true;
                    result.MaintenanceMessage = config.AppConfig.MaintenanceMessage 
                        ?? "Application en maintenance. Reessayez dans quelques minutes.";
                    Logger.Log("[!] Mode maintenance actif: " + result.MaintenanceMessage, Logger.LogLevel.WARNING);
                    return result;
                }

                // 4. Verifier les mises a jour
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

                // 5. Verifier les messages broadcast
                var broadcast = await CheckBroadcastMessageAsync();
                if (broadcast.hasMessage)
                {
                    result.HasBroadcastMessage = true;
                    result.BroadcastTitle = broadcast.title;
                    result.BroadcastMessage = broadcast.message;
                    result.BroadcastType = broadcast.type;
                    Logger.Log($"[i] Message broadcast recu: {broadcast.title}");
                }

                return result;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur Firebase: {ex.Message}", Logger.LogLevel.WARNING);
                // En cas d'erreur, continuer normalement (mode hors ligne)
                return FirebaseCheckResult.CreateSuccess();
            }
        }

        /// <summary>
        /// Verifie si l'utilisateur actuel est desactive
        /// </summary>
        private static async Task<(bool isDisabled, string message)> CheckUserStatusAsync()
        {
            try
            {
                // Chercher l'utilisateur par email ou username
                var users = await FetchJsonAsync<Dictionary<string, FirebaseUser>>("/users.json");
                if (users == null) return (false, null);

                foreach (var kvp in users)
                {
                    var user = kvp.Value;
                    if (user == null) continue;

                    // Verifier par email ou nom d'utilisateur
                    string userEmail = user.Email?.ToLowerInvariant() ?? "";
                    string userName = user.DisplayName?.ToLowerInvariant() ?? "";

                    if (userEmail.Contains(CURRENT_USER) || userName.Contains(CURRENT_USER) ||
                        CURRENT_USER.Contains(userEmail.Split('@')[0]))
                    {
                        if (!user.Enabled)
                        {
                            string message = user.DisabledMessage ?? 
                                $"Votre compte ({user.DisplayName ?? CURRENT_USER}) a ete desactive par l'administrateur.";
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
        /// Verifie s'il y a un message broadcast actif
        /// </summary>
        private static async Task<(bool hasMessage, string title, string message, string type)> CheckBroadcastMessageAsync()
        {
            try
            {
                var broadcasts = await FetchJsonAsync<Dictionary<string, BroadcastMessage>>("/broadcasts.json");
                if (broadcasts == null) return (false, null, null, null);

                long currentTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();

                foreach (var kvp in broadcasts)
                {
                    var broadcast = kvp.Value;
                    if (broadcast == null || !broadcast.Active) continue;

                    // Verifier l'expiration
                    if (broadcast.ExpiresAt > 0 && broadcast.ExpiresAt < currentTime) continue;

                    // Verifier si le message est cible
                    bool isTargeted = false;

                    // Message global (pas de cible)
                    if (string.IsNullOrEmpty(broadcast.TargetUser) && string.IsNullOrEmpty(broadcast.TargetDevice))
                    {
                        isTargeted = true;
                    }
                    // Cible par utilisateur
                    else if (!string.IsNullOrEmpty(broadcast.TargetUser) && 
                             broadcast.TargetUser.ToLowerInvariant().Contains(CURRENT_USER))
                    {
                        isTargeted = true;
                    }
                    // Cible par device
                    else if (!string.IsNullOrEmpty(broadcast.TargetDevice) && 
                             CURRENT_DEVICE.ToLowerInvariant().Contains(broadcast.TargetDevice.ToLowerInvariant()))
                    {
                        isTargeted = true;
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
        /// Recupere la configuration depuis Firebase
        /// </summary>
        private static async Task<FirebaseConfig> FetchFirebaseConfigAsync()
        {
            try
            {
                // Lire les endpoints necessaires en parallele
                var appConfigTask = FetchJsonAsync<AppConfig>("/appConfig.json");
                var commandsTask = FetchJsonAsync<GlobalCommands>("/commands.json");
                var versionInfoTask = FetchJsonAsync<VersionInfo>("/versionInfo.json");

                await Task.WhenAll(appConfigTask, commandsTask, versionInfoTask);

                return new FirebaseConfig
                {
                    AppConfig = await appConfigTask,
                    Commands = await commandsTask,
                    VersionInfo = await versionInfoTask
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
