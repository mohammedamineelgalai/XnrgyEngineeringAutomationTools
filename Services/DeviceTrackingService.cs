using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using Microsoft.Win32;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service pour tracker les appareils connectes dans Firebase
    /// Permet de voir en temps reel quels postes utilisent l'application
    /// </summary>
    public class DeviceTrackingService : IDisposable
    {
        // URL de la Firebase Realtime Database
        private const string FIREBASE_DATABASE_URL = "https://xeat-remote-control-default-rtdb.firebaseio.com";
        
        // Intervalle de heartbeat en millisecondes (60 secondes)
        private const int HEARTBEAT_INTERVAL_MS = 60000;

        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(10)
        };

        private readonly string _deviceId;
        private readonly System.Timers.Timer _heartbeatTimer;
        private bool _isRegistered;
        private bool _disposed;

        // Singleton pour acces global
        private static DeviceTrackingService _instance;
        public static DeviceTrackingService Instance => _instance;

        /// <summary>
        /// Initialise le service de tracking
        /// </summary>
        public DeviceTrackingService()
        {
            // Generer un ID unique base sur le nom de machine
            _deviceId = GenerateDeviceId();
            
            // Timer pour le heartbeat
            _heartbeatTimer = new System.Timers.Timer(HEARTBEAT_INTERVAL_MS);
            _heartbeatTimer.Elapsed += async (s, e) => await SendHeartbeatAsync();
            _heartbeatTimer.AutoReset = true;

            _instance = this;
        }

        /// <summary>
        /// Genere un ID unique pour l'appareil base sur le nom de machine
        /// </summary>
        private string GenerateDeviceId()
        {
            string machineName = Environment.MachineName ?? "Unknown";
            string userName = Environment.UserName ?? "Unknown";
            
            // Creer un ID lisible et unique
            string combined = $"{machineName}_{userName}";
            
            // Nettoyer pour Firebase (pas de caracteres speciaux)
            return SanitizeForFirebase(combined);
        }

        /// <summary>
        /// Nettoie une chaine pour etre compatible avec les cles Firebase
        /// </summary>
        private string SanitizeForFirebase(string input)
        {
            if (string.IsNullOrEmpty(input)) return "unknown";
            
            // Firebase n'accepte pas: . # $ [ ] /
            var result = new StringBuilder();
            foreach (char c in input)
            {
                if (char.IsLetterOrDigit(c) || c == '_' || c == '-')
                {
                    result.Append(c);
                }
                else
                {
                    result.Append('_');
                }
            }
            return result.ToString();
        }

        /// <summary>
        /// Enregistre l'appareil au demarrage de l'application
        /// Structure alignee avec firebase-init.json
        /// </summary>
        public async Task RegisterDeviceAsync()
        {
            try
            {
                // Collecter toutes les informations systeme
                var systemInfo = CollectSystemInfo();
                var autodeskInfo = CollectAutodeskSoftware();
                var networkInfo = CollectNetworkInfo();
                var storageInfo = CollectStorageInfo();
                var memoryInfo = CollectMemoryInfo();
                var hardwareInfo = CollectHardwareInfo();

                string nowUtc = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                string nowLocal = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Structure alignee avec firebase-init.json
                var deviceInfo = new
                {
                    // === REGISTRATION ===
                    registration = new
                    {
                        deviceId = _deviceId,
                        machineName = Environment.MachineName,
                        registeredAt = nowUtc,
                        registeredBy = Environment.UserName,
                        approved = true,
                        approvedAt = nowUtc,
                        approvedBy = "auto"
                    },
                    
                    // === ORGANIZATION ===
                    organization = new
                    {
                        site = GetSiteFromMachineName(),
                        department = GetDepartmentFromMachineName(),
                        location = "none",
                        assignedTo = Environment.UserName,
                        assetTag = "none",
                        purchaseDate = "none",
                        warrantyExpires = "none"
                    },
                    
                    // === STATUS ===
                    status = new
                    {
                        enabled = true,
                        blocked = false,
                        blockReason = "none",
                        blockedAt = "none",
                        blockedBy = "none",
                        online = true,
                        lastSeen = nowUtc,
                        currentUser = Environment.UserName,
                        currentUserId = SanitizeForFirebase($"{Environment.UserDomainName}_{Environment.UserName}")
                    },
                    
                    // === SYSTEM ===
                    system = new
                    {
                        osName = systemInfo.OsFriendlyName,
                        osVersion = systemInfo.OsVersion,
                        osBuild = systemInfo.OsBuild,
                        osArchitecture = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit",
                        computerName = Environment.MachineName,
                        domain = Environment.UserDomainName,
                        manufacturer = hardwareInfo.Manufacturer,
                        model = hardwareInfo.Model,
                        serialNumber = hardwareInfo.SerialNumber
                    },
                    
                    // === HARDWARE ===
                    hardware = new
                    {
                        processor = systemInfo.ProcessorName,
                        processorCores = Environment.ProcessorCount,
                        processorSpeed = hardwareInfo.ProcessorSpeed,
                        ramTotalGB = hardwareInfo.RamTotalGB,
                        ramAvailableGB = hardwareInfo.RamAvailableGB,
                        gpuName = hardwareInfo.GpuName,
                        gpuMemoryGB = hardwareInfo.GpuMemoryGB,
                        diskTotalGB = GetPrimaryDiskTotalGB(storageInfo),
                        diskFreeGB = GetPrimaryDiskFreeGB(storageInfo),
                        diskType = hardwareInfo.DiskType
                    },
                    
                    // === NETWORK ===
                    network = new
                    {
                        ipAddressLocal = GetNetworkValue(networkInfo, "ipAddress"),
                        ipAddressPublic = "none",
                        macAddress = GetNetworkValue(networkInfo, "macAddress"),
                        hostname = Environment.MachineName,
                        domainName = Environment.UserDomainName,
                        dnsServers = "none",
                        networkSpeed = GetNetworkValue(networkInfo, "speedDescription")
                    },
                    
                    // === SOFTWARE ===
                    software = new
                    {
                        xeatVersion = GetAppVersion(),
                        xeatInstalledAt = nowUtc,
                        xeatLastUpdated = nowUtc,
                        inventorVersion = GetAutodeskValue(autodeskInfo, "inventor", "version"),
                        inventorYear = GetAutodeskValue(autodeskInfo, "inventor", "year"),
                        vaultVersion = GetAutodeskValue(autodeskInfo, "vault", "version"),
                        vaultServer = GetVaultServerName(),
                        dotnetVersion = Environment.Version.ToString(),
                        officeVersion = GetOfficeVersion()
                    },
                    
                    // === HEARTBEAT ===
                    heartbeat = new
                    {
                        lastHeartbeat = nowUtc,
                        intervalMs = HEARTBEAT_INTERVAL_MS,
                        missedHeartbeats = 0,
                        status = "online",
                        cpuUsage = GetCurrentCpuUsage(),
                        ramUsage = GetCurrentRamUsage(),
                        diskUsage = GetPrimaryDiskUsage(storageInfo)
                    },
                    
                    // === USAGE ===
                    usage = new
                    {
                        totalSessions = 1,
                        totalUploads = 0,
                        totalModulesCreated = 0,
                        lastModuleCreated = "none",
                        lastUpload = "none"
                    },
                    
                    // === SECURITY ===
                    security = new
                    {
                        lastLoginAttempt = nowUtc,
                        failedLoginAttempts = 0,
                        trustedDevice = true,
                        requiresApproval = false
                    },
                    
                    // === METADATA ===
                    metadata = new
                    {
                        notes = "Auto-registered by XEAT",
                        tags = $"auto,{GetSiteFromMachineName()}",
                        createdAt = nowUtc,
                        updatedAt = nowUtc
                    }
                };

                string json = JsonSerializer.Serialize(deviceInfo, new JsonSerializerOptions { WriteIndented = false });
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                string url = $"{FIREBASE_DATABASE_URL}/devices/{_deviceId}.json";
                var response = await _httpClient.PutAsync(url, content);

                if (response.IsSuccessStatusCode)
                {
                    _isRegistered = true;
                    _heartbeatTimer.Start();
                    Logger.Log($"[+] Appareil enregistre dans Firebase: {_deviceId}");
                    
                    // Enregistrer aussi dans les statistiques
                    await IncrementStatisticAsync("statistics/devices/totalRegistered");
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    Logger.Log($"[!] Echec enregistrement Firebase: {error}", Logger.LogLevel.WARNING);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur enregistrement appareil: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Envoie un heartbeat pour indiquer que l'appareil est toujours actif
        /// Structure alignee avec firebase-init.json
        /// Verifie aussi les commandes (blocage, broadcasts) depuis la console admin
        /// </summary>
        private async Task SendHeartbeatAsync()
        {
            if (!_isRegistered) return;

            try
            {
                string nowUtc = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                
                // Heartbeat avec metriques de performance
                var heartbeat = new
                {
                    lastHeartbeat = nowUtc,
                    intervalMs = HEARTBEAT_INTERVAL_MS,
                    missedHeartbeats = 0,
                    status = "online",
                    cpuUsage = GetCurrentCpuUsage(),
                    ramUsage = GetCurrentRamUsage(),
                    diskUsage = GetPrimaryDiskUsageSimple()
                };

                string json = JsonSerializer.Serialize(heartbeat);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // PUT sur /devices/{id}/heartbeat.json
                string url = $"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/heartbeat.json";
                await _httpClient.PutAsync(url, content);
                
                // Aussi mettre a jour le status.online et status.lastSeen via PUT individuel
                // (PATCH non disponible en .NET 4.8 HttpClient standard)
                var onlineContent = new StringContent("true", Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status/online.json", onlineContent);
                
                var lastSeenContent = new StringContent($"\"{nowUtc}\"", Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status/lastSeen.json", lastSeenContent);

                Logger.Log($"[~] Heartbeat envoye: {_deviceId}", Logger.LogLevel.DEBUG);
                
                // === VERIFICATION DES COMMANDES ADMIN ===
                await CheckRemoteCommandsAsync();
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur heartbeat: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        /// <summary>
        /// Verifie les commandes distantes (blocage, broadcasts) depuis la console admin
        /// Appele a chaque heartbeat (toutes les 60 secondes)
        /// </summary>
        private async Task CheckRemoteCommandsAsync()
        {
            try
            {
                // Verifier si le device est bloque
                var statusResponse = await _httpClient.GetStringAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status.json");
                if (!string.IsNullOrEmpty(statusResponse) && statusResponse != "null")
                {
                    var status = JsonSerializer.Deserialize<DeviceStatusCheck>(statusResponse, new JsonSerializerOptions 
                    { 
                        PropertyNameCaseInsensitive = true 
                    });
                    
                    if (status?.Blocked == true)
                    {
                        string reason = status.BlockReason ?? "Suspendu par l'administrateur";
                        Logger.Log($"[-] DEVICE BLOQUE PAR ADMIN: {reason}", Logger.LogLevel.ERROR);
                        
                        // Afficher le message et fermer l'application
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            System.Windows.MessageBox.Show(
                                $"Ce poste a ete bloque par l'administrateur.\n\nRaison: {reason}\n\nL'application va se fermer.",
                                "Acces refuse - Poste bloque",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Error);
                            
                            System.Windows.Application.Current.Shutdown();
                        });
                        return;
                    }
                }
                
                // Verifier les broadcasts actifs
                await CheckBroadcastsAsync();
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification commandes: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        /// <summary>
        /// Verifie et affiche les broadcasts actifs
        /// </summary>
        private async Task CheckBroadcastsAsync()
        {
            try
            {
                var broadcastsResponse = await _httpClient.GetStringAsync($"{FIREBASE_DATABASE_URL}/broadcasts.json");
                if (string.IsNullOrEmpty(broadcastsResponse) || broadcastsResponse == "null") return;
                
                var broadcasts = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(broadcastsResponse);
                if (broadcasts == null) return;
                
                foreach (var kvp in broadcasts)
                {
                    if (kvp.Key == "placeholder") continue;
                    
                    try
                    {
                        var broadcast = kvp.Value;
                        
                        // Verifier si actif (status.active)
                        bool isActive = false;
                        if (broadcast.TryGetProperty("status", out var statusEl))
                        {
                            if (statusEl.TryGetProperty("active", out var activeEl))
                            {
                                isActive = activeEl.GetBoolean();
                            }
                        }
                        
                        if (!isActive) continue;
                        
                        // Verifier si doit afficher en popup (display.showAsPopup)
                        bool showPopup = false;
                        if (broadcast.TryGetProperty("display", out var displayEl))
                        {
                            if (displayEl.TryGetProperty("showAsPopup", out var popupEl))
                            {
                                showPopup = popupEl.GetBoolean();
                            }
                        }
                        
                        if (!showPopup) continue;
                        
                        // Verifier si deja affiche (eviter doublons)
                        string broadcastId = kvp.Key;
                        if (_displayedBroadcasts.Contains(broadcastId)) continue;
                        
                        // Recuperer titre et message
                        string title = broadcast.TryGetProperty("title", out var titleEl) ? titleEl.GetString() : "Message";
                        string message = broadcast.TryGetProperty("message", out var msgEl) ? msgEl.GetString() : "";
                        string type = broadcast.TryGetProperty("type", out var typeEl) ? typeEl.GetString() : "info";
                        
                        // Marquer comme affiche
                        _displayedBroadcasts.Add(broadcastId);
                        
                        Logger.Log($"[i] Broadcast popup: {title}");
                        
                        // Afficher le message
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            var icon = type switch
                            {
                                "warning" => System.Windows.MessageBoxImage.Warning,
                                "error" => System.Windows.MessageBoxImage.Error,
                                _ => System.Windows.MessageBoxImage.Information
                            };
                            
                            System.Windows.MessageBox.Show(
                                message,
                                $"[Admin] {title}",
                                System.Windows.MessageBoxButton.OK,
                                icon);
                        });
                        
                        // Incrementer le viewCount
                        await IncrementBroadcastViewCountAsync(broadcastId);
                    }
                    catch { /* Ignorer les broadcasts mal formes */ }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur check broadcasts: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        /// <summary>
        /// Incremente le compteur de vues d'un broadcast
        /// </summary>
        private async Task IncrementBroadcastViewCountAsync(string broadcastId)
        {
            try
            {
                // Lire le viewCount actuel
                var response = await _httpClient.GetStringAsync($"{FIREBASE_DATABASE_URL}/broadcasts/{broadcastId}/status/viewCount.json");
                int currentCount = 0;
                if (!string.IsNullOrEmpty(response) && response != "null")
                {
                    int.TryParse(response, out currentCount);
                }
                
                // Incrementer
                var content = new StringContent((currentCount + 1).ToString(), Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/broadcasts/{broadcastId}/status/viewCount.json", content);
            }
            catch { }
        }
        
        // Liste des broadcasts deja affiches pour eviter les doublons
        private readonly HashSet<string> _displayedBroadcasts = new HashSet<string>();
        
        /// <summary>
        /// Classe interne pour deserialiser le status du device
        /// </summary>
        private class DeviceStatusCheck
        {
            public bool Blocked { get; set; }
            public string BlockReason { get; set; }
            public bool Enabled { get; set; } = true;
        }
        
        /// <summary>
        /// Obtient l'utilisation disque de facon simple (sans l'objet storageInfo)
        /// </summary>
        private int GetPrimaryDiskUsageSimple()
        {
            try
            {
                var drive = new DriveInfo("C");
                if (drive.IsReady)
                {
                    double totalGB = drive.TotalSize / 1024.0 / 1024.0 / 1024.0;
                    double freeGB = drive.AvailableFreeSpace / 1024.0 / 1024.0 / 1024.0;
                    return (int)(((totalGB - freeGB) * 100) / totalGB);
                }
            }
            catch { }
            return 0;
        }

        /// <summary>
        /// Desenregistre l'appareil a la fermeture de l'application
        /// </summary>
        public async Task UnregisterDeviceAsync()
        {
            if (!_isRegistered) return;

            try
            {
                _heartbeatTimer.Stop();

                string nowUtc = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

                // Mettre a jour status.online = false et status.lastSeen
                var onlineContent = new StringContent("false", Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status/online.json", onlineContent);
                
                var lastSeenContent = new StringContent($"\"{nowUtc}\"", Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status/lastSeen.json", lastSeenContent);

                // Mettre a jour heartbeat.status = offline
                var heartbeatStatus = new
                {
                    lastHeartbeat = nowUtc,
                    status = "offline",
                    missedHeartbeats = 0
                };
                string heartbeatJson = JsonSerializer.Serialize(heartbeatStatus);
                var heartbeatContent = new StringContent(heartbeatJson, Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/heartbeat.json", heartbeatContent);

                _isRegistered = false;
                Logger.Log($"[+] Appareil desenregistre: {_deviceId}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur desenregistrement: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Obtient la version de l'application
        /// </summary>
        private string GetAppVersion()
        {
            try
            {
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "1.0.0";
            }
            catch
            {
                return "1.0.0";
            }
        }

        /// <summary>
        /// Detecte le site XNRGY base sur le nom de machine
        /// Utilise les identifiants du firebase-init.json (saintHubertQC, arizonaUS, lavalQC)
        /// </summary>
        private string GetSiteFromMachineName()
        {
            string machine = Environment.MachineName?.ToUpperInvariant() ?? "";
            
            if (machine.Contains("LAV") || machine.Contains("LAVAL"))
                return "lavalQC";
            if (machine.Contains("AZ") || machine.Contains("ARIZONA") || machine.Contains("PHX") || machine.Contains("PHOENIX"))
                return "arizonaUS";
            if (machine.Contains("HUB") || machine.Contains("STH") || machine.Contains("SH"))
                return "saintHubertQC";
            
            // Default to Saint-Hubert (primary site)
            return "saintHubertQC";
        }

        #region System Info Collection

        /// <summary>
        /// Collecte les informations systeme generales
        /// </summary>
        private (string OsFriendlyName, string OsVersion, string OsBuild, string ProcessorName) CollectSystemInfo()
        {
            string osFriendlyName = "Windows";
            string osVersion = "";
            string osBuild = "";
            string processorName = "Unknown";

            try
            {
                // Nom convivial de l'OS
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (key != null)
                    {
                        string productName = key.GetValue("ProductName")?.ToString() ?? "";
                        string displayVersion = key.GetValue("DisplayVersion")?.ToString() ?? "";
                        string currentBuild = key.GetValue("CurrentBuild")?.ToString() ?? "";
                        
                        osFriendlyName = string.IsNullOrEmpty(displayVersion) 
                            ? productName 
                            : $"{productName} ({displayVersion})";
                        osVersion = displayVersion;
                        osBuild = currentBuild;
                    }
                }

                // Nom du processeur
                using (var searcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Processor"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        processorName = obj["Name"]?.ToString()?.Trim() ?? "Unknown";
                        break;
                    }
                }
            }
            catch { }

            return (osFriendlyName, osVersion, osBuild, processorName);
        }
        
        /// <summary>
        /// Collecte les informations hardware detaillees (manufacturer, model, GPU, etc.)
        /// </summary>
        private (string Manufacturer, string Model, string SerialNumber, string ProcessorSpeed, 
                 double RamTotalGB, double RamAvailableGB, string GpuName, double GpuMemoryGB, string DiskType) CollectHardwareInfo()
        {
            string manufacturer = "Unknown";
            string model = "Unknown";
            string serialNumber = "Unknown";
            string processorSpeed = "Unknown";
            double ramTotalGB = 0;
            double ramAvailableGB = 0;
            string gpuName = "Unknown";
            double gpuMemoryGB = 0;
            string diskType = "Unknown";

            try
            {
                // Manufacturer et Model
                using (var searcher = new ManagementObjectSearcher("SELECT Manufacturer, Model FROM Win32_ComputerSystem"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        manufacturer = obj["Manufacturer"]?.ToString() ?? "Unknown";
                        model = obj["Model"]?.ToString() ?? "Unknown";
                        break;
                    }
                }

                // Serial Number
                using (var searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        serialNumber = obj["SerialNumber"]?.ToString() ?? "Unknown";
                        break;
                    }
                }

                // Processor Speed
                using (var searcher = new ManagementObjectSearcher("SELECT MaxClockSpeed FROM Win32_Processor"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        int speed = Convert.ToInt32(obj["MaxClockSpeed"]);
                        processorSpeed = $"{speed / 1000.0:F2} GHz";
                        break;
                    }
                }

                // RAM
                using (var searcher = new ManagementObjectSearcher("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        ulong totalKB = Convert.ToUInt64(obj["TotalVisibleMemorySize"]);
                        ulong freeKB = Convert.ToUInt64(obj["FreePhysicalMemory"]);
                        ramTotalGB = Math.Round(totalKB / 1024.0 / 1024.0, 1);
                        ramAvailableGB = Math.Round(freeKB / 1024.0 / 1024.0, 1);
                        break;
                    }
                }

                // GPU
                using (var searcher = new ManagementObjectSearcher("SELECT Name, AdapterRAM FROM Win32_VideoController"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        gpuName = obj["Name"]?.ToString() ?? "Unknown";
                        try
                        {
                            ulong adapterRam = Convert.ToUInt64(obj["AdapterRAM"]);
                            gpuMemoryGB = Math.Round(adapterRam / 1024.0 / 1024.0 / 1024.0, 1);
                        }
                        catch { }
                        break;
                    }
                }

                // Disk Type (SSD vs HDD)
                using (var searcher = new ManagementObjectSearcher("SELECT MediaType FROM Win32_DiskDrive WHERE Index=0"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        string mediaType = obj["MediaType"]?.ToString() ?? "";
                        diskType = mediaType.Contains("SSD") ? "SSD" : 
                                   mediaType.Contains("HDD") ? "HDD" : 
                                   DetectDiskTypeAlternative();
                        break;
                    }
                }
            }
            catch { }

            return (manufacturer, model, serialNumber, processorSpeed, ramTotalGB, ramAvailableGB, gpuName, gpuMemoryGB, diskType);
        }
        
        /// <summary>
        /// Detection alternative du type de disque si MediaType n'est pas disponible
        /// </summary>
        private string DetectDiskTypeAlternative()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT Model FROM Win32_DiskDrive WHERE Index=0"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        string model = obj["Model"]?.ToString()?.ToUpper() ?? "";
                        if (model.Contains("SSD") || model.Contains("NVME") || model.Contains("SOLID"))
                            return "SSD";
                        if (model.Contains("HDD") || model.Contains("HARD"))
                            return "HDD";
                    }
                }
            }
            catch { }
            return "Unknown";
        }
        
        /// <summary>
        /// Detecte le departement base sur le nom de machine
        /// </summary>
        private string GetDepartmentFromMachineName()
        {
            string machine = Environment.MachineName?.ToUpperInvariant() ?? "";
            
            if (machine.Contains("ENG") || machine.Contains("CAD") || machine.Contains("DESIGN"))
                return "engineering";
            if (machine.Contains("PROD") || machine.Contains("SHOP"))
                return "production";
            if (machine.Contains("QA") || machine.Contains("QC") || machine.Contains("QUAL"))
                return "quality";
            if (machine.Contains("IT") || machine.Contains("ADMIN") || machine.Contains("SRV"))
                return "it";
            
            return "engineering"; // Default pour XNRGY
        }
        
        /// <summary>
        /// Obtient le nom du serveur Vault depuis le fichier de config ou le registre
        /// </summary>
        private string GetVaultServerName()
        {
            try
            {
                // Essayer de lire depuis le registre Vault
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Autodesk\VaultClient"))
                {
                    if (key != null)
                    {
                        return key.GetValue("LastServer")?.ToString() ?? "XNRGY-SRV";
                    }
                }
            }
            catch { }
            return "XNRGY-SRV";
        }
        
        /// <summary>
        /// Obtient la version de Microsoft Office installee
        /// </summary>
        private string GetOfficeVersion()
        {
            try
            {
                // Office 365/2019/2021
                string[] paths = {
                    @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                    @"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
                    @"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot"
                };
                
                foreach (var path in paths)
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(path))
                    {
                        if (key != null)
                        {
                            string version = key.GetValue("VersionToReport")?.ToString() ?? 
                                           key.GetValue("ProductReleaseIds")?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(version))
                                return version;
                        }
                    }
                }
            }
            catch { }
            return "none";
        }
        
        /// <summary>
        /// Obtient l'utilisation CPU actuelle
        /// </summary>
        private int GetCurrentCpuUsage()
        {
            try
            {
                using (var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total"))
                {
                    cpuCounter.NextValue();
                    Thread.Sleep(100);
                    return (int)cpuCounter.NextValue();
                }
            }
            catch { }
            return 0;
        }
        
        /// <summary>
        /// Obtient l'utilisation RAM actuelle en pourcentage
        /// </summary>
        private int GetCurrentRamUsage()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        ulong totalKB = Convert.ToUInt64(obj["TotalVisibleMemorySize"]);
                        ulong freeKB = Convert.ToUInt64(obj["FreePhysicalMemory"]);
                        return (int)(((totalKB - freeKB) * 100) / totalKB);
                    }
                }
            }
            catch { }
            return 0;
        }
        
        /// <summary>
        /// Helper pour obtenir une valeur du dictionnaire network
        /// </summary>
        private string GetNetworkValue(object networkInfo, string key)
        {
            try
            {
                var dict = networkInfo as IDictionary<string, object>;
                if (dict != null && dict.TryGetValue(key, out var value))
                    return value?.ToString() ?? "none";
                    
                // Fallback avec reflexion pour les types anonymes
                var prop = networkInfo.GetType().GetProperty(key);
                if (prop != null)
                    return prop.GetValue(networkInfo)?.ToString() ?? "none";
            }
            catch { }
            return "none";
        }
        
        /// <summary>
        /// Helper pour obtenir une valeur de l'objet autodesk
        /// </summary>
        private string GetAutodeskValue(object autodeskInfo, string software, string property)
        {
            try
            {
                var softwareProp = autodeskInfo.GetType().GetProperty(software);
                if (softwareProp != null)
                {
                    var softwareObj = softwareProp.GetValue(autodeskInfo);
                    if (softwareObj != null)
                    {
                        var valueProp = softwareObj.GetType().GetProperty(property);
                        if (valueProp != null)
                            return valueProp.GetValue(softwareObj)?.ToString() ?? "none";
                    }
                }
            }
            catch { }
            return "none";
        }
        
        /// <summary>
        /// Helper pour obtenir la taille totale du disque principal
        /// </summary>
        private double GetPrimaryDiskTotalGB(object storageInfo)
        {
            try
            {
                var drivesProp = storageInfo.GetType().GetProperty("drives");
                if (drivesProp != null)
                {
                    var drives = drivesProp.GetValue(storageInfo) as System.Collections.IEnumerable;
                    if (drives != null)
                    {
                        foreach (var drive in drives)
                        {
                            var letterProp = drive.GetType().GetProperty("letter");
                            if (letterProp?.GetValue(drive)?.ToString()?.StartsWith("C") == true)
                            {
                                var totalProp = drive.GetType().GetProperty("totalGB");
                                return Convert.ToDouble(totalProp?.GetValue(drive) ?? 0);
                            }
                        }
                    }
                }
            }
            catch { }
            return 0;
        }
        
        /// <summary>
        /// Helper pour obtenir l'espace libre du disque principal
        /// </summary>
        private double GetPrimaryDiskFreeGB(object storageInfo)
        {
            try
            {
                var drivesProp = storageInfo.GetType().GetProperty("drives");
                if (drivesProp != null)
                {
                    var drives = drivesProp.GetValue(storageInfo) as System.Collections.IEnumerable;
                    if (drives != null)
                    {
                        foreach (var drive in drives)
                        {
                            var letterProp = drive.GetType().GetProperty("letter");
                            if (letterProp?.GetValue(drive)?.ToString()?.StartsWith("C") == true)
                            {
                                var freeProp = drive.GetType().GetProperty("freeGB");
                                return Convert.ToDouble(freeProp?.GetValue(drive) ?? 0);
                            }
                        }
                    }
                }
            }
            catch { }
            return 0;
        }
        
        /// <summary>
        /// Helper pour obtenir l'utilisation du disque principal
        /// </summary>
        private int GetPrimaryDiskUsage(object storageInfo)
        {
            try
            {
                var drivesProp = storageInfo.GetType().GetProperty("drives");
                if (drivesProp != null)
                {
                    var drives = drivesProp.GetValue(storageInfo) as System.Collections.IEnumerable;
                    if (drives != null)
                    {
                        foreach (var drive in drives)
                        {
                            var letterProp = drive.GetType().GetProperty("letter");
                            if (letterProp?.GetValue(drive)?.ToString()?.StartsWith("C") == true)
                            {
                                var usageProp = drive.GetType().GetProperty("usagePercent");
                                return Convert.ToInt32(usageProp?.GetValue(drive) ?? 0);
                            }
                        }
                    }
                }
            }
            catch { }
            return 0;
        }
        
        /// <summary>
        /// Incremente une statistique dans Firebase
        /// </summary>
        private async Task IncrementStatisticAsync(string path)
        {
            try
            {
                // Lire la valeur actuelle
                var response = await _httpClient.GetAsync($"{FIREBASE_DATABASE_URL}/{path}.json");
                int currentValue = 0;
                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(content) && content != "null")
                    {
                        int.TryParse(content, out currentValue);
                    }
                }
                
                // Incrementer
                currentValue++;
                
                // Sauvegarder
                var putContent = new StringContent(currentValue.ToString(), Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/{path}.json", putContent);
            }
            catch { }
        }

        /// <summary>
        /// Collecte les informations memoire
        /// </summary>
        private object CollectMemoryInfo()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        ulong totalKB = Convert.ToUInt64(obj["TotalVisibleMemorySize"]);
                        ulong freeKB = Convert.ToUInt64(obj["FreePhysicalMemory"]);
                        ulong usedKB = totalKB - freeKB;
                        
                        double totalGB = totalKB / 1024.0 / 1024.0;
                        double usedGB = usedKB / 1024.0 / 1024.0;
                        double freeGB = freeKB / 1024.0 / 1024.0;
                        int usagePercent = (int)((usedKB * 100) / totalKB);

                        return new
                        {
                            totalGB = Math.Round(totalGB, 1),
                            usedGB = Math.Round(usedGB, 1),
                            freeGB = Math.Round(freeGB, 1),
                            usagePercent = usagePercent,
                            status = usagePercent > 90 ? "critical" : usagePercent > 75 ? "warning" : "ok"
                        };
                    }
                }
            }
            catch { }

            return new { totalGB = 0, usedGB = 0, freeGB = 0, usagePercent = 0, status = "unknown" };
        }

        /// <summary>
        /// Collecte les informations de stockage (disques)
        /// </summary>
        private object CollectStorageInfo()
        {
            var drives = new List<object>();

            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    if (drive.IsReady && drive.DriveType == DriveType.Fixed)
                    {
                        double totalGB = drive.TotalSize / 1024.0 / 1024.0 / 1024.0;
                        double freeGB = drive.AvailableFreeSpace / 1024.0 / 1024.0 / 1024.0;
                        double usedGB = totalGB - freeGB;
                        int usagePercent = (int)((usedGB * 100) / totalGB);

                        drives.Add(new
                        {
                            letter = drive.Name.TrimEnd('\\'),
                            label = string.IsNullOrEmpty(drive.VolumeLabel) ? "Local Disk" : drive.VolumeLabel,
                            totalGB = Math.Round(totalGB, 1),
                            usedGB = Math.Round(usedGB, 1),
                            freeGB = Math.Round(freeGB, 1),
                            usagePercent = usagePercent,
                            status = usagePercent > 95 ? "critical" : usagePercent > 85 ? "warning" : "ok"
                        });
                    }
                }

                // Verifier specifiquement C:\Vault
                string vaultPath = @"C:\Vault";
                bool vaultExists = Directory.Exists(vaultPath);
                long vaultSizeMB = 0;
                
                if (vaultExists)
                {
                    try
                    {
                        var vaultDir = new DirectoryInfo(vaultPath);
                        vaultSizeMB = GetDirectorySizeMB(vaultDir);
                    }
                    catch { }
                }

                return new
                {
                    drives = drives,
                    vaultFolder = new
                    {
                        exists = vaultExists,
                        path = vaultPath,
                        sizeMB = vaultSizeMB,
                        sizeGB = Math.Round(vaultSizeMB / 1024.0, 2)
                    }
                };
            }
            catch { }

            return new { drives = drives, vaultFolder = new { exists = false, path = "", sizeMB = 0, sizeGB = 0 } };
        }

        /// <summary>
        /// Calcule la taille d'un dossier en MB (limite a 2 niveaux pour performance)
        /// </summary>
        private long GetDirectorySizeMB(DirectoryInfo dir, int depth = 0)
        {
            long size = 0;
            try
            {
                // Limiter la profondeur pour eviter les timeouts
                if (depth > 2) return 0;

                foreach (var file in dir.GetFiles())
                {
                    size += file.Length;
                }

                foreach (var subDir in dir.GetDirectories())
                {
                    size += GetDirectorySizeMB(subDir, depth + 1);
                }
            }
            catch { }

            return size / 1024 / 1024; // Convertir en MB
        }

        /// <summary>
        /// Collecte les informations reseau
        /// </summary>
        private object CollectNetworkInfo()
        {
            try
            {
                var activeInterface = NetworkInterface.GetAllNetworkInterfaces()
                    .FirstOrDefault(n => n.OperationalStatus == OperationalStatus.Up 
                                      && n.NetworkInterfaceType != NetworkInterfaceType.Loopback);

                if (activeInterface != null)
                {
                    var ipProps = activeInterface.GetIPProperties();
                    var ipv4 = ipProps.UnicastAddresses
                        .FirstOrDefault(a => a.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);

                    var stats = activeInterface.GetIPv4Statistics();

                    // Test de connectivite internet
                    bool internetConnected = false;
                    long pingMs = -1;
                    
                    try
                    {
                        using (var ping = new Ping())
                        {
                            var reply = ping.Send("8.8.8.8", 3000);
                            if (reply.Status == IPStatus.Success)
                            {
                                internetConnected = true;
                                pingMs = reply.RoundtripTime;
                            }
                        }
                    }
                    catch { }

                    // Vitesse de l'interface
                    long speedMbps = activeInterface.Speed / 1000000;

                    return new
                    {
                        interfaceName = activeInterface.Name,
                        interfaceType = activeInterface.NetworkInterfaceType.ToString(),
                        ipAddress = ipv4?.Address.ToString() ?? "Unknown",
                        macAddress = FormatMacAddress(activeInterface.GetPhysicalAddress().ToString()),
                        speedMbps = speedMbps,
                        speedDescription = speedMbps >= 1000 ? $"{speedMbps / 1000} Gbps" : $"{speedMbps} Mbps",
                        internetConnected = internetConnected,
                        pingMs = pingMs,
                        pingStatus = pingMs < 0 ? "offline" : pingMs < 50 ? "excellent" : pingMs < 100 ? "good" : pingMs < 200 ? "fair" : "poor",
                        bytesSent = FormatBytes(stats.BytesSent),
                        bytesReceived = FormatBytes(stats.BytesReceived),
                        status = internetConnected ? "connected" : "disconnected"
                    };
                }
            }
            catch { }

            return new
            {
                interfaceName = "Unknown",
                interfaceType = "Unknown",
                ipAddress = "Unknown",
                macAddress = "Unknown",
                speedMbps = 0,
                speedDescription = "Unknown",
                internetConnected = false,
                pingMs = -1,
                pingStatus = "unknown",
                bytesSent = "0 B",
                bytesReceived = "0 B",
                status = "unknown"
            };
        }

        /// <summary>
        /// Formate une adresse MAC
        /// </summary>
        private string FormatMacAddress(string mac)
        {
            if (string.IsNullOrEmpty(mac) || mac.Length != 12) return mac;
            return string.Join(":", Enumerable.Range(0, 6).Select(i => mac.Substring(i * 2, 2)));
        }

        /// <summary>
        /// Formate les bytes en unite lisible
        /// </summary>
        private string FormatBytes(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            double size = bytes;
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }
            return $"{size:0.##} {sizes[order]}";
        }

        /// <summary>
        /// Collecte les logiciels Autodesk installes EN VERIFIANT QUE LES FICHIERS EXISTENT REELLEMENT
        /// Ne se fie PAS uniquement au registre (peut contenir des anciennes installations)
        /// </summary>
        private object CollectAutodeskSoftware()
        {
            try
            {
                // === DETECTION INVENTOR - VERIFICATION PHYSIQUE ===
                var inventorInfo = DetectInventorInstallation();
                
                // === DETECTION VAULT - VERIFICATION PHYSIQUE ===
                var vaultInfo = DetectVaultInstallation();
                
                // === AUTRES LOGICIELS AUTODESK (seulement si installPath existe) ===
                var otherSoftware = DetectOtherAutodeskSoftware();
                
                // === PROCESSUS EN COURS ===
                bool inventorRunning = Process.GetProcessesByName("Inventor").Length > 0;
                bool vaultRunning = Process.GetProcessesByName("Connectivity.VaultPro").Length > 0 ||
                                   Process.GetProcessesByName("Autodesk.Connectivity.Explorer").Length > 0;

                return new
                {
                    inventor = new
                    {
                        installed = inventorInfo.Installed,
                        version = inventorInfo.Version,
                        year = inventorInfo.Year,
                        path = inventorInfo.Path,
                        running = inventorRunning,
                        verifiedOnDisk = inventorInfo.Installed // Toujours verifie physiquement
                    },
                    vault = new
                    {
                        installed = vaultInfo.Installed,
                        version = vaultInfo.Version,
                        path = vaultInfo.Path,
                        running = vaultRunning,
                        verifiedOnDisk = vaultInfo.Installed
                    },
                    allSoftware = otherSoftware.OrderBy(s => s.Name).ToList(),
                    totalAutodeskProducts = otherSoftware.Count + (inventorInfo.Installed ? 1 : 0) + (vaultInfo.Installed ? 1 : 0)
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    inventor = new { installed = false, version = "", year = "", path = "", running = false, verifiedOnDisk = false },
                    vault = new { installed = false, version = "", path = "", running = false, verifiedOnDisk = false },
                    allSoftware = new List<object>(),
                    totalAutodeskProducts = 0,
                    error = ex.Message
                };
            }
        }

        /// <summary>
        /// Detecte Inventor en verifiant que l'executable existe reellement
        /// </summary>
        private (bool Installed, string Version, string Year, string Path) DetectInventorInstallation()
        {
            // Chemins standards Inventor (du plus recent au plus ancien)
            var inventorPaths = new[]
            {
                (@"C:\Program Files\Autodesk\Inventor 2026\Bin\Inventor.exe", "2026"),
                (@"C:\Program Files\Autodesk\Inventor 2025\Bin\Inventor.exe", "2025"),
                (@"C:\Program Files\Autodesk\Inventor 2024\Bin\Inventor.exe", "2024"),
                (@"C:\Program Files\Autodesk\Inventor 2023\Bin\Inventor.exe", "2023"),
                (@"C:\Program Files\Autodesk\Inventor 2022\Bin\Inventor.exe", "2022"),
            };

            foreach (var (exePath, year) in inventorPaths)
            {
                if (File.Exists(exePath))
                {
                    // Obtenir la version du fichier
                    string version = "";
                    try
                    {
                        var versionInfo = FileVersionInfo.GetVersionInfo(exePath);
                        version = versionInfo.FileVersion ?? "";
                    }
                    catch { }

                    string installPath = Path.GetDirectoryName(Path.GetDirectoryName(exePath)) ?? "";
                    
                    return (true, version, year, installPath);
                }
            }

            return (false, "", "", "");
        }

        /// <summary>
        /// Detecte Vault en verifiant que l'executable existe reellement
        /// </summary>
        private (bool Installed, string Version, string Path) DetectVaultInstallation()
        {
            // Chemins standards Vault (du plus recent au plus ancien)
            var vaultPaths = new[]
            {
                @"C:\Program Files\Autodesk\Vault Client 2026\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Client 2025\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Client 2024\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Client 2023\Explorer\Connectivity.VaultPro.exe",
                // Vault Professional
                @"C:\Program Files\Autodesk\Vault Professional 2026\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Professional 2025\Explorer\Connectivity.VaultPro.exe",
            };

            foreach (var exePath in vaultPaths)
            {
                if (File.Exists(exePath))
                {
                    string version = "";
                    try
                    {
                        var versionInfo = FileVersionInfo.GetVersionInfo(exePath);
                        version = versionInfo.FileVersion ?? "";
                    }
                    catch { }

                    string installPath = Path.GetDirectoryName(Path.GetDirectoryName(exePath)) ?? "";
                    
                    return (true, version, installPath);
                }
            }

            return (false, "", "");
        }

        /// <summary>
        /// Detecte les autres logiciels Autodesk (AutoCAD, etc.) en verifiant physiquement
        /// </summary>
        private List<(string Name, string Version, string Path)> DetectOtherAutodeskSoftware()
        {
            var software = new List<(string Name, string Version, string Path)>();
            var foundPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // AutoCAD
            var autocadPaths = new[]
            {
                (@"C:\Program Files\Autodesk\AutoCAD 2026\acad.exe", "AutoCAD 2026"),
                (@"C:\Program Files\Autodesk\AutoCAD 2025\acad.exe", "AutoCAD 2025"),
                (@"C:\Program Files\Autodesk\AutoCAD 2024\acad.exe", "AutoCAD 2024"),
            };

            foreach (var (exePath, name) in autocadPaths)
            {
                if (File.Exists(exePath) && !foundPaths.Contains(exePath))
                {
                    foundPaths.Add(exePath);
                    string version = GetFileVersion(exePath);
                    software.Add((name, version, Path.GetDirectoryName(exePath) ?? ""));
                }
            }

            // Revit
            var revitPaths = new[]
            {
                (@"C:\Program Files\Autodesk\Revit 2026\Revit.exe", "Revit 2026"),
                (@"C:\Program Files\Autodesk\Revit 2025\Revit.exe", "Revit 2025"),
                (@"C:\Program Files\Autodesk\Revit 2024\Revit.exe", "Revit 2024"),
            };

            foreach (var (exePath, name) in revitPaths)
            {
                if (File.Exists(exePath) && !foundPaths.Contains(exePath))
                {
                    foundPaths.Add(exePath);
                    string version = GetFileVersion(exePath);
                    software.Add((name, version, Path.GetDirectoryName(exePath) ?? ""));
                }
            }

            // Navisworks
            var navisworksPaths = new[]
            {
                (@"C:\Program Files\Autodesk\Navisworks Manage 2026\Navisworks.exe", "Navisworks Manage 2026"),
                (@"C:\Program Files\Autodesk\Navisworks Manage 2025\Navisworks.exe", "Navisworks Manage 2025"),
            };

            foreach (var (exePath, name) in navisworksPaths)
            {
                if (File.Exists(exePath) && !foundPaths.Contains(exePath))
                {
                    foundPaths.Add(exePath);
                    string version = GetFileVersion(exePath);
                    software.Add((name, version, Path.GetDirectoryName(exePath) ?? ""));
                }
            }

            // 3ds Max
            var maxPaths = new[]
            {
                (@"C:\Program Files\Autodesk\3ds Max 2026\3dsmax.exe", "3ds Max 2026"),
                (@"C:\Program Files\Autodesk\3ds Max 2025\3dsmax.exe", "3ds Max 2025"),
            };

            foreach (var (exePath, name) in maxPaths)
            {
                if (File.Exists(exePath) && !foundPaths.Contains(exePath))
                {
                    foundPaths.Add(exePath);
                    string version = GetFileVersion(exePath);
                    software.Add((name, version, Path.GetDirectoryName(exePath) ?? ""));
                }
            }

            // Fusion 360
            string fusion360Path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"Autodesk\webdeploy\production\Fusion360.exe");
            
            if (File.Exists(fusion360Path))
            {
                string version = GetFileVersion(fusion360Path);
                software.Add(("Fusion 360", version, Path.GetDirectoryName(fusion360Path) ?? ""));
            }

            return software;
        }

        /// <summary>
        /// Obtient la version d'un fichier executable
        /// </summary>
        private string GetFileVersion(string filePath)
        {
            try
            {
                var versionInfo = FileVersionInfo.GetVersionInfo(filePath);
                return versionInfo.FileVersion ?? "";
            }
            catch
            {
                return "";
            }
        }

        #endregion

        /// <summary>
        /// Libere les ressources et met le statut offline
        /// Structure alignee avec firebase-init.json
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            
            _heartbeatTimer?.Stop();
            _heartbeatTimer?.Dispose();
            
            string nowUtc = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
            
            // Envoyer le statut offline de facon synchrone avec timeout
            try
            {
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) })
                {
                    // Mettre status.online = false
                    var onlineContent = new StringContent("false", Encoding.UTF8, "application/json");
                    var task1 = client.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status/online.json", onlineContent);
                    task1.Wait(1500);
                    
                    // Mettre status.lastSeen
                    var lastSeenContent = new StringContent($"\"{nowUtc}\"", Encoding.UTF8, "application/json");
                    var task2 = client.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status/lastSeen.json", lastSeenContent);
                    task2.Wait(1500);
                    
                    // Mettre heartbeat.status = offline
                    var heartbeatContent = new StringContent("\"offline\"", Encoding.UTF8, "application/json");
                    var task3 = client.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/heartbeat/status.json", heartbeatContent);
                    task3.Wait(1500);
                    
                    // Mettre heartbeat.lastHeartbeat
                    var lastHeartbeatContent = new StringContent($"\"{nowUtc}\"", Encoding.UTF8, "application/json");
                    var task4 = client.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/heartbeat/lastHeartbeat.json", lastHeartbeatContent);
                    task4.Wait(1500);
                }
                
                Logger.Log($"[+] Statut offline envoye pour: {_deviceId}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur envoi statut offline: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        #region Public Methods for Telemetry
        
        /// <summary>
        /// Incremente le compteur d'uploads pour cet appareil
        /// </summary>
        public async Task IncrementUploadsAsync()
        {
            try
            {
                // Lire la valeur actuelle
                var response = await _httpClient.GetAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/totalUploads.json");
                int currentValue = 0;
                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(content) && content != "null")
                    {
                        int.TryParse(content, out currentValue);
                    }
                }
                
                currentValue++;
                
                var putContent = new StringContent(currentValue.ToString(), Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/totalUploads.json", putContent);
                
                // Aussi mettre a jour lastUpload
                string nowUtc = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                var lastUploadContent = new StringContent($"\"{nowUtc}\"", Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/lastUpload.json", lastUploadContent);
                
                // Incrementer aussi dans les statistiques globales
                await IncrementStatisticAsync("statistics/usage/totalUploads");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur increment uploads: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        /// <summary>
        /// Incremente le compteur de modules crees pour cet appareil
        /// </summary>
        public async Task IncrementModulesCreatedAsync()
        {
            try
            {
                // Lire la valeur actuelle
                var response = await _httpClient.GetAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/totalModulesCreated.json");
                int currentValue = 0;
                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(content) && content != "null")
                    {
                        int.TryParse(content, out currentValue);
                    }
                }
                
                currentValue++;
                
                var putContent = new StringContent(currentValue.ToString(), Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/totalModulesCreated.json", putContent);
                
                // Aussi mettre a jour lastModuleCreated
                string nowUtc = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                var lastModuleContent = new StringContent($"\"{nowUtc}\"", Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/lastModuleCreated.json", lastModuleContent);
                
                // Incrementer aussi dans les statistiques globales
                await IncrementStatisticAsync("statistics/usage/totalModulesCreated");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur increment modules: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        /// <summary>
        /// Incremente le compteur de sessions pour cet appareil
        /// </summary>
        public async Task IncrementSessionsAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/totalSessions.json");
                int currentValue = 0;
                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(content) && content != "null")
                    {
                        int.TryParse(content, out currentValue);
                    }
                }
                
                currentValue++;
                
                var putContent = new StringContent(currentValue.ToString(), Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/usage/totalSessions.json", putContent);
                
                // Incrementer aussi dans les statistiques globales
                await IncrementStatisticAsync("statistics/sessions/total");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur increment sessions: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        /// <summary>
        /// Log un evenement de telemetrie
        /// </summary>
        public async Task LogTelemetryEventAsync(string category, string action, string label = null, int? value = null)
        {
            try
            {
                string eventId = $"evt_{DateTime.UtcNow:yyyyMMddHHmmssfff}";
                var eventData = new
                {
                    timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    deviceId = _deviceId,
                    userId = SanitizeForFirebase($"{Environment.UserDomainName}_{Environment.UserName}"),
                    category = category,
                    action = action,
                    label = label ?? "none",
                    value = value ?? 0,
                    site = GetSiteFromMachineName()
                };

                string json = JsonSerializer.Serialize(eventData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/telemetry/events/{eventId}.json", content);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur log telemetry: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        /// <summary>
        /// Log une erreur dans Firebase
        /// </summary>
        public async Task LogErrorAsync(string errorType, string message, string stackTrace = null)
        {
            try
            {
                string errorId = $"err_{DateTime.UtcNow:yyyyMMddHHmmssfff}";
                var errorData = new
                {
                    timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    deviceId = _deviceId,
                    userId = SanitizeForFirebase($"{Environment.UserDomainName}_{Environment.UserName}"),
                    errorType = errorType,
                    message = message,
                    stackTrace = stackTrace ?? "none",
                    appVersion = GetAppVersion(),
                    site = GetSiteFromMachineName()
                };

                string json = JsonSerializer.Serialize(errorData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/telemetry/errors/{errorId}.json", content);
                
                // Incrementer le compteur d'erreurs
                await IncrementStatisticAsync("statistics/errors/total");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur log error: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
        
        #endregion
    }
}
