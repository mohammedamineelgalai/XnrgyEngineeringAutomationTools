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

                var deviceInfo = new
                {
                    // === INFORMATIONS MACHINE ===
                    machineName = Environment.MachineName,
                    userName = Environment.UserName,
                    userDomainName = Environment.UserDomainName,
                    osVersion = Environment.OSVersion.ToString(),
                    osFriendlyName = systemInfo.OsFriendlyName,
                    processorCount = Environment.ProcessorCount,
                    processorName = systemInfo.ProcessorName,
                    is64BitOS = Environment.Is64BitOperatingSystem,
                    systemDirectory = Environment.SystemDirectory,
                    
                    // === INFORMATIONS APPLICATION ===
                    appVersion = GetAppVersion(),
                    dotNetVersion = Environment.Version.ToString(),
                    
                    // === INFORMATIONS SESSION ===
                    startTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    startTimeLocal = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    lastHeartbeat = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    status = "online",
                    
                    // === INFORMATIONS XNRGY ===
                    site = GetSiteFromMachineName(),
                    workingDirectory = Environment.CurrentDirectory,
                    
                    // === MEMOIRE ===
                    memory = memoryInfo,
                    
                    // === STOCKAGE ===
                    storage = storageInfo,
                    
                    // === RESEAU ===
                    network = networkInfo,
                    
                    // === LOGICIELS AUTODESK ===
                    autodesk = autodeskInfo
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
        /// </summary>
        private async Task SendHeartbeatAsync()
        {
            if (!_isRegistered) return;

            try
            {
                // Utiliser PUT sur un sous-noeud pour mise a jour partielle (compatible .NET 4.8)
                var heartbeat = new
                {
                    lastHeartbeat = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    status = "online"
                };

                string json = JsonSerializer.Serialize(heartbeat);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // PUT sur /devices/{id}/heartbeat.json pour mise a jour partielle
                string url = $"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/heartbeat.json";
                await _httpClient.PutAsync(url, content);

                Logger.Log($"[~] Heartbeat envoye: {_deviceId}", Logger.LogLevel.DEBUG);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur heartbeat: {ex.Message}", Logger.LogLevel.DEBUG);
            }
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

                // Mettre le statut principal a "offline"
                var offlineData = new
                {
                    status = "offline",
                    lastHeartbeat = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    disconnectTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    disconnectTimeLocal = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                };

                string json = JsonSerializer.Serialize(offlineData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // Mettre a jour le statut a la racine du device (pas dans heartbeat)
                string url = $"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/session.json";
                var response = await _httpClient.PutAsync(url, content);

                // Aussi mettre a jour le statut principal
                var statusContent = new StringContent("\"offline\"", Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status.json", statusContent);

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
        /// </summary>
        private string GetSiteFromMachineName()
        {
            string machine = Environment.MachineName?.ToUpperInvariant() ?? "";
            
            if (machine.Contains("LAV") || machine.Contains("LAVAL"))
                return "Laval";
            if (machine.Contains("MTL") || machine.Contains("MONTREAL"))
                return "Montreal";
            if (machine.Contains("QC") || machine.Contains("QUEBEC"))
                return "Quebec";
            if (machine.Contains("TOR") || machine.Contains("TORONTO"))
                return "Toronto";
            
            return "Unknown";
        }

        #region System Info Collection

        /// <summary>
        /// Collecte les informations systeme generales
        /// </summary>
        private (string OsFriendlyName, string ProcessorName) CollectSystemInfo()
        {
            string osFriendlyName = "Windows";
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
                        osFriendlyName = string.IsNullOrEmpty(displayVersion) 
                            ? productName 
                            : $"{productName} ({displayVersion})";
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

            return (osFriendlyName, processorName);
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
        /// Collecte les logiciels Autodesk installes
        /// </summary>
        private object CollectAutodeskSoftware()
        {
            var software = new List<object>();
            var vaultInfo = new { installed = false, version = "", path = "" };
            var inventorInfo = new { installed = false, version = "", year = "", path = "" };

            try
            {
                // Parcourir le registre pour trouver les logiciels Autodesk
                string[] registryPaths = new[]
                {
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                };

                var foundSoftware = new HashSet<string>();

                foreach (string regPath in registryPaths)
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(regPath))
                    {
                        if (key == null) continue;

                        foreach (string subKeyName in key.GetSubKeyNames())
                        {
                            using (var subKey = key.OpenSubKey(subKeyName))
                            {
                                if (subKey == null) continue;

                                string displayName = subKey.GetValue("DisplayName")?.ToString() ?? "";
                                string publisher = subKey.GetValue("Publisher")?.ToString() ?? "";
                                string version = subKey.GetValue("DisplayVersion")?.ToString() ?? "";
                                string installLocation = subKey.GetValue("InstallLocation")?.ToString() ?? "";

                                // Filtrer les logiciels Autodesk
                                if (publisher.Contains("Autodesk") && !string.IsNullOrEmpty(displayName))
                                {
                                    // Eviter les doublons
                                    string key2 = $"{displayName}|{version}";
                                    if (foundSoftware.Contains(key2)) continue;
                                    foundSoftware.Add(key2);

                                    // Filtrer les composants internes
                                    if (displayName.Contains("Material Library") ||
                                        displayName.Contains("Content Pack") ||
                                        displayName.Contains("Genuine Service") ||
                                        displayName.Contains("Single Sign") ||
                                        displayName.Contains("Identity Manager") ||
                                        displayName.Contains("Licensing") ||
                                        displayName.Contains("Application Manager"))
                                        continue;

                                    // Detecter Inventor
                                    if (displayName.Contains("Inventor") && !displayName.Contains("Content") && !displayName.Contains("Add-in"))
                                    {
                                        var yearMatch = System.Text.RegularExpressions.Regex.Match(displayName, @"20\d{2}");
                                        inventorInfo = new
                                        {
                                            installed = true,
                                            version = version,
                                            year = yearMatch.Success ? yearMatch.Value : "",
                                            path = installLocation
                                        };
                                    }

                                    // Detecter Vault
                                    if (displayName.Contains("Vault") && (displayName.Contains("Professional") || displayName.Contains("Client")))
                                    {
                                        vaultInfo = new
                                        {
                                            installed = true,
                                            version = version,
                                            path = installLocation
                                        };
                                    }

                                    software.Add(new
                                    {
                                        name = displayName,
                                        version = version,
                                        installPath = installLocation
                                    });
                                }
                            }
                        }
                    }
                }

                // Verifier si Inventor et Vault sont en cours d'execution
                bool inventorRunning = Process.GetProcessesByName("Inventor").Length > 0;
                bool vaultRunning = Process.GetProcessesByName("Connectivity.VaultPro").Length > 0 ||
                                   Process.GetProcessesByName("Autodesk.Connectivity.Explorer").Length > 0;

                // Verifier les chemins Inventor/Vault standards
                string inventorPath2026 = @"C:\Program Files\Autodesk\Inventor 2026\Bin\Inventor.exe";
                string inventorPath2025 = @"C:\Program Files\Autodesk\Inventor 2025\Bin\Inventor.exe";
                string inventorPath2024 = @"C:\Program Files\Autodesk\Inventor 2024\Bin\Inventor.exe";

                string detectedInventorPath = "";
                if (File.Exists(inventorPath2026)) detectedInventorPath = inventorPath2026;
                else if (File.Exists(inventorPath2025)) detectedInventorPath = inventorPath2025;
                else if (File.Exists(inventorPath2024)) detectedInventorPath = inventorPath2024;

                return new
                {
                    inventor = new
                    {
                        inventorInfo.installed,
                        inventorInfo.version,
                        inventorInfo.year,
                        path = !string.IsNullOrEmpty(inventorInfo.path) ? inventorInfo.path : detectedInventorPath,
                        running = inventorRunning
                    },
                    vault = new
                    {
                        vaultInfo.installed,
                        vaultInfo.version,
                        vaultInfo.path,
                        running = vaultRunning
                    },
                    allSoftware = software.OrderBy(s => ((dynamic)s).name).ToList(),
                    totalAutodeskProducts = software.Count
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    inventor = new { installed = false, version = "", year = "", path = "", running = false },
                    vault = new { installed = false, version = "", path = "", running = false },
                    allSoftware = new List<object>(),
                    totalAutodeskProducts = 0,
                    error = ex.Message
                };
            }
        }

        #endregion

        /// <summary>
        /// Libere les ressources et met le statut offline
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            
            _heartbeatTimer?.Stop();
            _heartbeatTimer?.Dispose();
            
            // Envoyer le statut offline de facon synchrone avec timeout
            try
            {
                // Utiliser un HttpClient synchrone pour s'assurer que ca part avant la fermeture
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) })
                {
                    var statusContent = new StringContent("\"offline\"", Encoding.UTF8, "application/json");
                    var task = client.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/status.json", statusContent);
                    task.Wait(3000);
                    
                    var disconnectContent = new StringContent(
                        $"{{\"time\":\"{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}\",\"timeLocal\":\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\"}}",
                        Encoding.UTF8, "application/json");
                    var task2 = client.PutAsync($"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/lastDisconnect.json", disconnectContent);
                    task2.Wait(2000);
                }
                
                Logger.Log($"[+] Statut offline envoye pour: {_deviceId}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur envoi statut offline: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }
    }
}
