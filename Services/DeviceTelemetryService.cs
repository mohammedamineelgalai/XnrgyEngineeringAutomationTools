using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service de telemetrie qui collecte les informations systeme et les envoie a Firebase
    /// Collecte: OS, Hardware, Network, Software (Inventor, Vault), Performance
    /// </summary>
    public class DeviceTelemetryService : IDisposable
    {
        private const string FIREBASE_DATABASE_URL = "https://xeat-remote-control-default-rtdb.firebaseio.com";
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        
        private readonly string _deviceId;
        private readonly string _userId;
        private Timer _heartbeatTimer;
        private bool _disposed;

        // Singleton
        private static DeviceTelemetryService _instance;
        public static DeviceTelemetryService Instance => _instance ?? (_instance = new DeviceTelemetryService());

        private DeviceTelemetryService()
        {
            _deviceId = GenerateDeviceId();
            _userId = GetCurrentUserId();
        }

        #region Device ID Generation

        private static string GenerateDeviceId()
        {
            // Format: MACHINENAME_username (sans caracteres speciaux)
            string machine = Environment.MachineName ?? "UNKNOWN";
            string user = Environment.UserName ?? "unknown";
            return SanitizeFirebaseKey($"{machine}_{user}");
        }

        private static string GetCurrentUserId()
        {
            // Convertir le username en format Firebase-safe
            string user = Environment.UserName?.ToLowerInvariant() ?? "unknown";
            return SanitizeFirebaseKey(user.Replace(".", "").Replace("@", ""));
        }

        /// <summary>
        /// Nettoie une cle pour Firebase (pas de . $ # [ ] /)
        /// </summary>
        private static string SanitizeFirebaseKey(string key)
        {
            if (string.IsNullOrEmpty(key)) return "unknown";
            return key.Replace(".", "")
                      .Replace("$", "")
                      .Replace("#", "")
                      .Replace("[", "")
                      .Replace("]", "")
                      .Replace("/", "")
                      .Replace(" ", "_");
        }

        #endregion

        #region Registration et Startup

        /// <summary>
        /// Enregistre le device au demarrage de l'application
        /// Collecte toutes les informations systeme et les envoie a Firebase
        /// </summary>
        public async Task RegisterDeviceAsync()
        {
            try
            {
                Logger.Log($"[>] Enregistrement device: {_deviceId}");

                // Collecter toutes les informations
                var deviceData = await CollectAllDeviceDataAsync();

                // Envoyer a Firebase
                await UpdateFirebaseAsync($"/devices/{_deviceId}", deviceData);

                Logger.Log($"[+] Device enregistre avec succes: {_deviceId}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur enregistrement device: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Demarre le heartbeat periodique
        /// </summary>
        public void StartHeartbeat(int intervalMs = 60000)
        {
            StopHeartbeat();
            _heartbeatTimer = new Timer(async _ => await SendHeartbeatAsync(), null, 0, intervalMs);
            Logger.Log($"[+] Heartbeat demarre (intervalle: {intervalMs}ms)");
        }

        /// <summary>
        /// Arrete le heartbeat
        /// </summary>
        public void StopHeartbeat()
        {
            _heartbeatTimer?.Dispose();
            _heartbeatTimer = null;
        }

        #endregion

        #region Data Collection

        /// <summary>
        /// Collecte toutes les donnees du device
        /// </summary>
        private async Task<Dictionary<string, object>> CollectAllDeviceDataAsync()
        {
            var data = new Dictionary<string, object>
            {
                ["registration"] = CollectRegistrationData(),
                ["organization"] = CollectOrganizationData(),
                ["status"] = CollectStatusData(),
                ["system"] = CollectSystemData(),
                ["hardware"] = await CollectHardwareDataAsync(),
                ["network"] = CollectNetworkData(),
                ["software"] = CollectSoftwareData(),
                ["heartbeat"] = CollectHeartbeatData(),
                ["usage"] = new Dictionary<string, object>
                {
                    ["totalSessions"] = 0,
                    ["totalUploads"] = 0,
                    ["totalModulesCreated"] = 0,
                    ["lastModuleCreated"] = "none",
                    ["lastUpload"] = "none"
                },
                ["security"] = new Dictionary<string, object>
                {
                    ["lastLoginAttempt"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    ["failedLoginAttempts"] = 0,
                    ["trustedDevice"] = false,
                    ["requiresApproval"] = false
                },
                ["metadata"] = new Dictionary<string, object>
                {
                    ["notes"] = "none",
                    ["tags"] = "auto-registered",
                    ["createdAt"] = DateTime.UtcNow.ToString("yyyy-MM-dd"),
                    ["updatedAt"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
                }
            };

            return data;
        }

        private Dictionary<string, object> CollectRegistrationData()
        {
            return new Dictionary<string, object>
            {
                ["deviceId"] = _deviceId,
                ["machineName"] = Environment.MachineName ?? "unknown",
                ["registeredAt"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                ["registeredBy"] = Environment.UserName ?? "unknown",
                ["approved"] = true,
                ["approvedAt"] = "auto",
                ["approvedBy"] = "system"
            };
        }

        private Dictionary<string, object> CollectOrganizationData()
        {
            // Detecter le site en fonction du nom de domaine ou IP
            string site = DetectSite();
            string department = "engineering"; // Par defaut

            return new Dictionary<string, object>
            {
                ["site"] = site,
                ["department"] = department,
                ["location"] = "none",
                ["assignedTo"] = _userId,
                ["assetTag"] = "none",
                ["purchaseDate"] = "none",
                ["warrantyExpires"] = "none"
            };
        }

        private Dictionary<string, object> CollectStatusData()
        {
            return new Dictionary<string, object>
            {
                ["enabled"] = true,
                ["blocked"] = false,
                ["blockReason"] = "none",
                ["blockedAt"] = "none",
                ["blockedBy"] = "none",
                ["online"] = true,
                ["lastSeen"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                ["currentUser"] = Environment.UserName ?? "unknown",
                ["currentUserId"] = _userId
            };
        }

        private Dictionary<string, object> CollectSystemData()
        {
            string osName = "Windows";
            string osVersion = "unknown";
            string osBuild = "unknown";

            try
            {
                // Obtenir les infos OS via Registry
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (key != null)
                    {
                        string productName = key.GetValue("ProductName")?.ToString() ?? "Windows";
                        string displayVersion = key.GetValue("DisplayVersion")?.ToString() ?? "";
                        string currentBuild = key.GetValue("CurrentBuild")?.ToString() ?? "";
                        string ubr = key.GetValue("UBR")?.ToString() ?? "";

                        osName = productName;
                        osVersion = displayVersion;
                        osBuild = string.IsNullOrEmpty(ubr) ? currentBuild : $"{currentBuild}.{ubr}";
                    }
                }
            }
            catch { }

            string manufacturer = "unknown";
            string model = "unknown";
            string serialNumber = "unknown";

            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        manufacturer = obj["Manufacturer"]?.ToString() ?? "unknown";
                        model = obj["Model"]?.ToString() ?? "unknown";
                        break;
                    }
                }

                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        serialNumber = obj["SerialNumber"]?.ToString() ?? "unknown";
                        break;
                    }
                }
            }
            catch { }

            return new Dictionary<string, object>
            {
                ["osName"] = osName,
                ["osVersion"] = osVersion,
                ["osBuild"] = osBuild,
                ["osArchitecture"] = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit",
                ["computerName"] = Environment.MachineName ?? "unknown",
                ["domain"] = Environment.UserDomainName ?? "WORKGROUP",
                ["manufacturer"] = manufacturer,
                ["model"] = model,
                ["serialNumber"] = serialNumber
            };
        }

        private async Task<Dictionary<string, object>> CollectHardwareDataAsync()
        {
            string processor = "unknown";
            int processorCores = Environment.ProcessorCount;
            string processorSpeed = "unknown";
            double ramTotalGB = 0;
            double ramAvailableGB = 0;
            string gpuName = "unknown";
            double gpuMemoryGB = 0;
            double diskTotalGB = 0;
            double diskFreeGB = 0;
            string diskType = "unknown";

            await Task.Run(() =>
            {
                try
                {
                    // CPU
                    using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor"))
                    {
                        foreach (var obj in searcher.Get())
                        {
                            processor = obj["Name"]?.ToString()?.Trim() ?? "unknown";
                            processorSpeed = $"{obj["MaxClockSpeed"]} MHz";
                            break;
                        }
                    }

                    // RAM
                    using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem"))
                    {
                        foreach (var obj in searcher.Get())
                        {
                            if (ulong.TryParse(obj["TotalPhysicalMemory"]?.ToString(), out ulong totalBytes))
                            {
                                ramTotalGB = Math.Round(totalBytes / 1024.0 / 1024.0 / 1024.0, 1);
                            }
                            break;
                        }
                    }

                    // RAM disponible
                    using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem"))
                    {
                        foreach (var obj in searcher.Get())
                        {
                            if (ulong.TryParse(obj["FreePhysicalMemory"]?.ToString(), out ulong freeKB))
                            {
                                ramAvailableGB = Math.Round(freeKB / 1024.0 / 1024.0, 1);
                            }
                            break;
                        }
                    }

                    // GPU
                    using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController"))
                    {
                        foreach (var obj in searcher.Get())
                        {
                            gpuName = obj["Name"]?.ToString() ?? "unknown";
                            if (uint.TryParse(obj["AdapterRAM"]?.ToString(), out uint adapterRam))
                            {
                                gpuMemoryGB = Math.Round(adapterRam / 1024.0 / 1024.0 / 1024.0, 1);
                            }
                            break;
                        }
                    }

                    // Disk (lecteur C:)
                    var cDrive = DriveInfo.GetDrives().FirstOrDefault(d => d.Name.StartsWith("C"));
                    if (cDrive != null && cDrive.IsReady)
                    {
                        diskTotalGB = Math.Round(cDrive.TotalSize / 1024.0 / 1024.0 / 1024.0, 1);
                        diskFreeGB = Math.Round(cDrive.AvailableFreeSpace / 1024.0 / 1024.0 / 1024.0, 1);
                        diskType = cDrive.DriveType.ToString();
                    }
                }
                catch { }
            });

            return new Dictionary<string, object>
            {
                ["processor"] = processor,
                ["processorCores"] = processorCores,
                ["processorSpeed"] = processorSpeed,
                ["ramTotalGB"] = ramTotalGB,
                ["ramAvailableGB"] = ramAvailableGB,
                ["gpuName"] = gpuName,
                ["gpuMemoryGB"] = gpuMemoryGB,
                ["diskTotalGB"] = diskTotalGB,
                ["diskFreeGB"] = diskFreeGB,
                ["diskType"] = diskType
            };
        }

        private Dictionary<string, object> CollectNetworkData()
        {
            string ipLocal = "unknown";
            string macAddress = "unknown";
            string hostname = Environment.MachineName ?? "unknown";

            try
            {
                // IP locale
                var host = Dns.GetHostEntry(Dns.GetHostName());
                var ip = host.AddressList.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                if (ip != null) ipLocal = ip.ToString();

                // MAC Address
                var nic = NetworkInterface.GetAllNetworkInterfaces()
                    .FirstOrDefault(n => n.OperationalStatus == OperationalStatus.Up && 
                                         n.NetworkInterfaceType != NetworkInterfaceType.Loopback);
                if (nic != null)
                {
                    macAddress = string.Join("-", nic.GetPhysicalAddress().GetAddressBytes().Select(b => b.ToString("X2")));
                }
            }
            catch { }

            return new Dictionary<string, object>
            {
                ["ipAddressLocal"] = ipLocal,
                ["ipAddressPublic"] = "none", // Ne pas collecter IP publique pour la vie privee
                ["macAddress"] = macAddress,
                ["hostname"] = hostname,
                ["domainName"] = Environment.UserDomainName ?? "WORKGROUP",
                ["dnsServers"] = "none",
                ["networkSpeed"] = "none"
            };
        }

        private Dictionary<string, object> CollectSoftwareData()
        {
            string xeatVersion = GetApplicationVersion();
            string inventorVersion = "none";
            string inventorYear = "none";
            string vaultVersion = "none";
            string vaultServer = "none";
            string dotnetVersion = GetDotNetVersion();
            string officeVersion = "none";

            try
            {
                // Inventor
                inventorVersion = GetInventorVersion(out inventorYear);

                // Vault
                vaultVersion = GetVaultVersion(out vaultServer);

                // Office
                officeVersion = GetOfficeVersion();
            }
            catch { }

            return new Dictionary<string, object>
            {
                ["xeatVersion"] = xeatVersion,
                ["xeatInstalledAt"] = GetXeatInstallDate(),
                ["xeatLastUpdated"] = DateTime.UtcNow.ToString("yyyy-MM-dd"),
                ["inventorVersion"] = inventorVersion,
                ["inventorYear"] = inventorYear,
                ["vaultVersion"] = vaultVersion,
                ["vaultServer"] = vaultServer,
                ["dotnetVersion"] = dotnetVersion,
                ["officeVersion"] = officeVersion
            };
        }

        private Dictionary<string, object> CollectHeartbeatData()
        {
            // Performance actuelle
            float cpuUsage = 0;
            float ramUsage = 0;
            float diskUsage = 0;

            try
            {
                // CPU Usage (approximatif)
                using (var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total"))
                {
                    cpuCounter.NextValue();
                    Thread.Sleep(100);
                    cpuUsage = cpuCounter.NextValue();
                }
            }
            catch { }

            try
            {
                // RAM Usage
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        if (ulong.TryParse(obj["TotalVisibleMemorySize"]?.ToString(), out ulong total) &&
                            ulong.TryParse(obj["FreePhysicalMemory"]?.ToString(), out ulong free))
                        {
                            ramUsage = (float)((total - free) * 100.0 / total);
                        }
                        break;
                    }
                }
            }
            catch { }

            try
            {
                // Disk Usage (C:)
                var cDrive = DriveInfo.GetDrives().FirstOrDefault(d => d.Name.StartsWith("C"));
                if (cDrive != null && cDrive.IsReady)
                {
                    diskUsage = (float)((cDrive.TotalSize - cDrive.AvailableFreeSpace) * 100.0 / cDrive.TotalSize);
                }
            }
            catch { }

            return new Dictionary<string, object>
            {
                ["lastHeartbeat"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                ["intervalMs"] = 60000,
                ["missedHeartbeats"] = 0,
                ["status"] = "online",
                ["cpuUsage"] = Math.Round(cpuUsage, 1),
                ["ramUsage"] = Math.Round(ramUsage, 1),
                ["diskUsage"] = Math.Round(diskUsage, 1)
            };
        }

        #endregion

        #region Software Detection

        private string GetApplicationVersion()
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

        private string GetXeatInstallDate()
        {
            try
            {
                string exePath = Assembly.GetExecutingAssembly().Location;
                if (File.Exists(exePath))
                {
                    return File.GetCreationTime(exePath).ToString("yyyy-MM-dd");
                }
            }
            catch { }
            return "none";
        }

        private string GetInventorVersion(out string year)
        {
            year = "none";
            try
            {
                // Chercher dans le Registry
                string[] inventorKeys = {
                    @"SOFTWARE\Autodesk\Inventor\RegistryVersion26.0",
                    @"SOFTWARE\Autodesk\Inventor\RegistryVersion25.0",
                    @"SOFTWARE\Autodesk\Inventor\RegistryVersion24.0"
                };

                foreach (var keyPath in inventorKeys)
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(keyPath))
                    {
                        if (key != null)
                        {
                            string version = key.GetValue("ProductVersion")?.ToString();
                            if (!string.IsNullOrEmpty(version))
                            {
                                // Extraire l'annee du chemin
                                if (keyPath.Contains("26.0")) year = "2026";
                                else if (keyPath.Contains("25.0")) year = "2025";
                                else if (keyPath.Contains("24.0")) year = "2024";
                                return version;
                            }
                        }
                    }
                }

                // Alternative: chercher le processus Inventor
                var inventorProcess = Process.GetProcessesByName("Inventor").FirstOrDefault();
                if (inventorProcess != null)
                {
                    var fileInfo = FileVersionInfo.GetVersionInfo(inventorProcess.MainModule.FileName);
                    year = $"20{fileInfo.FileMajorPart - 4}"; // Inventor 26 = 2026
                    return fileInfo.FileVersion;
                }
            }
            catch { }
            return "none";
        }

        private string GetVaultVersion(out string server)
        {
            server = "none";
            try
            {
                // Chercher dans le Registry
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk\VaultPro\CurrentVersion"))
                {
                    if (key != null)
                    {
                        string version = key.GetValue("Version")?.ToString();
                        server = key.GetValue("Server")?.ToString() ?? "none";
                        if (!string.IsNullOrEmpty(version)) return version;
                    }
                }

                // Alternative: fichier de config Vault
                string vaultConfigPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Autodesk", "VaultCommon", "Servers"
                );
                if (Directory.Exists(vaultConfigPath))
                {
                    var serverDirs = Directory.GetDirectories(vaultConfigPath);
                    if (serverDirs.Length > 0)
                    {
                        server = Path.GetFileName(serverDirs[0]);
                    }
                }
            }
            catch { }
            return "none";
        }

        private string GetDotNetVersion()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                {
                    if (key != null)
                    {
                        int release = (int)(key.GetValue("Release") ?? 0);
                        // Mapper le release number vers la version
                        if (release >= 533320) return "4.8.1";
                        if (release >= 528040) return "4.8";
                        if (release >= 461808) return "4.7.2";
                        if (release >= 461308) return "4.7.1";
                        if (release >= 460798) return "4.7";
                        if (release >= 394802) return "4.6.2";
                        return $"4.x (release {release})";
                    }
                }
            }
            catch { }
            return Environment.Version.ToString();
        }

        private string GetOfficeVersion()
        {
            try
            {
                string[] officeKeys = {
                    @"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
                    @"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot"
                };

                foreach (var keyPath in officeKeys)
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(keyPath))
                    {
                        if (key != null)
                        {
                            string path = key.GetValue("Path")?.ToString();
                            if (!string.IsNullOrEmpty(path))
                            {
                                if (keyPath.Contains("16.0")) return "Office 365/2016/2019/2021";
                                if (keyPath.Contains("15.0")) return "Office 2013";
                            }
                        }
                    }
                }
            }
            catch { }
            return "none";
        }

        private string DetectSite()
        {
            try
            {
                // Detection basee sur le domaine ou l'IP
                string domain = Environment.UserDomainName?.ToUpperInvariant() ?? "";
                
                if (domain.Contains("XNRGY") || domain.Contains("QC"))
                {
                    // Verifier l'IP pour distinguer les sites Quebec
                    var host = Dns.GetHostEntry(Dns.GetHostName());
                    var ip = host.AddressList.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                    if (ip != null)
                    {
                        string ipStr = ip.ToString();
                        // Ajuster selon les plages IP reelles
                        if (ipStr.StartsWith("10.10.")) return "saintHubertQC";
                        if (ipStr.StartsWith("10.20.")) return "lavalQC";
                        if (ipStr.StartsWith("10.30.")) return "arizonaUS";
                    }
                    return "saintHubertQC"; // Par defaut Quebec
                }
                
                if (domain.Contains("US") || domain.Contains("AZ"))
                {
                    return "arizonaUS";
                }
            }
            catch { }
            
            return "saintHubertQC"; // Site par defaut
        }

        #endregion

        #region Heartbeat

        /// <summary>
        /// Envoie un heartbeat a Firebase
        /// </summary>
        public async Task SendHeartbeatAsync()
        {
            try
            {
                var heartbeatData = CollectHeartbeatData();
                
                // Mettre a jour aussi le status online et lastSeen
                var statusUpdate = new Dictionary<string, object>
                {
                    ["online"] = true,
                    ["lastSeen"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    ["currentUser"] = Environment.UserName ?? "unknown"
                };

                // Envoyer en parallele
                var heartbeatTask = UpdateFirebaseAsync($"/devices/{_deviceId}/heartbeat", heartbeatData);
                var statusTask = UpdateFirebaseAsync($"/devices/{_deviceId}/status", statusUpdate);

                await Task.WhenAll(heartbeatTask, statusTask);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur heartbeat: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }

        /// <summary>
        /// Marque le device comme offline
        /// </summary>
        public async Task SetOfflineAsync()
        {
            try
            {
                var offlineData = new Dictionary<string, object>
                {
                    ["online"] = false,
                    ["lastSeen"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
                };

                var heartbeatData = new Dictionary<string, object>
                {
                    ["status"] = "offline",
                    ["lastHeartbeat"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
                };

                await Task.WhenAll(
                    UpdateFirebaseAsync($"/devices/{_deviceId}/status", offlineData),
                    UpdateFirebaseAsync($"/devices/{_deviceId}/heartbeat", heartbeatData)
                );

                Logger.Log($"[+] Device marque offline: {_deviceId}");
            }
            catch { }
        }

        #endregion

        #region Firebase Communication

        /// <summary>
        /// Met a jour des donnees dans Firebase (PATCH)
        /// </summary>
        private async Task UpdateFirebaseAsync(string path, object data)
        {
            try
            {
                string url = $"{FIREBASE_DATABASE_URL}{path}.json";
                string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = false });
                
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var request = new HttpRequestMessage(new HttpMethod("PATCH"), url) { Content = content };
                
                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur Firebase PATCH {path}: {ex.Message}", Logger.LogLevel.DEBUG);
                throw;
            }
        }

        /// <summary>
        /// Ecrit des donnees dans Firebase (PUT - remplace)
        /// </summary>
        private async Task WriteFirebaseAsync(string path, object data)
        {
            try
            {
                string url = $"{FIREBASE_DATABASE_URL}{path}.json";
                string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = false });
                
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PutAsync(url, content);
                response.EnsureSuccessStatusCode();
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur Firebase PUT {path}: {ex.Message}", Logger.LogLevel.DEBUG);
                throw;
            }
        }

        #endregion

        #region Telemetry Events

        /// <summary>
        /// Enregistre un evenement de telemetrie
        /// </summary>
        public async Task LogEventAsync(string eventType, Dictionary<string, object> data = null)
        {
            try
            {
                string eventId = $"{_deviceId}_{DateTime.UtcNow:yyyyMMddHHmmss}_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
                
                var eventData = new Dictionary<string, object>
                {
                    ["eventId"] = eventId,
                    ["eventType"] = eventType,
                    ["timestamp"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    ["userId"] = _userId,
                    ["deviceId"] = _deviceId,
                    ["data"] = data != null ? JsonSerializer.Serialize(data) : "none"
                };

                await WriteFirebaseAsync($"/telemetryEvents/{eventId}", eventData);
            }
            catch { }
        }

        /// <summary>
        /// Enregistre une erreur
        /// </summary>
        public async Task LogErrorAsync(Exception ex, string context = null)
        {
            try
            {
                string errorId = $"err_{DateTime.UtcNow:yyyyMMddHHmmss}_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
                
                var errorData = new Dictionary<string, object>
                {
                    ["errorId"] = errorId,
                    ["errorType"] = ex.GetType().Name,
                    ["message"] = ex.Message,
                    ["stackTrace"] = ex.StackTrace?.Substring(0, Math.Min(ex.StackTrace?.Length ?? 0, 1000)) ?? "none",
                    ["timestamp"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    ["userId"] = _userId,
                    ["deviceId"] = _deviceId,
                    ["appVersion"] = GetApplicationVersion(),
                    ["context"] = context ?? "none"
                };

                await WriteFirebaseAsync($"/errorReports/{errorId}", errorData);
            }
            catch { }
        }

        /// <summary>
        /// Incremente un compteur d'utilisation de module
        /// </summary>
        public async Task IncrementModuleUsageAsync(string moduleName)
        {
            try
            {
                var updateData = new Dictionary<string, object>
                {
                    ["lastUsed"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
                };

                await UpdateFirebaseAsync($"/statistics/byModule/{moduleName}", updateData);
            }
            catch { }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!_disposed)
            {
                StopHeartbeat();
                
                // Marquer offline de maniere synchrone au shutdown
                try
                {
                    SetOfflineAsync().GetAwaiter().GetResult();
                }
                catch { }
                
                _disposed = true;
            }
        }

        #endregion
    }
}
