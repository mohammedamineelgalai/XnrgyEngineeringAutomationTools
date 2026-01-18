using System;
using System.Collections.Generic;
using System.Management;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace XnrgyInstaller
{
    /// <summary>
    /// Service Firebase pour enregistrer les installations/desinstallations
    /// Utilise l'API REST Firebase (pas de SDK requis)
    /// </summary>
    public class FirebaseService
    {
        #region Constants

        private const string FIREBASE_DATABASE_URL = "https://xeat-remote-control-default-rtdb.firebaseio.com";
        private const string APP_VERSION = "1.0.0";

        #endregion

        #region Private Fields

        private static readonly HttpClient _httpClient = new HttpClient();
        private readonly string _deviceId;
        private readonly string _machineName;
        private readonly string _userName;
        private readonly string _domain;

        #endregion

        #region Constructor

        public FirebaseService()
        {
            _machineName = Environment.MachineName;
            _userName = Environment.UserName;
            _domain = Environment.UserDomainName;
            _deviceId = $"{_machineName}_{_userName}".Replace(".", "_").Replace(" ", "_");
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Enregistre une nouvelle installation dans Firebase
        /// </summary>
        public async Task<bool> RegisterInstallationAsync()
        {
            try
            {
                var timestamp = DateTime.UtcNow.ToString("o");
                var systemInfo = GetSystemInfo();

                // 1. Enregistrer/Mettre a jour le device
                var deviceData = new Dictionary<string, object>
                {
                    ["registration"] = new Dictionary<string, object>
                    {
                        ["deviceId"] = _deviceId,
                        ["machineName"] = _machineName,
                        ["registeredAt"] = timestamp,
                        ["registeredBy"] = _userName,
                        ["approved"] = true,
                        ["approvedAt"] = timestamp,
                        ["approvedBy"] = "auto"
                    },
                    ["organization"] = new Dictionary<string, object>
                    {
                        ["assignedTo"] = _userName,
                        ["department"] = "engineering",
                        ["site"] = "montrealQC",
                        ["assetTag"] = "none",
                        ["location"] = "none",
                        ["purchaseDate"] = "none",
                        ["warrantyExpires"] = "none"
                    },
                    ["status"] = new Dictionary<string, object>
                    {
                        ["online"] = false,
                        ["enabled"] = true,
                        ["blocked"] = false,
                        ["blockedAt"] = "none",
                        ["blockedBy"] = "none",
                        ["blockReason"] = "none",
                        ["currentUser"] = _userName,
                        ["currentUserId"] = $"{_domain}_{_userName}",
                        ["lastSeen"] = timestamp
                    },
                    ["system"] = systemInfo,
                    ["software"] = new Dictionary<string, object>
                    {
                        ["xeatVersion"] = APP_VERSION,
                        ["xeatInstalledAt"] = timestamp,
                        ["xeatLastUpdated"] = timestamp,
                        ["dotnetVersion"] = Environment.Version.ToString(),
                        ["inventorVersion"] = "none",
                        ["inventorYear"] = "none",
                        ["vaultVersion"] = "none",
                        ["vaultServer"] = "none",
                        ["officeVersion"] = "none"
                    },
                    ["heartbeat"] = new Dictionary<string, object>
                    {
                        ["status"] = "offline",
                        ["lastHeartbeat"] = "none",
                        ["missedHeartbeats"] = 0,
                        ["intervalMs"] = 60000,
                        ["cpuUsage"] = 0,
                        ["ramUsage"] = 0,
                        ["diskUsage"] = 0
                    },
                    ["hardware"] = GetHardwareInfo(),
                    ["network"] = GetNetworkInfo(),
                    ["security"] = new Dictionary<string, object>
                    {
                        ["trustedDevice"] = false,
                        ["requiresApproval"] = false,
                        ["failedLoginAttempts"] = 0,
                        ["lastLoginAttempt"] = "none"
                    },
                    ["usage"] = new Dictionary<string, object>
                    {
                        ["totalSessions"] = 0,
                        ["totalUploads"] = 0,
                        ["totalModulesCreated"] = 0,
                        ["lastUpload"] = "none",
                        ["lastModuleCreated"] = "none"
                    },
                    ["metadata"] = new Dictionary<string, object>
                    {
                        ["createdAt"] = timestamp,
                        ["updatedAt"] = timestamp,
                        ["notes"] = "Installed via XEAT Setup",
                        ["tags"] = "auto"
                    },
                    // NOUVEAU: Historique d'installation pour affichage dans Postes
                    ["installationHistory"] = new Dictionary<string, object>
                    {
                        [$"install_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}"] = new Dictionary<string, object>
                        {
                            ["action"] = "install",
                            ["version"] = APP_VERSION,
                            ["timestamp"] = timestamp,
                            ["installedBy"] = _userName,
                            ["machineName"] = _machineName,
                            ["installPath"] = GetInstallPath(),
                            ["osVersion"] = Environment.OSVersion.ToString(),
                            ["success"] = true,
                            ["notes"] = "Initial installation"
                        }
                    }
                };

                await PatchFirebaseAsync($"devices/{_deviceId}", deviceData);

                // 2. Enregistrer l'evenement d'installation dans auditLog
                var logId = $"log_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";
                var auditEntry = new Dictionary<string, object>
                {
                    ["id"] = logId,
                    ["action"] = "app_installed",
                    ["category"] = "installation",
                    ["userId"] = _userName,
                    ["userName"] = _userName,
                    ["deviceId"] = _deviceId,
                    ["timestamp"] = timestamp,
                    ["details"] = $"XEAT v{APP_VERSION} installed on {_machineName}",
                    ["success"] = true,
                    ["ipAddress"] = "local",
                    ["oldValue"] = "none",
                    ["newValue"] = APP_VERSION,
                    ["errorMessage"] = "none"
                };

                await PutFirebaseAsync($"auditLog/{logId}", auditEntry);

                // 3. Mettre a jour les statistiques
                await IncrementStatisticAsync("statistics/devices/totalRegistered");
                await IncrementSiteStatisticAsync("montrealQC", "devices");

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Firebase registration error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Enregistre une desinstallation dans Firebase
        /// </summary>
        public async Task<bool> RegisterUninstallationAsync()
        {
            try
            {
                var timestamp = DateTime.UtcNow.ToString("o");

                // 1. Marquer le device comme desinstalle (ne pas supprimer pour historique)
                var statusUpdate = new Dictionary<string, object>
                {
                    ["status/online"] = false,
                    ["status/enabled"] = false,
                    ["software/xeatVersion"] = "uninstalled",
                    ["software/xeatLastUpdated"] = timestamp,
                    ["metadata/updatedAt"] = timestamp,
                    ["metadata/notes"] = $"Uninstalled on {timestamp}"
                };

                await PatchFirebaseAsync($"devices/{_deviceId}", statusUpdate);

                // 1b. Ajouter l'evenement dans installationHistory
                var uninstallHistoryEntry = new Dictionary<string, object>
                {
                    ["action"] = "uninstall",
                    ["version"] = APP_VERSION,
                    ["timestamp"] = timestamp,
                    ["uninstalledBy"] = _userName,
                    ["machineName"] = _machineName,
                    ["success"] = true,
                    ["notes"] = "User initiated uninstall"
                };
                await PutFirebaseAsync($"devices/{_deviceId}/installationHistory/uninstall_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}", uninstallHistoryEntry);

                // 2. Enregistrer l'evenement de desinstallation
                var logId = $"log_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";
                var auditEntry = new Dictionary<string, object>
                {
                    ["id"] = logId,
                    ["action"] = "app_uninstalled",
                    ["category"] = "installation",
                    ["userId"] = _userName,
                    ["userName"] = _userName,
                    ["deviceId"] = _deviceId,
                    ["timestamp"] = timestamp,
                    ["details"] = $"XEAT uninstalled from {_machineName}",
                    ["success"] = true,
                    ["ipAddress"] = "local",
                    ["oldValue"] = APP_VERSION,
                    ["newValue"] = "none",
                    ["errorMessage"] = "none"
                };

                await PutFirebaseAsync($"auditLog/{logId}", auditEntry);

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Firebase unregistration error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Envoie un rapport d'erreur a Firebase (pour Audit Logs)
        /// </summary>
        public static async Task<bool> ReportErrorAsync(string errorType, string message, string stackTrace = null)
        {
            try
            {
                var timestamp = DateTime.UtcNow.ToString("o");
                var machineName = Environment.MachineName;
                var userName = Environment.UserName;
                var deviceId = $"{machineName}_{userName}".Replace(".", "_").Replace(" ", "_");

                var logId = $"log_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";
                var auditEntry = new Dictionary<string, object>
                {
                    ["id"] = logId,
                    ["action"] = "error_reported",
                    ["category"] = "error",
                    ["userId"] = userName,
                    ["userName"] = userName,
                    ["deviceId"] = deviceId,
                    ["timestamp"] = timestamp,
                    ["details"] = $"[{errorType}] {message}",
                    ["success"] = false,
                    ["ipAddress"] = "local",
                    ["oldValue"] = "none",
                    ["newValue"] = "none",
                    ["errorMessage"] = stackTrace ?? message
                };

                var json = DictionaryToJson(auditEntry);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/auditLog/{logId}.json", content);

                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Private Methods - Firebase API

        private async Task PutFirebaseAsync(string path, Dictionary<string, object> data)
        {
            var json = DictionaryToJson(data);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/{path}.json", content);
        }

        private async Task PatchFirebaseAsync(string path, Dictionary<string, object> data)
        {
            var json = DictionaryToJson(data);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var request = new HttpRequestMessage(new HttpMethod("PATCH"), $"{FIREBASE_DATABASE_URL}/{path}.json")
            {
                Content = content
            };
            await _httpClient.SendAsync(request);
        }

        private async Task IncrementStatisticAsync(string path)
        {
            try
            {
                // Lire la valeur actuelle
                var response = await _httpClient.GetStringAsync($"{FIREBASE_DATABASE_URL}/{path}.json");
                int currentValue = 0;
                if (int.TryParse(response.Trim('"'), out int val))
                    currentValue = val;

                // Incrementer
                var content = new StringContent((currentValue + 1).ToString(), Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/{path}.json", content);
            }
            catch { }
        }

        private async Task IncrementSiteStatisticAsync(string site, string stat)
        {
            await IncrementStatisticAsync($"statistics/bySite/{site}/{stat}");
        }

        #endregion

        #region Private Methods - System Info

        private Dictionary<string, object> GetSystemInfo()
        {
            var info = new Dictionary<string, object>
            {
                ["computerName"] = _machineName,
                ["domain"] = _domain,
                ["osName"] = GetOSName(),
                ["osVersion"] = Environment.OSVersion.Version.ToString(),
                ["osBuild"] = GetOSBuild(),
                ["osArchitecture"] = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit",
                ["manufacturer"] = GetWmiValue("Win32_ComputerSystem", "Manufacturer"),
                ["model"] = GetWmiValue("Win32_ComputerSystem", "Model"),
                ["serialNumber"] = GetWmiValue("Win32_BIOS", "SerialNumber")
            };
            return info;
        }

        private Dictionary<string, object> GetHardwareInfo()
        {
            return new Dictionary<string, object>
            {
                ["processor"] = GetWmiValue("Win32_Processor", "Name"),
                ["processorCores"] = Environment.ProcessorCount,
                ["processorSpeed"] = "N/A",
                ["ramTotalGB"] = Math.Round(GetTotalMemoryGB(), 1),
                ["ramAvailableGB"] = 0,
                ["diskTotalGB"] = 0,
                ["diskFreeGB"] = 0,
                ["diskType"] = "Unknown",
                ["gpuName"] = GetWmiValue("Win32_VideoController", "Name"),
                ["gpuMemoryGB"] = 0
            };
        }

        private Dictionary<string, object> GetNetworkInfo()
        {
            return new Dictionary<string, object>
            {
                ["hostname"] = _machineName,
                ["domainName"] = _domain,
                ["ipAddressLocal"] = GetLocalIPAddress(),
                ["ipAddressPublic"] = "none",
                ["macAddress"] = GetMacAddress(),
                ["networkSpeed"] = "Unknown",
                ["dnsServers"] = "none"
            };
        }

        private string GetOSName()
        {
            try
            {
                var name = GetWmiValue("Win32_OperatingSystem", "Caption");
                var version = GetWmiValue("Win32_OperatingSystem", "Version");
                return $"{name} ({version})";
            }
            catch
            {
                return Environment.OSVersion.VersionString;
            }
        }

        private string GetOSBuild()
        {
            try
            {
                return GetWmiValue("Win32_OperatingSystem", "BuildNumber");
            }
            catch
            {
                return Environment.OSVersion.Version.Build.ToString();
            }
        }

        private string GetWmiValue(string wmiClass, string property)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher($"SELECT {property} FROM {wmiClass}"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var value = obj[property]?.ToString();
                        if (!string.IsNullOrEmpty(value))
                            return value.Trim();
                    }
                }
            }
            catch { }
            return "Unknown";
        }

        private double GetTotalMemoryGB()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var bytes = Convert.ToUInt64(obj["TotalPhysicalMemory"]);
                        return bytes / (1024.0 * 1024.0 * 1024.0);
                    }
                }
            }
            catch { }
            return 0;
        }

        private string GetLocalIPAddress()
        {
            try
            {
                var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        return ip.ToString();
                }
            }
            catch { }
            return "Unknown";
        }

        private string GetMacAddress()
        {
            try
            {
                foreach (var nic in System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (nic.OperationalStatus == System.Net.NetworkInformation.OperationalStatus.Up &&
                        nic.NetworkInterfaceType != System.Net.NetworkInformation.NetworkInterfaceType.Loopback)
                    {
                        return nic.GetPhysicalAddress().ToString();
                    }
                }
            }
            catch { }
            return "Unknown";
        }

        private string GetInstallPath()
        {
            try
            {
                string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                return System.IO.Path.Combine(programFiles, "XNRGY Climate Systems ULC", "XNRGY Engineering Automation Tools");
            }
            catch
            {
                return "C:\\Program Files\\XNRGY Climate Systems ULC\\XNRGY Engineering Automation Tools";
            }
        }

        #endregion

        #region Private Methods - JSON Helpers

        private static string DictionaryToJson(Dictionary<string, object> dict)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            bool first = true;
            foreach (var kvp in dict)
            {
                if (!first) sb.Append(",");
                first = false;
                sb.Append($"\"{kvp.Key}\":");
                sb.Append(ObjectToJson(kvp.Value));
            }
            sb.Append("}");
            return sb.ToString();
        }

        private static string ObjectToJson(object obj)
        {
            if (obj == null) return "null";
            if (obj is string s) return $"\"{EscapeJson(s)}\"";
            if (obj is bool b) return b.ToString().ToLower();
            if (obj is int || obj is long || obj is double || obj is float) return obj.ToString();
            if (obj is Dictionary<string, object> dict) return DictionaryToJson(dict);
            return $"\"{EscapeJson(obj.ToString())}\"";
        }

        private static string EscapeJson(string s)
        {
            return s.Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\n", "\\n")
                    .Replace("\r", "\\r")
                    .Replace("\t", "\\t");
        }

        #endregion
    }
}
