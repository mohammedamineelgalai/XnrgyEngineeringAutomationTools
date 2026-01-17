using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

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
                var deviceInfo = new
                {
                    machineName = Environment.MachineName,
                    userName = Environment.UserName,
                    osVersion = Environment.OSVersion.ToString(),
                    appVersion = GetAppVersion(),
                    startTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    lastHeartbeat = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    status = "online",
                    site = GetSiteFromMachineName()
                };

                string json = JsonSerializer.Serialize(deviceInfo);
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

                // Mettre le statut a "offline" via PUT sur un sous-noeud
                var offlineStatus = new
                {
                    lastHeartbeat = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    status = "offline",
                    disconnectTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
                };

                string json = JsonSerializer.Serialize(offlineStatus);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                string url = $"{FIREBASE_DATABASE_URL}/devices/{_deviceId}/heartbeat.json";
                await _httpClient.PutAsync(url, content);

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

        /// <summary>
        /// Libere les ressources
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            
            _heartbeatTimer?.Stop();
            _heartbeatTimer?.Dispose();
            
            // Essayer de desenregistrer de facon synchrone
            try
            {
                Task.Run(async () => await UnregisterDeviceAsync()).Wait(3000);
            }
            catch { }
            
            _disposed = true;
        }
    }
}
