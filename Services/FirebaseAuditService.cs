using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service Firebase pour l'application XEAT
    /// - Envoie les erreurs/bugs a Firebase Audit Logs
    /// - Enregistre les sessions utilisateur
    /// - Met a jour le heartbeat device
    /// PERFORMANCE: Seules les erreurs sont envoyees, pas les logs INFO/DEBUG
    /// </summary>
    public class FirebaseAuditService
    {
        #region Constants

        private const string FIREBASE_DATABASE_URL = "https://xeat-remote-control-default-rtdb.firebaseio.com";
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };

        #endregion

        #region Private Fields

        private static FirebaseAuditService _instance;
        private static readonly object _lock = new object();

        private readonly string _deviceId;
        private readonly string _machineName;
        private readonly string _userName;
        private readonly string _appVersion;
        private bool _isInitialized;

        // Queue pour batch les erreurs (eviter trop de requetes)
        private readonly Queue<AuditLogEntry> _errorQueue = new Queue<AuditLogEntry>();
        private readonly object _queueLock = new object();
        private DateTime _lastFlush = DateTime.MinValue;
        private const int FLUSH_INTERVAL_SECONDS = 30;
        private const int MAX_QUEUE_SIZE = 10;

        #endregion

        #region Singleton

        public static FirebaseAuditService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new FirebaseAuditService();
                        }
                    }
                }
                return _instance;
            }
        }

        private FirebaseAuditService()
        {
            _machineName = Environment.MachineName;
            _userName = Environment.UserName;
            _deviceId = $"{_machineName}_{_userName}".Replace(".", "_").Replace(" ", "_");
            _appVersion = "1.0.0"; // TODO: Lire depuis assembly
            _isInitialized = false;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Initialise le service et enregistre le demarrage de session
        /// </summary>
        public async Task InitializeAsync()
        {
            if (_isInitialized) return;

            try
            {
                await RegisterSessionStartAsync();
                _isInitialized = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Firebase init warning: {ex.Message}");
                // Continue meme si Firebase echoue
            }
        }

        /// <summary>
        /// Enregistre une erreur dans Firebase Audit Logs
        /// ASYNC - N'attend pas la reponse pour ne pas bloquer l'UI
        /// </summary>
        public void LogError(string errorType, string message, string stackTrace = null, string context = null)
        {
            var entry = new AuditLogEntry
            {
                Action = "error_reported",
                Category = "error",
                ErrorType = errorType,
                Message = message,
                StackTrace = stackTrace,
                Context = context,
                Timestamp = DateTime.UtcNow
            };

            lock (_queueLock)
            {
                _errorQueue.Enqueue(entry);

                // Flush si queue pleine ou intervalle depasse
                if (_errorQueue.Count >= MAX_QUEUE_SIZE ||
                    (DateTime.Now - _lastFlush).TotalSeconds >= FLUSH_INTERVAL_SECONDS)
                {
                    _ = FlushErrorQueueAsync();
                }
            }
        }

        /// <summary>
        /// Enregistre une exception dans Firebase
        /// </summary>
        public void LogException(Exception ex, string context = null)
        {
            LogError(
                errorType: ex.GetType().Name,
                message: ex.Message,
                stackTrace: ex.StackTrace,
                context: context
            );
        }

        /// <summary>
        /// Enregistre une action utilisateur (module utilise, etc.)
        /// </summary>
        public async Task LogUserActionAsync(string action, string details = null)
        {
            try
            {
                var logId = $"log_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";
                var timestamp = DateTime.UtcNow.ToString("o");

                var auditEntry = new Dictionary<string, object>
                {
                    ["id"] = logId,
                    ["action"] = action,
                    ["category"] = "user_action",
                    ["userId"] = _userName,
                    ["userName"] = _userName,
                    ["deviceId"] = _deviceId,
                    ["timestamp"] = timestamp,
                    ["details"] = details ?? "none",
                    ["success"] = true,
                    ["ipAddress"] = "local",
                    ["oldValue"] = "none",
                    ["newValue"] = "none",
                    ["errorMessage"] = "none"
                };

                await PutFirebaseAsync($"auditLog/{logId}", auditEntry);
            }
            catch
            {
                // Silencieux - ne pas bloquer l'app
            }
        }

        /// <summary>
        /// Met a jour le heartbeat du device (appeler periodiquement)
        /// </summary>
        public async Task UpdateHeartbeatAsync()
        {
            try
            {
                var timestamp = DateTime.UtcNow.ToString("o");
                var heartbeatData = new Dictionary<string, object>
                {
                    ["status"] = "online",
                    ["lastHeartbeat"] = timestamp,
                    ["missedHeartbeats"] = 0,
                    ["cpuUsage"] = GetCpuUsage(),
                    ["ramUsage"] = GetRamUsage(),
                    ["diskUsage"] = 0
                };

                await PatchFirebaseAsync($"devices/{_deviceId}/heartbeat", heartbeatData);

                // Mettre a jour le status online
                await PatchFirebaseAsync($"devices/{_deviceId}/status", new Dictionary<string, object>
                {
                    ["online"] = true,
                    ["lastSeen"] = timestamp
                });
            }
            catch
            {
                // Silencieux
            }
        }

        /// <summary>
        /// Enregistre la fin de session
        /// </summary>
        public async Task RegisterSessionEndAsync()
        {
            try
            {
                // Flush les erreurs en attente
                await FlushErrorQueueAsync();

                var timestamp = DateTime.UtcNow.ToString("o");

                // Marquer le device comme offline
                await PatchFirebaseAsync($"devices/{_deviceId}/status", new Dictionary<string, object>
                {
                    ["online"] = false,
                    ["lastSeen"] = timestamp
                });

                await PatchFirebaseAsync($"devices/{_deviceId}/heartbeat", new Dictionary<string, object>
                {
                    ["status"] = "offline"
                });

                // Log de fin de session
                var logId = $"log_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";
                await PutFirebaseAsync($"auditLog/{logId}", new Dictionary<string, object>
                {
                    ["id"] = logId,
                    ["action"] = "session_end",
                    ["category"] = "session",
                    ["userId"] = _userName,
                    ["userName"] = _userName,
                    ["deviceId"] = _deviceId,
                    ["timestamp"] = timestamp,
                    ["details"] = "Application closed",
                    ["success"] = true,
                    ["ipAddress"] = "local",
                    ["oldValue"] = "none",
                    ["newValue"] = "none",
                    ["errorMessage"] = "none"
                });
            }
            catch
            {
                // Silencieux
            }
        }

        #endregion

        #region Private Methods

        private async Task RegisterSessionStartAsync()
        {
            var timestamp = DateTime.UtcNow.ToString("o");

            // Mettre a jour le device comme online
            await PatchFirebaseAsync($"devices/{_deviceId}/status", new Dictionary<string, object>
            {
                ["online"] = true,
                ["lastSeen"] = timestamp,
                ["currentUser"] = _userName
            });

            await PatchFirebaseAsync($"devices/{_deviceId}/heartbeat", new Dictionary<string, object>
            {
                ["status"] = "online",
                ["lastHeartbeat"] = timestamp
            });

            await PatchFirebaseAsync($"devices/{_deviceId}/software", new Dictionary<string, object>
            {
                ["xeatVersion"] = _appVersion,
                ["xeatLastUpdated"] = timestamp
            });

            // Log de debut de session
            var logId = $"log_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";
            await PutFirebaseAsync($"auditLog/{logId}", new Dictionary<string, object>
            {
                ["id"] = logId,
                ["action"] = "session_start",
                ["category"] = "session",
                ["userId"] = _userName,
                ["userName"] = _userName,
                ["deviceId"] = _deviceId,
                ["timestamp"] = timestamp,
                ["details"] = $"XEAT v{_appVersion} started",
                ["success"] = true,
                ["ipAddress"] = "local",
                ["oldValue"] = "none",
                ["newValue"] = "none",
                ["errorMessage"] = "none"
            });

            // Incrementer les statistiques
            await IncrementStatAsync("statistics/global/totalSessions");
        }

        private async Task FlushErrorQueueAsync()
        {
            List<AuditLogEntry> entriesToSend;

            lock (_queueLock)
            {
                if (_errorQueue.Count == 0) return;

                entriesToSend = new List<AuditLogEntry>(_errorQueue);
                _errorQueue.Clear();
                _lastFlush = DateTime.Now;
            }

            foreach (var entry in entriesToSend)
            {
                try
                {
                    var logId = $"log_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
                    var auditEntry = new Dictionary<string, object>
                    {
                        ["id"] = logId,
                        ["action"] = entry.Action,
                        ["category"] = entry.Category,
                        ["userId"] = _userName,
                        ["userName"] = _userName,
                        ["deviceId"] = _deviceId,
                        ["timestamp"] = entry.Timestamp.ToString("o"),
                        ["details"] = $"[{entry.ErrorType}] {entry.Message}" + (entry.Context != null ? $" (Context: {entry.Context})" : ""),
                        ["success"] = false,
                        ["ipAddress"] = "local",
                        ["oldValue"] = "none",
                        ["newValue"] = "none",
                        ["errorMessage"] = entry.StackTrace ?? entry.Message
                    };

                    await PutFirebaseAsync($"auditLog/{logId}", auditEntry);

                    // Incrementer le compteur d'erreurs
                    await IncrementStatAsync("statistics/global/totalErrors");
                }
                catch
                {
                    // Silencieux
                }
            }
        }

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

        private async Task IncrementStatAsync(string path)
        {
            try
            {
                var response = await _httpClient.GetStringAsync($"{FIREBASE_DATABASE_URL}/{path}.json");
                int currentValue = 0;
                if (int.TryParse(response.Trim('"'), out int val))
                    currentValue = val;

                var content = new StringContent((currentValue + 1).ToString(), Encoding.UTF8, "application/json");
                await _httpClient.PutAsync($"{FIREBASE_DATABASE_URL}/{path}.json", content);
            }
            catch { }
        }

        private int GetCpuUsage()
        {
            // Approximation simple
            return 0;
        }

        private int GetRamUsage()
        {
            try
            {
                var proc = System.Diagnostics.Process.GetCurrentProcess();
                var usedMB = proc.WorkingSet64 / (1024 * 1024);
                // Approximation du pourcentage
                return (int)(usedMB / 10); // Rough estimate
            }
            catch
            {
                return 0;
            }
        }

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
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\n", "\\n")
                    .Replace("\r", "\\r")
                    .Replace("\t", "\\t");
        }

        #endregion

        #region Inner Classes

        private class AuditLogEntry
        {
            public string Action { get; set; }
            public string Category { get; set; }
            public string ErrorType { get; set; }
            public string Message { get; set; }
            public string StackTrace { get; set; }
            public string Context { get; set; }
            public DateTime Timestamp { get; set; }
        }

        #endregion
    }
}
