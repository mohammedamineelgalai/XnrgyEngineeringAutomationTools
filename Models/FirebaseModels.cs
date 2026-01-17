using System;

namespace XnrgyEngineeringAutomationTools.Models
{
    /// <summary>
    /// Modeles pour la configuration Firebase Realtime Database
    /// </summary>
    public class FirebaseConfig
    {
        public AppConfig AppConfig { get; set; }
        public GlobalCommands Commands { get; set; }
        public VersionInfo VersionInfo { get; set; }
    }

    public class AppConfig
    {
        public string CurrentVersion { get; set; }
        public string MinVersion { get; set; }
        public bool MaintenanceMode { get; set; }
        public string MaintenanceMessage { get; set; }
        public bool ForceUpdate { get; set; }
        public string UpdateUrl { get; set; }
        public long LastUpdated { get; set; }
    }

    public class GlobalCommands
    {
        public GlobalCommandSettings Global { get; set; }
    }

    public class GlobalCommandSettings
    {
        public bool KillSwitch { get; set; }
        public string KillSwitchMessage { get; set; }
        public bool ForceUpdate { get; set; }
    }

    public class VersionInfo
    {
        public LatestVersion Latest { get; set; }
    }

    public class LatestVersion
    {
        public string Version { get; set; }
        public string DownloadUrl { get; set; }
        public string ReleaseDate { get; set; }
        public string Changelog { get; set; }
    }

    /// <summary>
    /// Resultat de la verification Firebase
    /// </summary>
    public class FirebaseCheckResult
    {
        public bool Success { get; set; }
        public bool KillSwitchActive { get; set; }
        public string KillSwitchMessage { get; set; }
        public bool MaintenanceMode { get; set; }
        public string MaintenanceMessage { get; set; }
        public bool UpdateAvailable { get; set; }
        public bool ForceUpdate { get; set; }
        public string LatestVersion { get; set; }
        public string DownloadUrl { get; set; }
        public string Changelog { get; set; }
        public string ErrorMessage { get; set; }
        
        // Controle par utilisateur
        public bool UserDisabled { get; set; }
        public string UserDisabledMessage { get; set; }
        
        // Messages personnalises
        public bool HasBroadcastMessage { get; set; }
        public string BroadcastTitle { get; set; }
        public string BroadcastMessage { get; set; }
        public string BroadcastType { get; set; } // "info", "warning", "error"

        public static FirebaseCheckResult CreateError(string message)
        {
            return new FirebaseCheckResult
            {
                Success = false,
                ErrorMessage = message
            };
        }

        public static FirebaseCheckResult CreateSuccess()
        {
            return new FirebaseCheckResult
            {
                Success = true
            };
        }
    }

    /// <summary>
    /// Utilisateur Firebase
    /// </summary>
    public class FirebaseUser
    {
        public string Email { get; set; }
        public string DisplayName { get; set; }
        public bool Enabled { get; set; }
        public string Role { get; set; }
        public string Site { get; set; }
        public string DisabledMessage { get; set; }
        public long CreatedAt { get; set; }
    }

    /// <summary>
    /// Message broadcast
    /// </summary>
    public class BroadcastMessage
    {
        public bool Active { get; set; }
        public string Title { get; set; }
        public string Message { get; set; }
        public string Type { get; set; } // "info", "warning", "error"
        public long CreatedAt { get; set; }
        public long ExpiresAt { get; set; }
        public string TargetUser { get; set; } // null = tous, sinon username specifique
        public string TargetDevice { get; set; } // null = tous, sinon deviceId specifique
    }
}
