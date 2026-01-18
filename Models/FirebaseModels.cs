using System;
using System.Collections.Generic;

namespace XnrgyEngineeringAutomationTools.Models
{
    /// <summary>
    /// Modeles pour la configuration Firebase Realtime Database
    /// Structure alignee avec firebase-init.json
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
    
    // ==========================================
    // NOUVELLES CLASSES - Structure Firebase-init.json
    // ==========================================
    
    /// <summary>
    /// Configuration Kill Switch - /killSwitch
    /// </summary>
    public class KillSwitchConfig
    {
        public KillSwitchGlobal Global { get; set; }
        public Dictionary<string, KillSwitchSite> BySite { get; set; }
        public Dictionary<string, KillSwitchDept> ByDepartment { get; set; }
    }
    
    public class KillSwitchGlobal
    {
        public bool Enabled { get; set; }
        public string Reason { get; set; }
        public string Message { get; set; }
        public string ActivatedAt { get; set; }
        public string ActivatedBy { get; set; }
        public bool AllowAdmins { get; set; }
    }
    
    public class KillSwitchSite
    {
        public bool Enabled { get; set; }
        public string Reason { get; set; }
        public string Message { get; set; }
    }
    
    public class KillSwitchDept
    {
        public bool Enabled { get; set; }
        public string Reason { get; set; }
    }
    
    /// <summary>
    /// Configuration Maintenance - /maintenance
    /// </summary>
    public class MaintenanceConfig
    {
        public bool Enabled { get; set; }
        public string Message { get; set; }
        public string DetailedMessage { get; set; }
        public string StartTime { get; set; }
        public string EstimatedEndTime { get; set; }
        public bool AllowReadOnly { get; set; }
        public bool AllowAdmins { get; set; }
        public bool ShowCountdown { get; set; }
        public MaintenanceScheduled ScheduledMaintenance { get; set; }
    }
    
    public class MaintenanceScheduled
    {
        public bool Enabled { get; set; }
        public string ScheduledStart { get; set; }
        public string ScheduledEnd { get; set; }
        public int NotifyBeforeMinutes { get; set; }
    }
    
    /// <summary>
    /// Configuration Force Update - /forceUpdate
    /// </summary>
    public class ForceUpdateConfig
    {
        public bool Enabled { get; set; }
        public string Message { get; set; }
        public string MinimumVersion { get; set; }
        public bool CriticalUpdate { get; set; }
        public bool BlockAppUntilUpdated { get; set; }
        public int GracePeriodHours { get; set; }
        public string DownloadUrl { get; set; }
        public string Sha256Checksum { get; set; }
    }
    
    /// <summary>
    /// Configuration Updates - /updates
    /// </summary>
    public class UpdatesConfig
    {
        public UpdateLatest Latest { get; set; }
        public Dictionary<string, UpdateHistory> History { get; set; }
        public UpdateChannels Channels { get; set; }
    }
    
    public class UpdateLatest
    {
        public string Version { get; set; }
        public string DownloadUrl { get; set; }
        public string DownloadUrlMirror { get; set; }
        public string ReleaseNotes { get; set; }
        public string ReleaseNotesUrl { get; set; }
        public string PublishedAt { get; set; }
        public string PublishedBy { get; set; }
        public bool IsCritical { get; set; }
        public string FileSize { get; set; }
        public string Sha256 { get; set; }
    }
    
    public class UpdateHistory
    {
        public string Version { get; set; }
        public string PublishedAt { get; set; }
        public string Notes { get; set; }
    }
    
    public class UpdateChannels
    {
        public string Stable { get; set; }
        public string Beta { get; set; }
        public string Dev { get; set; }
    }

    /// <summary>
    /// Configuration du message de bienvenue (legacy - single message)
    /// </summary>
    public class WelcomeMessageConfig
    {
        public bool Enabled { get; set; }
        public string Title { get; set; }
        public string Message { get; set; }
        public string Type { get; set; } // "info", "warning", "success"
        public string UpdatedAt { get; set; }
        public string UpdatedBy { get; set; }
    }

    /// <summary>
    /// Configuration des messages de bienvenue (nouveau systeme multi-messages)
    /// </summary>
    public class WelcomeMessagesConfig
    {
        /// <summary>Message affiche au premier lancement apres installation</summary>
        public WelcomeMessageItem FirstInstall { get; set; }
        
        /// <summary>Message affiche pour un nouvel utilisateur Windows sur un poste deja installe</summary>
        public WelcomeMessageItem NewUserOnDevice { get; set; }
        
        /// <summary>Message global optionnel pour tous les utilisateurs</summary>
        public WelcomeMessageItem Global { get; set; }
    }

    /// <summary>
    /// Element de message de bienvenue
    /// </summary>
    public class WelcomeMessageItem
    {
        public bool Enabled { get; set; }
        public string Title { get; set; }
        public string Message { get; set; }
        public string Type { get; set; } // "info", "warning", "success"
        public bool ShowOnce { get; set; } // Si true, ne s'affiche qu'une fois par utilisateur
        public string TargetUsers { get; set; } // "all", "new", "specific"
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
        
        // Controle par utilisateur (global - optionnel)
        public bool UserDisabled { get; set; }
        public string UserDisabledMessage { get; set; }
        
        // Controle par device (poste de travail entier)
        public bool DeviceDisabled { get; set; }
        public string DeviceDisabledMessage { get; set; }
        public string DeviceDisabledReason { get; set; }
        
        // Controle par utilisateur SUR un device specifique (devices/[ID]/users/[USER])
        public bool DeviceUserDisabled { get; set; }
        public string DeviceUserDisabledMessage { get; set; }
        public string DeviceUserDisabledReason { get; set; }
        
        // Messages personnalises
        public bool HasBroadcastMessage { get; set; }
        public string BroadcastTitle { get; set; }
        public string BroadcastMessage { get; set; }
        public string BroadcastType { get; set; } // "info", "warning", "error"
        
        // Message de bienvenue (affiche au demarrage)
        public bool HasWelcomeMessage { get; set; }
        public string WelcomeTitle { get; set; }
        public string WelcomeMessage { get; set; }
        public string WelcomeType { get; set; }

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
    /// Utilisateur Firebase - Structure compatible ancienne ET nouvelle
    /// Ancienne: { Email, DisplayName, Enabled, ... }
    /// Nouvelle: { profile: { email, fullName }, status: { enabled }, organization: { role, site } }
    /// </summary>
    public class FirebaseUser
    {
        // Ancienne structure (compatibilite)
        public string Email { get; set; }
        public string DisplayName { get; set; }
        public bool Enabled { get; set; } = true; // Par defaut ACTIF
        public string Role { get; set; }
        public string Site { get; set; }
        public string DisabledMessage { get; set; }
        public long CreatedAt { get; set; }
        
        // Nouvelle structure
        public FirebaseUserProfile profile { get; set; }
        public FirebaseUserStatus status { get; set; }
        public FirebaseUserOrganization organization { get; set; }
        
        // Proprietes calculees pour compatibilite
        public string GetEmail() => profile?.email ?? Email ?? "";
        public string GetDisplayName() => profile?.fullName ?? profile?.username ?? DisplayName ?? "";
        public bool IsEnabled() => status?.enabled ?? Enabled;
        public string GetRole() => organization?.role ?? Role ?? "user";
        public string GetSite() => organization?.site ?? Site ?? "";
        public string GetDisabledMessage() => status?.disabledMessage ?? DisabledMessage;
    }
    
    public class FirebaseUserProfile
    {
        public string email { get; set; }
        public string fullName { get; set; }
        public string username { get; set; }
        public string createdAt { get; set; }
    }
    
    public class FirebaseUserStatus
    {
        public bool enabled { get; set; } = true;
        public bool blocked { get; set; } = false;
        public bool online { get; set; }
        public string lastActive { get; set; }
        public string disabledMessage { get; set; }
    }
    
    public class FirebaseUserOrganization
    {
        public string role { get; set; }
        public string site { get; set; }
        public string department { get; set; }
        public string team { get; set; }
    }

    /// <summary>
    /// Message broadcast - Structure mise a jour pour correspondre a Firebase
    /// </summary>
    public class BroadcastMessage
    {
        // Proprietes directes (ancienne structure)
        public bool Active { get; set; }
        public string Title { get; set; }
        public string Message { get; set; }
        public string Type { get; set; } // "info", "warning", "error"
        public long CreatedAt { get; set; }
        public long ExpiresAt { get; set; }
        public string TargetUser { get; set; }
        public string TargetDevice { get; set; }
        
        // Nouvelle structure Firebase imbriquee
        public BroadcastStatus Status { get; set; }
        public BroadcastTarget Target { get; set; }
        public BroadcastDisplay Display { get; set; }
        public string Priority { get; set; }
        public string Id { get; set; }
        
        /// <summary>
        /// Retourne si le broadcast est actif (compatible ancienne et nouvelle structure)
        /// </summary>
        public bool IsActive()
        {
            // Nouvelle structure: status.active
            if (Status != null)
                return Status.Active;
            // Ancienne structure: Active direct
            return Active;
        }
        
        /// <summary>
        /// Retourne le type de cible (compatible ancienne et nouvelle structure)
        /// </summary>
        public string GetTargetType()
        {
            if (Target != null)
                return Target.Type ?? "all";
            return "all";
        }
        
        /// <summary>
        /// Verifie si doit afficher en popup
        /// </summary>
        public bool ShouldShowPopup()
        {
            return Display?.ShowAsPopup ?? false;
        }
    }
    
    public class BroadcastStatus
    {
        public bool Active { get; set; }
        public string SentAt { get; set; }
        public string SentBy { get; set; }
        public int ViewCount { get; set; }
    }
    
    public class BroadcastTarget
    {
        public string Type { get; set; } // "all", "user", "device", "site"
        public string UserId { get; set; }
        public string DeviceId { get; set; }
        public string Site { get; set; }
    }
    
    public class BroadcastDisplay
    {
        public bool Dismissible { get; set; } = true;
        public bool RequireAcknowledgment { get; set; }
        public bool ShowAsBanner { get; set; }
        public bool ShowAsPopup { get; set; }
    }

    /// <summary>
    /// Device enregistre dans Firebase - Structure mise a jour
    /// </summary>
    public class FirebaseDevice
    {
        // Proprietes directes (ancienne structure)
        public string MachineName { get; set; }
        public string UserName { get; set; }
        public string AppVersion { get; set; }
        public bool Enabled { get; set; } = true;
        public string DisabledMessage { get; set; }
        public string DisabledReason { get; set; }
        public long DisabledAt { get; set; }
        public string DisabledBy { get; set; }
        public object Heartbeat { get; set; }
        public object SystemInfo { get; set; }
        
        // Nouvelle structure Firebase imbriquee
        public FirebaseDeviceRegistration Registration { get; set; }
        public FirebaseDeviceStatus Status { get; set; }
        public Dictionary<string, DeviceUser> Users { get; set; }
        
        /// <summary>
        /// Retourne si le device est bloque (compatible ancienne et nouvelle structure)
        /// </summary>
        public bool IsBlocked()
        {
            // Nouvelle structure: status.blocked
            if (Status != null)
                return Status.Blocked;
            // Ancienne structure: Enabled = false
            return !Enabled;
        }
        
        /// <summary>
        /// Retourne le message de blocage
        /// </summary>
        public string GetBlockedMessage()
        {
            if (Status != null && Status.Blocked)
            {
                string reason = Status.BlockReason ?? "suspended";
                switch (reason.ToLowerInvariant())
                {
                    case "maintenance":
                        return $"Ce poste ({GetMachineName()}) est en maintenance.";
                    case "unauthorized":
                        return $"Ce poste ({GetMachineName()}) n'est pas autorise.";
                    case "security":
                        return $"Ce poste ({GetMachineName()}) a ete bloque pour raison de securite.";
                    default:
                        return $"Ce poste ({GetMachineName()}) a ete suspendu par l'administrateur.";
                }
            }
            return DisabledMessage ?? "Ce poste a ete desactive.";
        }
        
        /// <summary>
        /// Retourne le nom de la machine
        /// </summary>
        public string GetMachineName()
        {
            return Registration?.MachineName ?? MachineName ?? Environment.MachineName;
        }
    }
    
    public class FirebaseDeviceRegistration
    {
        public string MachineName { get; set; }
        public string RegisteredAt { get; set; }
        public string AppVersion { get; set; }
    }
    
    public class FirebaseDeviceStatus
    {
        public bool Online { get; set; }
        public bool Enabled { get; set; } = true;
        public bool Blocked { get; set; }
        public string BlockReason { get; set; }
        public string BlockedAt { get; set; }
        public string BlockedBy { get; set; }
        public string CurrentUser { get; set; }
        public string LastSeen { get; set; }
    }

    /// <summary>
    /// Utilisateur specifique sur un device - permet de bloquer un utilisateur Windows sur un poste precis
    /// </summary>
    public class DeviceUser
    {
        public bool Enabled { get; set; } = true;
        public string DisabledMessage { get; set; }
        public string DisabledReason { get; set; } // "suspended", "unauthorized", "revoked"
        public long DisabledAt { get; set; }
        public string DisabledBy { get; set; }
        public long LastSeen { get; set; }
    }
}
