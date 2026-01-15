using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace XnrgyEngineeringAutomationTools.Modules.ChecklistHVAC.Models
{
    /// <summary>
    /// Modèle de données pour les réponses de checklist HVAC
    /// Structure JSON stockée dans Vault pour synchronisation multi-utilisateurs
    /// </summary>
    public class ChecklistDataModel
    {
        [JsonPropertyName("moduleId")]
        public string ModuleId { get; set; } = "";  // Format: "PROJECT-REF-MODULE" ex: "25001-01-01"

        [JsonPropertyName("projectNumber")]
        public string ProjectNumber { get; set; } = "";

        [JsonPropertyName("reference")]
        public string Reference { get; set; } = "";

        [JsonPropertyName("module")]
        public string Module { get; set; } = "";

        [JsonPropertyName("lastModifiedBy")]
        public string LastModifiedBy { get; set; } = "";  // Initiales utilisateur

        [JsonPropertyName("lastModifiedDate")]
        public DateTime LastModifiedDate { get; set; } = DateTime.UtcNow;

        [JsonPropertyName("version")]
        public int Version { get; set; } = 1;  // Incrémenté à chaque modification

        [JsonPropertyName("responses")]
        public Dictionary<int, CheckpointResponse> Responses { get; set; } = new Dictionary<int, CheckpointResponse>();

        [JsonPropertyName("moduleName")]
        public string ModuleName { get; set; } = "";  // Nom du module pour affichage
    }

    /// <summary>
    /// Réponse d'un checkpoint individuel
    /// </summary>
    public class CheckpointResponse
    {
        [JsonPropertyName("checkpointId")]
        public int CheckpointId { get; set; }

        [JsonPropertyName("status")]
        public string Status { get; set; } = "";  // "fait", "non_applicable", "pas_fait"

        [JsonPropertyName("comment")]
        public string Comment { get; set; } = "";

        [JsonPropertyName("userInitials")]
        public string UserInitials { get; set; } = "";  // Initiales de l'utilisateur qui a répondu

        [JsonPropertyName("modifiedDate")]
        public DateTime ModifiedDate { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Métadonnées de synchronisation pour comparer les versions
    /// </summary>
    public class ChecklistSyncMetadata
    {
        public string ModuleId { get; set; } = "";
        public DateTime LastModifiedDate { get; set; }
        public string LastModifiedBy { get; set; } = "";
        public int Version { get; set; }
        public long FileSize { get; set; }
        public string VaultFilePath { get; set; } = "";
        public long? VaultFileId { get; set; }
    }
}



