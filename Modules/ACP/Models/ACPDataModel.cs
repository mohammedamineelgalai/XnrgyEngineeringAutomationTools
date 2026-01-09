using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace XnrgyEngineeringAutomationTools.Modules.ACP.Models
{
    /// <summary>
    /// Modèle de données ACP (Assistant de Conception de Projet)
    /// Structure hiérarchique : Unité → Modules → Points Critiques
    /// </summary>
    public class ACPDataModel
    {
        [JsonPropertyName("unitId")]
        public string UnitId { get; set; } = "";  // Format: "PROJECT-REF" ex: "10516-01"

        [JsonPropertyName("projectNumber")]
        public string ProjectNumber { get; set; } = "";

        [JsonPropertyName("reference")]
        public string Reference { get; set; } = "";

        [JsonPropertyName("unitName")]
        public string UnitName { get; set; } = "";  // Nom de l'unité pour affichage

        [JsonPropertyName("createdBy")]
        public string CreatedBy { get; set; } = "";  // Lead CAD ou Gestionnaire qui a créé l'unité

        [JsonPropertyName("createdDate")]
        public DateTime CreatedDate { get; set; } = DateTime.UtcNow;

        [JsonPropertyName("lastModifiedBy")]
        public string LastModifiedBy { get; set; } = "";

        [JsonPropertyName("lastModifiedDate")]
        public DateTime LastModifiedDate { get; set; } = DateTime.UtcNow;

        [JsonPropertyName("version")]
        public int Version { get; set; } = 1;

        [JsonPropertyName("modules")]
        public Dictionary<string, ACPModule> Modules { get; set; } = new Dictionary<string, ACPModule>();  // Clé: "M01", "M02", etc.
    }

    /// <summary>
    /// Module ACP avec ses points critiques
    /// </summary>
    public class ACPModule
    {
        [JsonPropertyName("moduleId")]
        public string ModuleId { get; set; } = "";  // "M01", "M02", etc.

        [JsonPropertyName("moduleName")]
        public string ModuleName { get; set; } = "";

        [JsonPropertyName("assignedTo")]
        public string AssignedTo { get; set; } = "";  // Initiales du dessinateur assigné

        [JsonPropertyName("criticalPoints")]
        public List<CriticalPoint> CriticalPoints { get; set; } = new List<CriticalPoint>();

        [JsonPropertyName("status")]
        public string Status { get; set; } = "En cours";  // "En cours", "Validé", "Approuvé"

        [JsonPropertyName("validatedBy")]
        public string ValidatedBy { get; set; } = "";  // Dessinateur qui a validé

        [JsonPropertyName("validatedDate")]
        public DateTime? ValidatedDate { get; set; }

        [JsonPropertyName("approvedBy")]
        public string ApprovedBy { get; set; } = "";  // Admin ou Lead CAD qui a approuvé

        [JsonPropertyName("approvedDate")]
        public DateTime? ApprovedDate { get; set; }
    }

    /// <summary>
    /// Point critique à vérifier dans un module
    /// </summary>
    public class CriticalPoint
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("category")]
        public string Category { get; set; } = "";  // "Portes", "Plancher", "Murs", "Communication", etc.

        [JsonPropertyName("title")]
        public string Title { get; set; } = "";  // Titre du point critique

        [JsonPropertyName("description")]
        public string Description { get; set; } = "";  // Description détaillée

        [JsonPropertyName("notes")]
        public string Notes { get; set; } = "";  // Notes spéciales pour ce module

        [JsonPropertyName("isValidated")]
        public bool IsValidated { get; set; } = false;

        [JsonPropertyName("validatedBy")]
        public string ValidatedBy { get; set; } = "";  // Initiales du dessinateur

        [JsonPropertyName("validatedDate")]
        public DateTime? ValidatedDate { get; set; }

        [JsonPropertyName("validationComment")]
        public string ValidationComment { get; set; } = "";  // Commentaire du dessinateur

        [JsonPropertyName("isApproved")]
        public bool IsApproved { get; set; } = false;  // Approuvé par Admin/Lead CAD

        [JsonPropertyName("approvedBy")]
        public string ApprovedBy { get; set; } = "";  // Initiales de l'approbateur

        [JsonPropertyName("approvedDate")]
        public DateTime? ApprovedDate { get; set; }

        [JsonPropertyName("approvalComment")]
        public string ApprovalComment { get; set; } = "";  // Commentaire de l'approbateur

        [JsonPropertyName("priority")]
        public string Priority { get; set; } = "Normal";  // "Haute", "Normal", "Basse"

        [JsonPropertyName("isApplicable")]
        public bool IsApplicable { get; set; } = true;  // Si le point est applicable à ce module
    }

    /// <summary>
    /// Données de validation d'un module complet
    /// </summary>
    public class ModuleValidationData
    {
        [JsonPropertyName("unitId")]
        public string UnitId { get; set; } = "";

        [JsonPropertyName("moduleId")]
        public string ModuleId { get; set; } = "";

        [JsonPropertyName("validatedPoints")]
        public Dictionary<int, CriticalPointValidation> ValidatedPoints { get; set; } = new Dictionary<int, CriticalPointValidation>();

        [JsonPropertyName("moduleStatus")]
        public string ModuleStatus { get; set; } = "En cours";

        [JsonPropertyName("lastValidatedBy")]
        public string LastValidatedBy { get; set; } = "";

        [JsonPropertyName("lastValidatedDate")]
        public DateTime? LastValidatedDate { get; set; }
    }

    /// <summary>
    /// Validation d'un point critique individuel
    /// </summary>
    public class CriticalPointValidation
    {
        [JsonPropertyName("pointId")]
        public int PointId { get; set; }

        [JsonPropertyName("isValidated")]
        public bool IsValidated { get; set; }

        [JsonPropertyName("validatedBy")]
        public string ValidatedBy { get; set; } = "";

        [JsonPropertyName("validatedDate")]
        public DateTime ValidatedDate { get; set; } = DateTime.UtcNow;

        [JsonPropertyName("comment")]
        public string Comment { get; set; } = "";

        [JsonPropertyName("isApproved")]
        public bool IsApproved { get; set; }

        [JsonPropertyName("approvedBy")]
        public string ApprovedBy { get; set; } = "";

        [JsonPropertyName("approvedDate")]
        public DateTime? ApprovedDate { get; set; }

        [JsonPropertyName("approvalComment")]
        public string ApprovalComment { get; set; } = "";
    }
}

