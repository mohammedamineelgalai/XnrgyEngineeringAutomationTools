using System.Collections.Generic;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models
{
    /// <summary>
    /// Configuration pour le module "Place Equipment" - modifiable par les admins Vault
    /// Stocke les parametres specifiques au placement d'equipements dans les modules Inventor
    /// </summary>
    public class PlaceEquipmentSettings
    {
        /// <summary>
        /// Liste des initiales des dessinateurs XNRGY
        /// </summary>
        public List<string> DesignerInitials { get; set; } = new List<string>
        {
            "N/A", "AC", "AM", "AR", "CC", "DC", "DL", "FL",
            "IM", "KB", "KJ", "MAE", "MC", "NJ", "RO", "SB", "TG", "TV", "VK", "YS"
        };

        /// <summary>
        /// Chemin local vers le dossier des equipements (Library)
        /// </summary>
        public string EquipmentLocalPath { get; set; } = @"C:\Vault\Engineering\Library\Equipment";

        /// <summary>
        /// Chemin Vault vers le dossier des equipements (Library)
        /// </summary>
        public string EquipmentVaultPath { get; set; } = "$/Engineering/Library/Equipment";

        /// <summary>
        /// Chemin local de base pour les projets (destination des equipements copies)
        /// </summary>
        public string ProjectsLocalBasePath { get; set; } = @"C:\Vault\Engineering\Projects";

        /// <summary>
        /// Chemin Vault de base pour les projets
        /// </summary>
        public string ProjectsVaultBasePath { get; set; } = "$/Engineering/Projects";

        /// <summary>
        /// Nom du sous-dossier pour les equipements dans les modules (ex: "1-Equipment")
        /// </summary>
        public string EquipmentSubfolderName { get; set; } = "1-Equipment";

        /// <summary>
        /// Prefixe pour les dossiers de reference (ex: "REF" pour REF01, REF02...)
        /// </summary>
        public string ReferenceFolderPrefix { get; set; } = "REF";

        /// <summary>
        /// Prefixe pour les dossiers de module (ex: "M" pour M01, M02...)
        /// </summary>
        public string ModuleFolderPrefix { get; set; } = "M";

        /// <summary>
        /// Nombre max de references/modules dans les ComboBox
        /// </summary>
        public int MaxReferenceModuleNumber { get; set; } = 50;

        /// <summary>
        /// Extensions de fichiers Inventor reconnues
        /// </summary>
        public List<string> InventorExtensions { get; set; } = new List<string>
        {
            ".ipt", ".iam", ".idw", ".dwg", ".ipn", ".ide"
        };

        /// <summary>
        /// Extensions de fichiers a exclure
        /// </summary>
        public List<string> ExcludedExtensions { get; set; } = new List<string>
        {
            ".v", ".v1", ".v2", ".v3", ".v4", ".v5", ".vbak", ".bak", ".lck", ".log", ".dwl", ".dwl2"
        };

        /// <summary>
        /// Dossiers a exclure lors du scan
        /// </summary>
        public List<string> ExcludedFolders { get; set; } = new List<string>
        {
            "_V", "OldVersions", "oldversions", ".git", ".vs", "Backup"
        };

        /// <summary>
        /// Dossiers Library dont les liens doivent etre preserves (pas copies)
        /// </summary>
        public List<string> LibraryPathPatterns { get; set; } = new List<string>
        {
            @"\Library\",
            @"\IPT_Typical_Drawing\",
            @"\Cabinet\"
        };

        /// <summary>
        /// Liste des equipements disponibles avec leurs fichiers IPJ et IAM
        /// </summary>
        public List<EquipmentDefinition> AvailableEquipments { get; set; } = new List<EquipmentDefinition>
        {
            new EquipmentDefinition
            {
                Name = "Silencer",
                DisplayName = "Silencer",
                ProjectFileName = "Silencer.ipj",
                AssemblyFileName = "Silencer.iam",
                VaultPath = "$/Engineering/Library/Equipment/Silencer"
            },
            new EquipmentDefinition
            {
                Name = "AngularFilter",
                DisplayName = "Angular Filter",
                ProjectFileName = "AngularFilter.ipj",
                AssemblyFileName = "AngularFilter.iam",
                VaultPath = "$/Engineering/Library/Equipment/AngularFilter"
            },
            new EquipmentDefinition
            {
                Name = "BandCooler",
                DisplayName = "Band Cooler",
                ProjectFileName = "BandCooler.ipj",
                AssemblyFileName = "BandCooler.iam",
                VaultPath = "$/Engineering/Library/Equipment/BandCooler"
            },
            new EquipmentDefinition
            {
                Name = "GasHeater",
                DisplayName = "Gas Heater",
                ProjectFileName = "GasHeater.ipj",
                AssemblyFileName = "GasHeater.iam",
                VaultPath = "$/Engineering/Library/Equipment/GasHeater"
            }
        };

        /// <summary>
        /// Proprietes iProperties a appliquer aux fichiers d'equipement copies
        /// </summary>
        public PlaceEquipmentIPropertySettings IPropertyNames { get; set; } = new PlaceEquipmentIPropertySettings();

        /// <summary>
        /// Option: Activer l'upload automatique vers Vault apres le placement
        /// </summary>
        public bool AutoUploadToVault { get; set; } = true;

        /// <summary>
        /// Option: Ouvrir automatiquement le Top Assembly apres le placement
        /// </summary>
        public bool AutoOpenTopAssembly { get; set; } = true;

        /// <summary>
        /// Option: Creer automatiquement le dossier 1-Equipment s'il n'existe pas
        /// </summary>
        public bool AutoCreateEquipmentFolder { get; set; } = true;
    }

    /// <summary>
    /// Definition d'un equipement disponible dans la Library
    /// </summary>
    public class EquipmentDefinition
    {
        /// <summary>
        /// Nom technique de l'equipement (utilise pour les chemins)
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Nom d'affichage dans l'interface utilisateur
        /// </summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>
        /// Nom du fichier projet (.ipj) de l'equipement
        /// </summary>
        public string ProjectFileName { get; set; } = string.Empty;

        /// <summary>
        /// Nom du fichier assemblage principal (.iam) de l'equipement
        /// </summary>
        public string AssemblyFileName { get; set; } = string.Empty;

        /// <summary>
        /// Chemin Vault complet de l'equipement
        /// </summary>
        public string VaultPath { get; set; } = string.Empty;

        /// <summary>
        /// Description de l'equipement (optionnel)
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// Icone ou image de l'equipement (chemin relatif, optionnel)
        /// </summary>
        public string IconPath { get; set; } = string.Empty;
    }

    /// <summary>
    /// Noms des iProperties utilisees pour les equipements
    /// </summary>
    public class PlaceEquipmentIPropertySettings
    {
        public string DesignerInitial { get; set; } = "Initiale_du_Dessinateur";
        public string CoDesignerInitial { get; set; } = "Initiale_du_Co_Dessinateur";
        public string CreationDate { get; set; } = "Creation_Date";
        public string JobTitle { get; set; } = "Job_Title";
        public string Project { get; set; } = "Projet";
        public string Reference { get; set; } = "Reference";
        public string Module { get; set; } = "Module";
        public string EquipmentType { get; set; } = "Equipment_Type";
        public string EquipmentInstance { get; set; } = "Equipment_Instance";
    }
}
