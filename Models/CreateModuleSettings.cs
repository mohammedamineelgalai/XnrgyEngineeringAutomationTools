using System.Collections.Generic;

namespace XnrgyEngineeringAutomationTools.Models
{
    /// <summary>
    /// Configuration pour le module "Créer Module" - modifiable par les admins Vault
    /// </summary>
    public class CreateModuleSettings
    {
        /// <summary>
        /// Liste des initiales des dessinateurs XNRGY
        /// </summary>
        public List<string> DesignerInitials { get; set; } = new List<string>
        {
            "N/A", "AC", "AM", "AP", "AR", "BL", "CC", "CP", "DC", "DL", "DM", "FL",
            "IM", "KB", "KJ", "MAE", "MC", "NJ", "RO", "SB", "TG", "TV", "VK", "YS", "ZM"
        };

        /// <summary>
        /// Chemin local vers le dossier des templates
        /// </summary>
        public string TemplateLocalPath { get; set; } = @"C:\Vault\Engineering\Library\Xnrgy_Module";

        /// <summary>
        /// Chemin Vault vers le dossier des templates
        /// </summary>
        public string TemplateVaultPath { get; set; } = "$/Engineering/Library/Xnrgy_Module";

        /// <summary>
        /// Chemin local de base pour les projets
        /// </summary>
        public string ProjectsLocalBasePath { get; set; } = @"C:\Vault\Engineering\Projects";

        /// <summary>
        /// Chemin Vault de base pour les projets
        /// </summary>
        public string ProjectsVaultBasePath { get; set; } = "$/Engineering/Projects";

        /// <summary>
        /// Chemin vers le fichier IPJ template
        /// </summary>
        public string TemplateIpjFileName { get; set; } = "XXXXX-XX-XX_2026.ipj";

        /// <summary>
        /// Nom du fichier Top Assembly template
        /// </summary>
        public string TemplateTopAssemblyName { get; set; } = "Module_.iam";

        /// <summary>
        /// Préfixe pour les dossiers de référence (ex: "REF" pour REF01, REF02...)
        /// </summary>
        public string ReferenceFolderPrefix { get; set; } = "REF";

        /// <summary>
        /// Préfixe pour les dossiers de module (ex: "M" pour M01, M02...)
        /// </summary>
        public string ModuleFolderPrefix { get; set; } = "M";

        /// <summary>
        /// Nombre max de références/modules dans les ComboBox
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
        /// Extensions de fichiers à exclure
        /// </summary>
        public List<string> ExcludedExtensions { get; set; } = new List<string>
        {
            ".v", ".v1", ".v2", ".v3", ".v4", ".v5", ".vbak", ".bak", ".lck", ".log", ".dwl", ".dwl2"
        };

        /// <summary>
        /// Dossiers à exclure lors du scan
        /// </summary>
        public List<string> ExcludedFolders { get; set; } = new List<string>
        {
            "_V", "OldVersions", "oldversions", ".git", ".vs", "Backup"
        };

        /// <summary>
        /// Dossiers Library dont les liens doivent être préservés (pas copiés)
        /// </summary>
        public List<string> LibraryPathPatterns { get; set; } = new List<string>
        {
            @"\Library\",
            @"\IPT_Typical_Drawing\",
            @"\Cabinet\"
        };

        /// <summary>
        /// Propriétés iProperties à appliquer aux fichiers
        /// </summary>
        public IPropertySettings IPropertyNames { get; set; } = new IPropertySettings();

        /// <summary>
        /// Paramètres Inventor à appliquer au Top Assembly
        /// </summary>
        public InventorParameterSettings InventorParameters { get; set; } = new InventorParameterSettings();
    }

    /// <summary>
    /// Noms des iProperties utilisées
    /// </summary>
    public class IPropertySettings
    {
        public string DesignerInitial { get; set; } = "Initiale_du_Dessinateur";
        public string CoDesignerInitial { get; set; } = "Initiale_du_Co_Dessinateur";
        public string CreationDate { get; set; } = "Creation_Date";
        public string JobTitle { get; set; } = "Job_Title";
        public string Project { get; set; } = "Projet";
        public string Reference { get; set; } = "Reference";
        public string Module { get; set; } = "Module";
    }

    /// <summary>
    /// Noms des paramètres Inventor du Top Assembly
    /// </summary>
    public class InventorParameterSettings
    {
        public string DesignerInitialForm { get; set; } = "Initiale_du_Dessinateur_Form";
        public string CoDesignerInitialForm { get; set; } = "Initiale_du_Co_Dessinateur_Form";
        public string CreationDateForm { get; set; } = "Creation_Date_Form";
    }
}
