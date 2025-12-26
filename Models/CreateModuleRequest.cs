using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Models
{
    /// <summary>
    /// Modèle de données pour la création d'un nouveau module XNRGY
    /// Contient toutes les propriétés nécessaires pour les iProperties et paramètres Inventor
    /// </summary>
    public class CreateModuleRequest : INotifyPropertyChanged
    {
        #region Propriétés Dossier (Structure)
        
        private string _project = string.Empty;
        /// <summary>
        /// Numéro de projet (ex: 12345) - Utilisé pour le chemin dossier
        /// </summary>
        public string Project
        {
            get => _project;
            set { _project = value; OnPropertyChanged(); OnPropertyChanged(nameof(FullProjectNumber)); OnPropertyChanged(nameof(DestinationPath)); }
        }

        private string _reference = "01";
        /// <summary>
        /// Référence du projet (ex: 01) - Utilisé pour REF01
        /// </summary>
        public string Reference
        {
            get => _reference;
            set { _reference = value; OnPropertyChanged(); OnPropertyChanged(nameof(FullProjectNumber)); OnPropertyChanged(nameof(DestinationPath)); }
        }

        private string _module = "01";
        /// <summary>
        /// Numéro de module (ex: 01) - Utilisé pour M01
        /// </summary>
        public string Module
        {
            get => _module;
            set { _module = value; OnPropertyChanged(); OnPropertyChanged(nameof(FullProjectNumber)); OnPropertyChanged(nameof(DestinationPath)); }
        }

        #endregion

        #region Propriétés Custom iProperties

        private string _initialeDessinateur = string.Empty;
        /// <summary>
        /// Initiales du dessinateur principal (ex: MAE)
        /// Linké au paramètre Initiale_du_Dessinateur_Form dans Top Assy
        /// </summary>
        public string InitialeDessinateur
        {
            get => _initialeDessinateur;
            set { _initialeDessinateur = value; OnPropertyChanged(); }
        }

        private string _initialeCoDessinateur = string.Empty;
        /// <summary>
        /// Initiales du co-dessinateur (ex: JD)
        /// Linké au paramètre Initiale_du_Co_Dessinateur_Form dans Top Assy
        /// </summary>
        public string InitialeCoDessinateur
        {
            get => _initialeCoDessinateur;
            set { _initialeCoDessinateur = value; OnPropertyChanged(); }
        }

        private DateTime _creationDate = DateTime.Now;
        /// <summary>
        /// Date de création du module
        /// Linké au paramètre Creation_Date_Form dans Top Assy
        /// </summary>
        public DateTime CreationDate
        {
            get => _creationDate;
            set { _creationDate = value; OnPropertyChanged(); OnPropertyChanged(nameof(CreationDateFormatted)); }
        }

        /// <summary>
        /// Date de création formatée pour affichage
        /// </summary>
        public string CreationDateFormatted => _creationDate.ToString("yyyy-MM-dd");

        private string _jobTitle = string.Empty;
        /// <summary>
        /// Titre du projet (Job Title)
        /// Utilisé pour iProperties Custom et pages de couverture PDF
        /// </summary>
        public string JobTitle
        {
            get => _jobTitle;
            set { _jobTitle = value; OnPropertyChanged(); }
        }

        #endregion

        #region Propriétés Calculées

        /// <summary>
        /// Numéro de projet complet (ex: 123450101 = Project + Reference + Module)
        /// Format: XXXXX + XX + XX où X = chiffre
        /// </summary>
        public string FullProjectNumber
        {
            get
            {
                var proj = Project?.PadLeft(5, '0') ?? "00000";
                var refNum = Reference?.PadLeft(2, '0') ?? "01";
                var mod = Module?.PadLeft(2, '0') ?? "01";
                return $"{proj}{refNum}{mod}";
            }
        }

        /// <summary>
        /// Chemin de destination calculé (ex: C:\Vault\Engineering\Projects\12345\REF01\M01)
        /// </summary>
        public string DestinationPath
        {
            get
            {
                var basePath = DestinationBasePath ?? @"C:\Vault\Engineering\Projects";
                var proj = Project ?? "00000";
                var refNum = Reference?.PadLeft(2, '0') ?? "01";
                var mod = Module?.PadLeft(2, '0') ?? "01";
                return System.IO.Path.Combine(basePath, proj, $"REF{refNum}", $"M{mod}");
            }
        }

        /// <summary>
        /// Nom du fichier Top Assembly renommé (ex: 123450101.iam)
        /// </summary>
        public string TopAssemblyNewName => $"{FullProjectNumber}.iam";

        #endregion

        #region Chemins Source et Destination

        private string _sourceTemplatePath = @"$/Engineering/Library";
        /// <summary>
        /// Chemin Vault du template source
        /// </summary>
        public string SourceTemplatePath
        {
            get => _sourceTemplatePath;
            set { _sourceTemplatePath = value; OnPropertyChanged(); }
        }

        private string _sourceExistingProjectPath = string.Empty;
        /// <summary>
        /// Chemin du projet existant à copier
        /// </summary>
        public string SourceExistingProjectPath
        {
            get => _sourceExistingProjectPath;
            set { _sourceExistingProjectPath = value; OnPropertyChanged(); }
        }

        private string _destinationBasePath = @"C:\Vault\Engineering\Projects";
        /// <summary>
        /// Chemin de base de destination (local)
        /// </summary>
        public string DestinationBasePath
        {
            get => _destinationBasePath;
            set { _destinationBasePath = value; OnPropertyChanged(); OnPropertyChanged(nameof(DestinationPath)); }
        }

        private CreateModuleSource _source = CreateModuleSource.FromTemplate;
        /// <summary>
        /// Source de création: Template ou Projet existant
        /// </summary>
        public CreateModuleSource Source
        {
            get => _source;
            set { _source = value; OnPropertyChanged(); }
        }

        #endregion

        #region Options de Renommage

        private string _searchPattern = string.Empty;
        /// <summary>
        /// Pattern de recherche pour renommer les fichiers
        /// </summary>
        public string SearchPattern
        {
            get => _searchPattern;
            set { _searchPattern = value; OnPropertyChanged(); }
        }

        private string _replacePattern = string.Empty;
        /// <summary>
        /// Pattern de remplacement pour renommer les fichiers
        /// </summary>
        public string ReplacePattern
        {
            get => _replacePattern;
            set { _replacePattern = value; OnPropertyChanged(); }
        }

        private string _prefix = string.Empty;
        /// <summary>
        /// Préfixe à ajouter aux noms de fichiers
        /// </summary>
        public string Prefix
        {
            get => _prefix;
            set { _prefix = value; OnPropertyChanged(); }
        }

        private string _suffix = string.Empty;
        /// <summary>
        /// Suffixe à ajouter aux noms de fichiers
        /// </summary>
        public string Suffix
        {
            get => _suffix;
            set { _suffix = value; OnPropertyChanged(); }
        }

        private bool _applyToAllFiles = true;
        /// <summary>
        /// Appliquer le renommage à tous les fichiers
        /// </summary>
        public bool ApplyToAllFiles
        {
            get => _applyToAllFiles;
            set { _applyToAllFiles = value; OnPropertyChanged(); }
        }

        #endregion

        #region Liste des fichiers

        /// <summary>
        /// Liste des fichiers à copier avec leurs nouveaux noms
        /// </summary>
        public ObservableCollection<FileRenameItem> FilesToCopy { get; set; } = new ObservableCollection<FileRenameItem>();

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region Validation

        /// <summary>
        /// Valide que toutes les propriétés requises sont remplies
        /// </summary>
        public (bool IsValid, string ErrorMessage) Validate()
        {
            if (string.IsNullOrWhiteSpace(Project))
                return (false, "Le numéro de projet est requis");
            
            if (Project.Length < 4 || Project.Length > 6)
                return (false, "Le numéro de projet doit contenir entre 4 et 6 chiffres");

            if (string.IsNullOrWhiteSpace(Reference))
                return (false, "La référence est requise");

            if (string.IsNullOrWhiteSpace(Module))
                return (false, "Le numéro de module est requis");

            if (string.IsNullOrWhiteSpace(InitialeDessinateur))
                return (false, "Les initiales du dessinateur sont requises");

            if (Source == CreateModuleSource.FromTemplate && string.IsNullOrWhiteSpace(SourceTemplatePath))
                return (false, "Le chemin du template est requis");

            if (Source == CreateModuleSource.FromExistingProject && string.IsNullOrWhiteSpace(SourceExistingProjectPath))
                return (false, "Le chemin du projet existant est requis");

            return (true, string.Empty);
        }

        #endregion
    }

    /// <summary>
    /// Source de création du module
    /// </summary>
    public enum CreateModuleSource
    {
        /// <summary>
        /// Créer depuis un template ($/Engineering/Library)
        /// </summary>
        FromTemplate,
        
        /// <summary>
        /// Créer depuis un projet existant
        /// </summary>
        FromExistingProject
    }

    /// <summary>
    /// Élément de fichier à renommer dans la liste
    /// </summary>
    public class FileRenameItem : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        /// <summary>
        /// Sélectionné pour la copie
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        private string _originalPath = string.Empty;
        /// <summary>
        /// Chemin original du fichier
        /// </summary>
        public string OriginalPath
        {
            get => _originalPath;
            set { _originalPath = value; OnPropertyChanged(); OnPropertyChanged(nameof(OriginalFileName)); }
        }

        /// <summary>
        /// Nom de fichier original (sans chemin)
        /// </summary>
        public string OriginalFileName => System.IO.Path.GetFileName(OriginalPath);

        private string _relativePath = string.Empty;
        /// <summary>
        /// Chemin relatif depuis la source (pour conserver la structure)
        /// </summary>
        public string RelativePath
        {
            get => _relativePath;
            set { _relativePath = value; OnPropertyChanged(); }
        }

        private string _newFileName = string.Empty;
        /// <summary>
        /// Nouveau nom de fichier
        /// </summary>
        public string NewFileName
        {
            get => _newFileName;
            set { _newFileName = value; OnPropertyChanged(); }
        }

        private string _newPath = string.Empty;
        /// <summary>
        /// Nouveau chemin complet
        /// </summary>
        public string NewPath
        {
            get => _newPath;
            set { _newPath = value; OnPropertyChanged(); }
        }

        private string _fileType = string.Empty;
        /// <summary>
        /// Type de fichier (IAM, IPT, IDW, etc.)
        /// </summary>
        public string FileType
        {
            get => _fileType;
            set { _fileType = value; OnPropertyChanged(); }
        }

        private string _status = "En attente";
        /// <summary>
        /// Status de la copie
        /// </summary>
        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        private bool _isTopAssembly = false;
        /// <summary>
        /// Est-ce le fichier Top Assembly (Module_.iam) ?
        /// </summary>
        public bool IsTopAssembly
        {
            get => _isTopAssembly;
            set { _isTopAssembly = value; OnPropertyChanged(); }
        }

        private bool _isInventorFile = false;
        /// <summary>
        /// Est-ce un fichier Inventor (.iam, .ipt, .idw, .dwg, .ipn) ?
        /// </summary>
        public bool IsInventorFile
        {
            get => _isInventorFile;
            set { _isInventorFile = value; OnPropertyChanged(); }
        }

        private string _destinationPath = string.Empty;
        /// <summary>
        /// Chemin de destination complet
        /// </summary>
        public string DestinationPath
        {
            get => _destinationPath;
            set { _destinationPath = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
