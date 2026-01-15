using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Modules.OpenVaultProject.Models
{
    /// <summary>
    /// Represente un element de projet Vault (Projet, Reference ou Module)
    /// </summary>
    public class VaultProjectItem : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private string _path = string.Empty;
        private string _type = string.Empty;
        private long _entityId;
        private bool _isSelected;
        private bool _isExpanded;
        private DateTime _lastModified;
        private string _status = string.Empty;

        /// <summary>
        /// Nom de l'element (ex: "10359", "REF09", "M03")
        /// </summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Chemin complet dans Vault (ex: "$/Engineering/Projects/10359/REF09/M03")
        /// </summary>
        public string Path
        {
            get => _path;
            set { _path = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Type d'element: "Project", "Reference", "Module"
        /// </summary>
        public string Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// ID de l'entite dans Vault
        /// </summary>
        public long EntityId
        {
            get => _entityId;
            set { _entityId = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Indique si l'element est selectionne
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Indique si le noeud est expanse dans le TreeView
        /// </summary>
        public bool IsExpanded
        {
            get => _isExpanded;
            set { _isExpanded = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Date de derniere modification
        /// </summary>
        public DateTime LastModified
        {
            get => _lastModified;
            set { _lastModified = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Statut de l'element (ex: "Released", "Work in Progress")
        /// </summary>
        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Icone selon le type
        /// </summary>
        public string TypeIcon => Type switch
        {
            "Project" => "📁",
            "Reference" => "📂",
            "Module" => "📦",
            _ => "📄"
        };

        /// <summary>
        /// Affichage complet avec icone
        /// </summary>
        public string DisplayName => $"{TypeIcon} {Name}";

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
