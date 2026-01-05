using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models
{
    /// <summary>
    /// Modèle représentant un équipement avec ses fichiers .ipj et .iam principaux
    /// </summary>
    public class EquipmentItem : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        /// <summary>
        /// Nom de l'équipement (ex: "Silencer", "AngularFilter")
        /// </summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        private string _displayName = string.Empty;
        /// <summary>
        /// Nom d'affichage de l'équipement (ex: "Silencer", "Angular Filter")
        /// </summary>
        public string DisplayName
        {
            get => _displayName;
            set { _displayName = value; OnPropertyChanged(); }
        }

        private string _projectFileName = string.Empty;
        /// <summary>
        /// Nom du fichier projet (.ipj) - ESSENTIEL pour Copy Design
        /// </summary>
        public string ProjectFileName
        {
            get => _projectFileName;
            set { _projectFileName = value; OnPropertyChanged(); }
        }

        private string _assemblyFileName = string.Empty;
        /// <summary>
        /// Nom du fichier assemblage principal (.iam) à insérer dans le top assembly
        /// </summary>
        public string AssemblyFileName
        {
            get => _assemblyFileName;
            set { _assemblyFileName = value; OnPropertyChanged(); }
        }

        private string _vaultPath = string.Empty;
        /// <summary>
        /// Chemin Vault de l'équipement (ex: $/Engineering/Library/Equipment/Silencer)
        /// </summary>
        public string VaultPath
        {
            get => _vaultPath;
            set { _vaultPath = value; OnPropertyChanged(); }
        }

        private string _localTempPath = string.Empty;
        /// <summary>
        /// Chemin local temporaire après téléchargement (ex: C:\Vault\Engineering\Library\Equipment\Silencer)
        /// </summary>
        public string LocalTempPath
        {
            get => _localTempPath;
            set { _localTempPath = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}


