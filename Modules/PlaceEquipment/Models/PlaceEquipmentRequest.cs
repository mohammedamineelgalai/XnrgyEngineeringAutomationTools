using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models
{
    /// <summary>
    /// Modèle de données pour le placement d'un équipement dans un module
    /// </summary>
    public class PlaceEquipmentRequest : INotifyPropertyChanged
    {
        private EquipmentItem _selectedEquipment;
        /// <summary>
        /// Équipement sélectionné
        /// </summary>
        public EquipmentItem SelectedEquipment
        {
            get => _selectedEquipment;
            set { _selectedEquipment = value; OnPropertyChanged(); }
        }

        private string _projectNumber = string.Empty;
        /// <summary>
        /// Numéro de projet (détecté depuis Inventor ou saisi manuellement)
        /// </summary>
        public string ProjectNumber
        {
            get => _projectNumber;
            set { _projectNumber = value; OnPropertyChanged(); OnPropertyChanged(nameof(DestinationPath)); }
        }

        private string _reference = string.Empty;
        /// <summary>
        /// Référence du projet (détectée depuis Inventor ou saisie manuellement)
        /// </summary>
        public string Reference
        {
            get => _reference;
            set { _reference = value; OnPropertyChanged(); OnPropertyChanged(nameof(DestinationPath)); }
        }

        private string _module = string.Empty;
        /// <summary>
        /// Numéro de module (détecté depuis Inventor ou saisi manuellement)
        /// </summary>
        public string Module
        {
            get => _module;
            set { _module = value; OnPropertyChanged(); OnPropertyChanged(nameof(DestinationPath)); }
        }

        private string _topAssemblyPath = string.Empty;
        /// <summary>
        /// Chemin du top assembly actif dans Inventor (ex: C:\Vault\Engineering\Projects\12345\REF01\M02\Master XXXXXXXXX.iam)
        /// </summary>
        public string TopAssemblyPath
        {
            get => _topAssemblyPath;
            set { _topAssemblyPath = value; OnPropertyChanged(); }
        }

        private string _currentProjectPath = string.Empty;
        /// <summary>
        /// Chemin du projet actuel dans Inventor (ex: C:\Vault\Engineering\Projects\12345\REF01\M02)
        /// </summary>
        public string CurrentProjectPath
        {
            get => _currentProjectPath;
            set { _currentProjectPath = value; OnPropertyChanged(); }
        }

        private string _currentProjectFile = string.Empty;
        /// <summary>
        /// Fichier projet (.ipj) actuel dans Inventor
        /// </summary>
        public string CurrentProjectFile
        {
            get => _currentProjectFile;
            set { _currentProjectFile = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Chemin de destination de l'équipement (ex: C:\Vault\Engineering\Projects\12345\REF01\M02\1-Equipment\Silencer)
        /// </summary>
        public string DestinationPath
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ProjectNumber) || string.IsNullOrWhiteSpace(Reference) || string.IsNullOrWhiteSpace(Module))
                    return string.Empty;

                var basePath = @"C:\Vault\Engineering\Projects";
                var refNum = Reference.PadLeft(2, '0');
                var mod = Module.PadLeft(2, '0');
                var equipmentName = SelectedEquipment?.Name ?? string.Empty;
                
                return System.IO.Path.Combine(basePath, ProjectNumber, $"REF{refNum}", $"M{mod}", "1-Equipment", equipmentName);
            }
        }

        /// <summary>
        /// Chemin temporaire de téléchargement (ex: C:\Vault\Engineering\Library\Equipment\Silencer)
        /// </summary>
        public string TempDownloadPath
        {
            get
            {
                var equipmentName = SelectedEquipment?.Name ?? string.Empty;
                return System.IO.Path.Combine(@"C:\Vault\Engineering\Library\Equipment", equipmentName);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Valide que toutes les propriétés requises sont remplies
        /// </summary>
        public (bool IsValid, string ErrorMessage) Validate()
        {
            if (SelectedEquipment == null)
                return (false, "Un équipement doit être sélectionné");

            if (string.IsNullOrWhiteSpace(ProjectNumber))
                return (false, "Le numéro de projet est requis");

            if (string.IsNullOrWhiteSpace(Reference))
                return (false, "La référence est requise");

            if (string.IsNullOrWhiteSpace(Module))
                return (false, "Le numéro de module est requis");

            if (string.IsNullOrWhiteSpace(TopAssemblyPath))
                return (false, "Le top assembly doit être détecté depuis Inventor");

            return (true, string.Empty);
        }
    }
}


