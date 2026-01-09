#nullable enable
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Modules.UploadModule.Models
{
    /// <summary>
    /// Represente un fichier a uploader vers Vault
    /// </summary>
    public class VaultUploadFileItem : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        private bool _isInventorFile;
        private string _status = "En attente";
        private string _vaultPath = string.Empty;

        public string FileName { get; set; } = string.Empty;
        public string FileType { get; set; } = string.Empty;
        public string FileSizeFormatted { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string RelativePath { get; set; } = string.Empty;
        public string FileExtension { get; set; } = string.Empty;

        /// <summary>
        /// Chemin de destination dans Vault (ex: $/Engineering/Projects/12345/REF01/M01)
        /// </summary>
        public string VaultPath
        {
            get => _vaultPath;
            set
            {
                _vaultPath = value;
                OnPropertyChanged();
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public bool IsInventorFile
        {
            get => _isInventorFile;
            set
            {
                _isInventorFile = value;
                OnPropertyChanged();
            }
        }

        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName ?? string.Empty));
        }
    }
}
