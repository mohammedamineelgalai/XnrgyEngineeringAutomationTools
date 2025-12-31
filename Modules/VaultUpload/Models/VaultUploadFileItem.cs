#nullable enable
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Modules.VaultUpload.Models
{
    /// <summary>
    /// Represente un fichier a uploader vers Vault
    /// </summary>
    public class VaultUploadFileItem : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        private bool _isInventorFile;
        private string _status = "En attente";

        public string FileName { get; set; } = string.Empty;
        public string FileType { get; set; } = string.Empty;
        public string FileSizeFormatted { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string FileExtension { get; set; } = string.Empty;

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
