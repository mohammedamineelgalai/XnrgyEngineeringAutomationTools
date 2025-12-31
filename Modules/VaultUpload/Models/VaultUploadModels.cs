#nullable enable
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Modules.VaultUpload.Models
{
    /// <summary>
    /// Categorie Vault pour classification des fichiers
    /// </summary>
    public class VaultCategoryItem
    {
        public long Id { get; set; }
        public string Name { get; set; } = string.Empty;

        public override string ToString() => Name;
    }

    /// <summary>
    /// Lifecycle State Item pour gestion du cycle de vie
    /// </summary>
    public class VaultLifecycleStateItem : INotifyPropertyChanged
    {
        public long Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public long LifecycleDefinitionId { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName ?? string.Empty));
        }
    }

    /// <summary>
    /// Lifecycle Definition Item avec ses etats
    /// </summary>
    public class VaultLifecycleDefinitionItem : INotifyPropertyChanged
    {
        public long Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public ObservableCollection<VaultLifecycleStateItem> States { get; set; } = new();

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName ?? string.Empty));
        }
    }

    /// <summary>
    /// Proprietes du projet detectees depuis le chemin
    /// </summary>
    public class VaultProjectProperties
    {
        public string ProjectNumber { get; set; } = string.Empty;
        public string Reference { get; set; } = string.Empty;
        public string Module { get; set; } = string.Empty;
    }

    /// <summary>
    /// Informations statistiques du projet
    /// </summary>
    public class VaultProjectInfo
    {
        public int TotalFiles { get; set; }
        public int InventorFiles { get; set; }
        public int NonInventorFiles { get; set; }
    }

    /// <summary>
    /// Information sur un module detecte
    /// </summary>
    public class VaultModuleInfo
    {
        public string FullPath { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string ProjectNumber { get; set; } = string.Empty;
        public string Reference { get; set; } = string.Empty;
        public string Module { get; set; } = string.Empty;

        public override string ToString() => DisplayName;
    }
}
