#nullable enable
// AppMainViewModel.cs - Version .NET Framework 4.8 SANS Source Generators
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using XnrgyEngineeringAutomationTools.Models;
using XnrgyEngineeringAutomationTools.Views;
using XnrgyEngineeringAutomationTools.Shared.Views;

namespace XnrgyEngineeringAutomationTools.ViewModels
{
    public class AppMainViewModel : INotifyPropertyChanged
    {
        private readonly Services.VaultSdkService _vaultService;
      
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name ?? string.Empty));
      
        public AppMainViewModel()
        {
            _vaultService = new Services.VaultSdkService();
  LoadConfiguration();
        }
    
        // Configuration
        public ApplicationConfiguration Configuration { get; } = new();
        public bool SaveCredentials { get; set; }
        public string VaultPassword { get; set; } = "";
   
    private bool _isConnected;
        public bool IsConnected { get => _isConnected; set { _isConnected = value; OnPropertyChanged(); OnPropertyChanged(nameof(ConnectionButtonText)); } }
        public string ConnectionButtonText => IsConnected ? "Déconnecter" : "Connecter";
      
        // Collections
      public ObservableCollection<FileItem> InventorFiles { get; } = new();
        public ObservableCollection<FileItem> NonInventorFiles { get; } = new();
        
        // Liste complète pour les filtres
        private readonly List<FileItem> _allFiles = new();

        // [+] Extensions a EXCLURE (backup, temporaires, systeme)
        private static readonly HashSet<string> ExcludedExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".v", ".bak", ".old",          // Backup Vault
            ".tmp", ".temp",               // Temporaires
            ".ipj",                        // Projet Inventor (ne pas uploader!)
            ".lck", ".lock", ".log",       // Systeme/logs
            ".dwl", ".dwl2"                // AutoCAD locks
        };
        
        // [+] Prefixes de fichiers temporaires a exclure
        private static readonly string[] ExcludedPrefixes = new[]
        {
            "~$",      // Office temporaire
            "._",      // macOS temporaire
            "Backup_", // Backup generique
            ".~"       // Temporaire generique
        };
        
        // [+] Dossiers a exclure completement
        private static readonly string[] ExcludedFolders = new[]
        {
            "OldVersions", "oldversions", 
            "Backup", "backup",
            ".vault", ".git", ".vs"
        };
        
        // Filtres de recherche
        private string _searchFilterInventor = string.Empty;
        public string SearchFilterInventor 
        { 
            get => _searchFilterInventor; 
            set 
            { 
                _searchFilterInventor = value; 
                OnPropertyChanged(); 
                ApplySearchFilterInventor();
            } 
        }
        
        private string _searchFilterNonInventor = string.Empty;
        public string SearchFilterNonInventor 
        { 
            get => _searchFilterNonInventor; 
            set 
            { 
                _searchFilterNonInventor = value; 
                OnPropertyChanged(); 
                ApplySearchFilterNonInventor();
            } 
        }
        
        // Catégories Vault
        private ObservableCollection<CategoryItem> _availableCategories = new();
        public ObservableCollection<CategoryItem> AvailableCategories
        {
            get => _availableCategories;
            set
            {
                _availableCategories = value;
                OnPropertyChanged();
            }
        }
        
        private CategoryItem? _selectedCategoryInventor;
        public CategoryItem? SelectedCategoryInventor
        {
            get => _selectedCategoryInventor;
            set
            {
                _selectedCategoryInventor = value;
                OnPropertyChanged();
                // Mettre à jour les états quand la catégorie change
                UpdateAvailableStates();
            }
        }
        
        private CategoryItem? _selectedCategoryNonInventor;
        public CategoryItem? SelectedCategoryNonInventor
        {
            get => _selectedCategoryNonInventor;
            set
            {
                _selectedCategoryNonInventor = value;
                OnPropertyChanged();
                // Mettre à jour les états quand la catégorie change
                UpdateAvailableStates();
            }
        }

        // Lifecycle Definitions et States
        private ObservableCollection<Models.LifecycleDefinitionItem> _availableLifecycleDefinitions = new();
        public ObservableCollection<Models.LifecycleDefinitionItem> AvailableLifecycleDefinitions
        {
            get => _availableLifecycleDefinitions;
            set
            {
                _availableLifecycleDefinitions = value;
                OnPropertyChanged();
            }
        }

        private ObservableCollection<Models.LifecycleStateItem> _availableStatesInventor = new();
        public ObservableCollection<Models.LifecycleStateItem> AvailableStatesInventor
        {
            get => _availableStatesInventor;
            set
            {
                _availableStatesInventor = value;
                OnPropertyChanged();
            }
        }

        private ObservableCollection<Models.LifecycleStateItem> _availableStatesNonInventor = new();
        public ObservableCollection<Models.LifecycleStateItem> AvailableStatesNonInventor
        {
            get => _availableStatesNonInventor;
            set
            {
                _availableStatesNonInventor = value;
                OnPropertyChanged();
            }
        }

        private Models.LifecycleStateItem? _selectedStateInventor;
        public Models.LifecycleStateItem? SelectedStateInventor
        {
            get => _selectedStateInventor;
            set
            {
                _selectedStateInventor = value;
                OnPropertyChanged();
                // Log pour debug
                if (value != null)
                {
                    Services.Logger.Log($"[ViewModel] [>] SelectedStateInventor change: ID={value.Id}, Name='{value.Name}'", Services.Logger.LogLevel.INFO);
                }
            }
        }

        private Models.LifecycleStateItem? _selectedStateNonInventor;
        public Models.LifecycleStateItem? SelectedStateNonInventor
        {
            get => _selectedStateNonInventor;
            set
            {
                _selectedStateNonInventor = value;
                OnPropertyChanged();
                // Log pour debug
                if (value != null)
                {
                    Services.Logger.Log($"[ViewModel] [>] SelectedStateNonInventor change: ID={value.Id}, Name='{value.Name}'", Services.Logger.LogLevel.INFO);
                }
            }
        }
        
        // Status
        private string _statusMessage = "Pret";
        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage == value) return;
                _statusMessage = value;
                OnPropertyChanged(nameof(StatusMessage));
            }
        }
        public string ProgressText { get; set; } = "0/0";
        public int ProgressValue { get; set; } = 0;
        public int ProgressMaximum { get; set; } = 100;
        public bool IsProcessing => IsCheckingIn || IsAddingToVault;
        public bool IsCheckingIn { get; set; }
        public bool IsAddingToVault { get; set; }
        
        private bool _isPaused;
        public bool IsPaused { get => _isPaused; set { _isPaused = value; OnPropertyChanged(); OnPropertyChanged(nameof(PauseButtonText)); } }
        public string PauseButtonText => IsPaused ? "REPRENDRE" : "PAUSE";
        
        // Projet
        public string ProjectPath { get; set; } = "";
        public ProjectInfo? ProjectInfo { get; set; }
        public ProjectProperties? ProjectProperties { get; set; }
        
        // Comment for check-in
        private string _comment = "Premier check-in effectué par Vault Automation Tool";
        public string Comment 
        { 
            get => _comment; 
            set 
            { 
                _comment = value; 
                OnPropertyChanged(); 
            } 
        }
        
        // Charger les catégories depuis Vault
        private void LoadCategories()
        {
            try
            {
                var categories = _vaultService.GetAvailableCategories();
                var categoryItems = categories.Select(c => new CategoryItem { Id = c.Id, Name = c.Name }).ToList();
                
                Application.Current.Dispatcher.Invoke(() =>
                {
                    AvailableCategories.Clear();
                    foreach (var item in categoryItems)
                    {
                        AvailableCategories.Add(item);
                    }
                    
                    // Sélectionner "Engineering" par défaut si disponible, sinon la première
                    var engineeringCategory = AvailableCategories.FirstOrDefault(c => c.Name.Equals("Engineering", StringComparison.OrdinalIgnoreCase));
                    if (engineeringCategory != null)
                    {
                        SelectedCategoryInventor = engineeringCategory;
                        SelectedCategoryNonInventor = engineeringCategory;
                    }
                    else if (AvailableCategories.Count > 0)
                    {
                        SelectedCategoryInventor = AvailableCategories.First();
                        SelectedCategoryNonInventor = AvailableCategories.First();
                    }
                    
                    // Mettre a jour les etats disponibles apres selection de la categorie
                    UpdateAvailableStates();
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"[!] Erreur lors du chargement des categories: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        private void LoadLifecycleDefinitions()
        {
            try
            {
                var lifecycleDefs = _vaultService.GetAvailableLifecycleDefinitions();
                
                Application.Current.Dispatcher.Invoke(() =>
                {
                    AvailableLifecycleDefinitions.Clear();
                    foreach (var def in lifecycleDefs)
                    {
                        AvailableLifecycleDefinitions.Add(def);
                    }
                    
                    // Mettre a jour les etats disponibles apres chargement
                    UpdateAvailableStates();
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"[!] Erreur lors du chargement des Lifecycle Definitions: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        private void UpdateAvailableStates()
        {
            try
            {
                // Déterminer la catégorie sélectionnée et obtenir le Lifecycle Definition correspondant
                CategoryItem? selectedCategory = SelectedCategoryInventor ?? SelectedCategoryNonInventor;
                if (selectedCategory == null || string.IsNullOrEmpty(selectedCategory.Name))
                {
                    AvailableStatesInventor.Clear();
                    AvailableStatesNonInventor.Clear();
                    return;
                }

                // Vérifier si la catégorie est "Base" (pas de Lifecycle)
                string categoryLower = selectedCategory.Name.ToLowerInvariant().Trim();
                if (categoryLower == "base")
                {
                    // Base n'a pas de States - vider les listes
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (SelectedCategoryInventor != null && selectedCategory == SelectedCategoryInventor)
                        {
                            AvailableStatesInventor.Clear();
                            SelectedStateInventor = null;
                        }
                        if (SelectedCategoryNonInventor != null && selectedCategory == SelectedCategoryNonInventor)
                        {
                            AvailableStatesNonInventor.Clear();
                            SelectedStateNonInventor = null;
                        }
                    });
                    return;
                }

                // Obtenir le Lifecycle Definition ID selon la catégorie
                long? lifecycleDefId = _vaultService.GetLifecycleDefinitionIdByCategory(selectedCategory.Name);
                
                // Trouver le Lifecycle Definition correspondant
                Models.LifecycleDefinitionItem? lifecycleDef = null;
                
                if (lifecycleDefId.HasValue)
                {
                    // Trouver le Lifecycle Definition dans la collection
                    lifecycleDef = AvailableLifecycleDefinitions.FirstOrDefault(d => d.Id == lifecycleDefId.Value);
                }
                
                // Si pas de mapping trouvé, ne pas utiliser de fallback pour respecter la config Vault
                if (lifecycleDef == null || lifecycleDef.States.Count == 0)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (SelectedCategoryInventor != null && selectedCategory == SelectedCategoryInventor)
                        {
                            AvailableStatesInventor.Clear();
                            SelectedStateInventor = null;
                        }
                        if (SelectedCategoryNonInventor != null && selectedCategory == SelectedCategoryNonInventor)
                        {
                            AvailableStatesNonInventor.Clear();
                            SelectedStateNonInventor = null;
                        }
                    });
                    return;
                }

                // Filtrer les States selon ce qui est vraiment disponible dans Vault Client
                // Engineering: For Review, Work in Progress, Released, Obsolete (PAS Quick-Change)
                // Office: Work in Progress, Released, Obsolete
                // Design Representation: Released, Work in Progress, Obsolete
                var allowedStates = GetAllowedStatesForCategory(categoryLower);

                // Mettre à jour les états disponibles
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (SelectedCategoryInventor != null && selectedCategory == SelectedCategoryInventor)
                    {
                        AvailableStatesInventor.Clear();
                        foreach (var state in lifecycleDef.States)
                        {
                            // Filtrer si une liste de States autorisés existe
                            if (allowedStates == null || allowedStates.Any(allowed => 
                                state.Name.IndexOf(allowed, StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                AvailableStatesInventor.Add(state);
                            }
                        }
                        // Sélectionner l'état "Work in Progress" par défaut
                        SelectedStateInventor = AvailableStatesInventor.FirstOrDefault(s => 
                            s.Name.IndexOf("Work", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            s.Name.IndexOf("Progress", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            s.Name.IndexOf("WIP", StringComparison.OrdinalIgnoreCase) >= 0) 
                            ?? AvailableStatesInventor.FirstOrDefault();
                    }
                    
                    if (SelectedCategoryNonInventor != null && selectedCategory == SelectedCategoryNonInventor)
                    {
                        AvailableStatesNonInventor.Clear();
                        foreach (var state in lifecycleDef.States)
                        {
                            // Filtrer si une liste de States autorisés existe
                            if (allowedStates == null || allowedStates.Any(allowed => 
                                state.Name.IndexOf(allowed, StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                AvailableStatesNonInventor.Add(state);
                            }
                        }
                        // Sélectionner l'état "Work in Progress" par défaut
                        SelectedStateNonInventor = AvailableStatesNonInventor.FirstOrDefault(s => 
                            s.Name.IndexOf("Work", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            s.Name.IndexOf("Progress", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            s.Name.IndexOf("WIP", StringComparison.OrdinalIgnoreCase) >= 0) 
                            ?? AvailableStatesNonInventor.FirstOrDefault();
                    }
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"[!] Erreur lors de la mise a jour des etats: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        /// <summary>
        /// Retourne la liste des States autorises pour une categorie donnee.
        /// Ces listes correspondent exactement a ce qui est visible dans Vault Client.
        /// </summary>
        private List<string>? GetAllowedStatesForCategory(string categoryLower)
        {
            // Mapping des States autorises par categorie (correspondant a Vault Client)
            var categoryAllowedStates = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
            {
                // Engineering: For Review, Work in Progress, Released, Obsolete (PAS Quick-Change)
                { "engineering", new List<string> { "For Review", "Work in Progress", "Released", "Obsolete" } },
                
                // Office: Pas de filtre - afficher tous les States de Simple Release Process
                // (Work in Progress, Released, Obsolete)
                { "office", null! },
                
                // Design Representation: Pas de filtre - afficher tous les States
                { "design representation", null! },
                
                // Standard: utiliser tous les States du Lifecycle (pas de filtre)
                { "standard", null! }
            };

            if (categoryAllowedStates.TryGetValue(categoryLower, out var allowedStates))
            {
                return allowedStates;
            }

            // Par defaut, pas de filtre (tous les States sont autorises)
            return null;
        }

        private void LoadCategoriesOld()
        {
            try
            {
                AvailableCategories.Clear();
                
                // Ajouter une option "Aucune" (par defaut)
                AvailableCategories.Add(new CategoryItem { Id = -1, Name = "Aucune (par defaut)" });
                
                // Charger les categories depuis Vault
                var categories = _vaultService.GetAvailableCategories();
                foreach (var category in categories)
                {
                    AvailableCategories.Add(new CategoryItem { Id = category.Id, Name = category.Name });
                }
                
                // Selectionner "Base" par defaut si disponible, sinon "Aucune"
                var baseCategory = AvailableCategories.FirstOrDefault(c => c.Name.Equals("Base", StringComparison.OrdinalIgnoreCase));
                if (baseCategory != null)
                {
                    SelectedCategoryInventor = baseCategory;
                    SelectedCategoryNonInventor = baseCategory;
                }
                else
                {
                    SelectedCategoryInventor = AvailableCategories.First();
                    SelectedCategoryNonInventor = AvailableCategories.First();
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"[!] Erreur chargement categories: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }
        
        // Commandes
        public ICommand ToggleConnectionCommand => new RelayCommand(ToggleConnection);
        public ICommand SelectModuleCommand => new RelayCommand(SelectModule);
        public ICommand BrowseFolderCommand => new RelayCommand(BrowseFolder);
        public ICommand GetFromInventorCommand => new RelayCommand(GetPathFromInventor);
        public ICommand ScanProjectCommand => new RelayCommand(ScanProject);
        public ICommand SelectAllInventorCommand => new RelayCommand(SelectAllInventor);
        public ICommand DeselectAllInventorCommand => new RelayCommand(DeselectAllInventor);
        public ICommand SelectAllNonInventorCommand => new RelayCommand(SelectAllNonInventor);
        public ICommand DeselectAllNonInventorCommand => new RelayCommand(DeselectAllNonInventor);
        public ICommand AutoCheckInCommand => new RelayCommand(async () => await AutoCheckInAsync());
        public ICommand PauseCommand => new RelayCommand(() => IsPaused = !IsPaused);
        public ICommand StopProcessingCommand => new RelayCommand(StopProcessing);
        public ICommand CancelProcessingCommand => new RelayCommand(CancelProcessing);
        
        private void SelectModule()
        {
            try
            {
                StatusMessage = "[>] Scan des modules disponibles...";
                OnPropertyChanged(nameof(StatusMessage));

                const string basePath = @"C:\Vault\Engineering\Projects";

                if (!Directory.Exists(basePath))
                {
                    StatusMessage = $"[-] Dossier Vault introuvable: {basePath}";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                // Scanner tous les modules disponibles
                var modules = ScanAvailableModules(basePath);

                if (modules.Count == 0)
                {
                    StatusMessage = $"[i] Aucun module trouve dans {basePath}. Verifiez la structure.";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                StatusMessage = $"[+] {modules.Count} modules trouves";
                OnPropertyChanged(nameof(StatusMessage));

                // Creer la fenetre de selection
                var window = new ModuleSelectionWindow(modules);
                
                // Definir la fenetre parente pour centrer sur MainWindow
                window.Owner = Application.Current.MainWindow;
                
                if (window.ShowDialog() == true)
                {
                    var selected = window.SelectedModule;
                    if (selected != null)
                    {
                        ProjectPath = selected.FullPath;
                        ProjectProperties = new ProjectProperties
                        {
                            ProjectNumber = selected.ProjectNumber,
                            Reference = selected.Reference,
                            Module = selected.Module
                        };
                        StatusMessage = $"[+] Module selectionne: {selected.DisplayName}";
                        OnPropertyChanged(nameof(ProjectPath));
                        OnPropertyChanged(nameof(ProjectProperties));
                        OnPropertyChanged(nameof(StatusMessage));
                        
                        // Scanner automatiquement le module selectionne
                        ScanProject();
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"[-] Erreur scan modules: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        /// <summary>
        /// Scanne tous les modules disponibles dans C:\Vault\Engineering\Projects
        /// Pattern: [NUMERO]\REF[NUM]\M[NUM]
        /// Exemple: 10359\REF09\M03
        /// </summary>
        private List<ModuleInfo> ScanAvailableModules(string basePath)
        {
            var modules = new List<ModuleInfo>();

            try
            {
                // Scanner tous les dossiers projets (numériques: 10359, 10360, etc.)
                var projectDirs = Directory.GetDirectories(basePath)
                    .Where(d => Regex.IsMatch(Path.GetFileName(d), @"^\d+$"))
                    .OrderBy(d => d);

                foreach (var projectDir in projectDirs)
                {
                    string projectNum = Path.GetFileName(projectDir);

                    // Scanner les références REF[NUM]
                    var refDirs = Directory.GetDirectories(projectDir)
                        .Where(d => Regex.IsMatch(Path.GetFileName(d), @"^REF\d+$", RegexOptions.IgnoreCase))
                        .OrderBy(d => d);

                    foreach (var refDir in refDirs)
                    {
                        string refFull = Path.GetFileName(refDir); // "REF09"
                        string refNum = refFull.Substring(3); // "09"

                        // Scanner les modules M[NUM]
                        var moduleDirs = Directory.GetDirectories(refDir)
                            .Where(d => Regex.IsMatch(Path.GetFileName(d), @"^M\d+$", RegexOptions.IgnoreCase))
                            .OrderBy(d => d);

                        foreach (var moduleDir in moduleDirs)
                        {
                            string moduleFull = Path.GetFileName(moduleDir); // "M03"
                            string moduleNum = moduleFull.Substring(1); // "03"

                            // Compter les fichiers dans le module
                            int fileCount = 0;
                            try
                            {
                                fileCount = Directory.GetFiles(moduleDir, "*.*", SearchOption.AllDirectories).Length;
                            }
                            catch
                            {
                                // Ignorer les erreurs d'accès
                            }

                            if (fileCount > 0)
                            {
                                modules.Add(new ModuleInfo
                                {
                                    ProjectNumber = projectNum,
                                    Reference = refNum,
                                    Module = moduleNum,
                                    FullPath = moduleDir,
                                    DisplayName = $"{projectNum} / {refFull} / {moduleFull} ({fileCount} fichiers)"
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"[-] Erreur lors du scan: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }

            return modules;
        }

        // Scan les fichiers du module selectionne (ProjectPath) et remplit les collections
        private void ScanProject()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ProjectPath) || !Directory.Exists(ProjectPath))
                {
                    StatusMessage = $"[-] Dossier module introuvable: {ProjectPath}";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                // ═══════════════════════════════════════════════════════════════════════════════
                // EXTRACTION AUTOMATIQUE DES PROPRIETES DEPUIS LE CHEMIN
                // Pattern: C:\Vault\Engineering\Projects\[PROJECT]\[REF]\[MODULE]
                // Exemple: C:\Vault\Engineering\Projects\12345\REF01\M01
                // ═══════════════════════════════════════════════════════════════════════════════
                if (ProjectProperties == null || 
                    string.IsNullOrWhiteSpace(ProjectProperties.ProjectNumber) ||
                    string.IsNullOrWhiteSpace(ProjectProperties.Reference) ||
                    string.IsNullOrWhiteSpace(ProjectProperties.Module))
                {
                    var extractedProps = ExtractPropertiesFromPath(ProjectPath);
                    if (extractedProps != null)
                    {
                        ProjectProperties = extractedProps;
                        OnPropertyChanged(nameof(ProjectProperties));
                        Services.Logger.Log($"[+] Proprietes extraites du chemin: Project={extractedProps.ProjectNumber}, Ref={extractedProps.Reference}, Module={extractedProps.Module}", Services.Logger.LogLevel.INFO);
                    }
                }

                InventorFiles.Clear();
                NonInventorFiles.Clear();
                _allFiles.Clear();

                var allFiles = Directory.GetFiles(ProjectPath, "*.*", SearchOption.AllDirectories);
                var inventorExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".ipt", ".iam", ".idw", ".ipn" };

                int excludedCount = 0;
                int inventorCount = 0;
                int nonInventorCount = 0;

                // Trier par nom de fichier alphabetiquement
                foreach (var f in allFiles.OrderBy(p => Path.GetFileName(p), StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var fi = new FileInfo(f);
                        
                        // [+] EXCLURE fichiers backup/temporaires/systeme
                        bool shouldExclude = false;
                        
                        // Verifier extension
                        if (ExcludedExtensions.Contains(fi.Extension))
                        {
                            excludedCount++;
                            shouldExclude = true;
                        }
                        
                        // Verifier prefixe
                        if (!shouldExclude && ExcludedPrefixes.Any(prefix => fi.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                        {
                            excludedCount++;
                            shouldExclude = true;
                        }
                        
                        // Verifier dossier parent (OldVersions, Backup, etc.)
                        if (!shouldExclude)
                        {
                            var dirName = fi.Directory?.Name ?? "";
                            if (ExcludedFolders.Any(excluded => dirName.Equals(excluded, StringComparison.OrdinalIgnoreCase)))
                            {
                                excludedCount++;
                                shouldExclude = true;
                            }
                        }
                        
                        // Si exclu, passer au suivant
                        if (shouldExclude)
                        {
                            continue;
                        }
                        
                        var item = new FileItem
                        {
                            FileName = fi.Name,
                            FullPath = fi.FullName,
                            FileType = fi.Extension,
                            FileExtension = fi.Extension,
                            FileSizeFormatted = FormatSize(fi.Length),
                            IsInventorFile = inventorExt.Contains(fi.Extension),
                            IsSelected = true, // [+] AUTO-COCHE pour TOUS les fichiers (Inventor ET Non-Inventor)
                            Status = "En attente"
                        };

                        _allFiles.Add(item);
                        
                        if (item.IsInventorFile)
                        {
                            InventorFiles.Add(item);
                            inventorCount++;
                        }
                        else
                        {
                            NonInventorFiles.Add(item);
                            nonInventorCount++;
                        }
                    }
                    catch
                    {
                        // Ignore file access errors
                    }
                }

                ProjectInfo = new ProjectInfo
                {
                    TotalFiles = InventorFiles.Count + NonInventorFiles.Count,
                    InventorFiles = InventorFiles.Count,
                    NonInventorFiles = NonInventorFiles.Count
                };
                OnPropertyChanged(nameof(ProjectInfo));
                
                var statusMsg = $"[+] Scanne: {ProjectInfo.TotalFiles} fichiers ({inventorCount} inventor, {nonInventorCount} non-inventor)";
                if (excludedCount > 0)
                {
                    statusMsg += $" | {excludedCount} exclus (backup/temp)";
                }
                StatusMessage = statusMsg;
                OnPropertyChanged(nameof(StatusMessage));
            }
            catch (Exception ex)
            {
                StatusMessage = $"[-] Erreur scan projet: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        /// <summary>
        /// Extrait les proprietes Project/Reference/Module depuis un chemin de dossier
        /// Pattern: ...\Projects\[PROJECT]\[REF]\[MODULE]
        /// Exemple: C:\Vault\Engineering\Projects\12345\REF01\M01
        /// Résultat: Project=12345, Reference=01, Module=01 (SANS préfixes REF/M)
        /// </summary>
        private ProjectProperties? ExtractPropertiesFromPath(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                    return null;

                // Normaliser le chemin
                path = path.TrimEnd('\\', '/');
                var parts = path.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
                
                // Chercher "Projects" dans le chemin et extraire les 3 niveaux suivants
                int projectsIndex = -1;
                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i].Equals("Projects", StringComparison.OrdinalIgnoreCase))
                    {
                        projectsIndex = i;
                        break;
                    }
                }
                
                if (projectsIndex >= 0 && projectsIndex + 3 < parts.Length)
                {
                    // Pattern: Projects\[PROJECT]\[REF]\[MODULE]
                    string projectNumber = parts[projectsIndex + 1];
                    string refFolder = parts[projectsIndex + 2];
                    string moduleFolder = parts[projectsIndex + 3];
                    
                    // Valider que ça ressemble à nos patterns
                    // Project: numérique (12345)
                    // Reference: REF suivi de numéros (REF01)
                    // Module: M suivi de numéros (M01)
                    if (Regex.IsMatch(projectNumber, @"^\d+$") &&
                        Regex.IsMatch(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase) &&
                        Regex.IsMatch(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase))
                    {
                        // Extraire les numéros SANS les préfixes REF et M
                        var refMatch = Regex.Match(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase);
                        var moduleMatch = Regex.Match(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase);
                        
                        return new ProjectProperties
                        {
                            ProjectNumber = projectNumber,
                            Reference = refMatch.Groups[1].Value,      // "01" au lieu de "REF01"
                            Module = moduleMatch.Groups[1].Value       // "01" au lieu de "M01"
                        };
                    }
                }
                
                // Alternative: essayer d'extraire depuis les 3 derniers dossiers du chemin
                if (parts.Length >= 3)
                {
                    string projectNumber = parts[parts.Length - 3];
                    string refFolder = parts[parts.Length - 2];
                    string moduleFolder = parts[parts.Length - 1];
                    
                    // Valider que ça ressemble à nos patterns
                    if (Regex.IsMatch(projectNumber, @"^\d+$") &&
                        Regex.IsMatch(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase) &&
                        Regex.IsMatch(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase))
                    {
                        // Extraire les numeros SANS les prefixes REF et M
                        var refMatch = Regex.Match(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase);
                        var moduleMatch = Regex.Match(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase);
                        
                        return new ProjectProperties
                        {
                            ProjectNumber = projectNumber,
                            Reference = refMatch.Groups[1].Value,      // "01" au lieu de "REF01"
                            Module = moduleMatch.Groups[1].Value       // "01" au lieu de "M01"
                        };
                    }
                }
                
                Services.Logger.Log($"[!] Impossible d'extraire les proprietes du chemin: {path}", Services.Logger.LogLevel.DEBUG);
                return null;
            }
            catch (Exception ex)
            {
                Services.Logger.Log($"[!] Erreur extraction proprietes: {ex.Message}", Services.Logger.LogLevel.DEBUG);
                return null;
            }
        }

        private void SelectAllInventor()
        {
            foreach (var it in InventorFiles) it.IsSelected = true;
        }

        private void DeselectAllInventor()
        {
            foreach (var it in InventorFiles) it.IsSelected = false;
        }

        private void SelectAllNonInventor()
        {
            foreach (var it in NonInventorFiles) it.IsSelected = true;
        }

        private void DeselectAllNonInventor()
        {
            foreach (var it in NonInventorFiles) it.IsSelected = false;
        }

        private string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            double kb = bytes / 1024.0;
            if (kb < 1024) return kb.ToString("0.0") + " KB";
            double mb = kb / 1024.0;
            if (mb < 1024) return mb.ToString("0.0") + " MB";
            double gb = mb / 1024.0;
            return gb.ToString("0.00") + " GB";
        }
        
        // Filtres de recherche avec support wildcards
        private void ApplySearchFilterInventor()
        {
            var temp = new List<FileItem>();
            
            if (string.IsNullOrWhiteSpace(SearchFilterInventor))
            {
                temp.AddRange(_allFiles.Where(f => f.IsInventorFile));
            }
            else
            {
                var filter = SearchFilterInventor.Trim();
                temp.AddRange(_allFiles.Where(f => f.IsInventorFile && MatchesFilter(f, filter)));
            }
            
            InventorFiles.Clear();
            foreach (var file in temp)
            {
                InventorFiles.Add(file);
            }
        }

        private void ApplySearchFilterNonInventor()
        {
            var temp = new List<FileItem>();
            
            if (string.IsNullOrWhiteSpace(SearchFilterNonInventor))
            {
                temp.AddRange(_allFiles.Where(f => !f.IsInventorFile));
            }
            else
            {
                var filter = SearchFilterNonInventor.Trim();
                temp.AddRange(_allFiles.Where(f => !f.IsInventorFile && MatchesFilter(f, filter)));
            }
            
            NonInventorFiles.Clear();
            foreach (var file in temp)
            {
                NonInventorFiles.Add(file);
            }
        }

        /// <summary>
        /// Filtre avancé avec support wildcards
        /// Exemples: "*.pdf", "Panel*", "*report*", "drawing"
        /// </summary>
        private bool MatchesFilter(FileItem file, string filter)
        {
            var filterLower = filter.ToLowerInvariant();
            var fileNameLower = file.FileName.ToLowerInvariant();
            var pathLower = file.FullPath.ToLowerInvariant();
            
            // Cas 1: Filtre par extension "*.ext"
            if (filterLower.StartsWith("*."))
            {
                var extension = filterLower.Substring(1);
                if (extension == ".*") return true;
                return fileNameLower.EndsWith(extension);
            }
            
            // Cas 2: Wildcard à la fin "nom*"
            if (filterLower.EndsWith("*"))
            {
                var prefix = filterLower.Substring(0, filterLower.Length - 1);
                return fileNameLower.StartsWith(prefix) || pathLower.Contains(prefix);
            }
            
            // Cas 3: Wildcard au début "*nom"
            if (filterLower.StartsWith("*"))
            {
                var suffix = filterLower.Substring(1);
                return fileNameLower.EndsWith(suffix) || pathLower.Contains(suffix);
            }
            
            // Cas 4: Wildcards des deux côtés "*nom*"
            if (filterLower.Contains("*"))
            {
                var parts = filterLower.Split(new[] { '*' }, StringSplitOptions.RemoveEmptyEntries);
                return parts.All(part => fileNameLower.Contains(part) || pathLower.Contains(part));
            }
            
            // Cas 5: Recherche simple (contient le texte)
            return fileNameLower.Contains(filterLower) || pathLower.Contains(filterLower);
        }
        
        // [+] Connexion/Deconnexion REELLE a Vault via VaultSDK
        private void ToggleConnection()
        {
            try
            {
                if (IsConnected)
                {
                    // Deconnexion
                    _vaultService.Disconnect();
                    IsConnected = false;
                    StatusMessage = "[>] Deconnecte de Vault";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                // Validation des champs
                if (string.IsNullOrWhiteSpace(Configuration.VaultConfig.ServerName))
                {
                    StatusMessage = "[!] Veuillez saisir le nom du serveur Vault";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                if (string.IsNullOrWhiteSpace(Configuration.VaultConfig.VaultName))
                {
                    StatusMessage = "[!] Veuillez saisir le nom du Vault";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                if (string.IsNullOrWhiteSpace(Configuration.VaultConfig.Username))
                {
                    StatusMessage = "[!] Veuillez saisir le nom d'utilisateur";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                // [+] Mot de passe obligatoire SAUF pour Administrator
                bool isAdministrator = Configuration.VaultConfig.Username.Equals("Administrator", StringComparison.OrdinalIgnoreCase);
                if (string.IsNullOrWhiteSpace(VaultPassword) && !isAdministrator)
                {
                    StatusMessage = "[!] Veuillez saisir le mot de passe";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                StatusMessage = "[>] Connexion a Vault en cours...";
                OnPropertyChanged(nameof(StatusMessage));

                // Connexion reelle via VaultSDK
                // [+] Toujours utiliser une chaine (jamais null)
                // Pour Administrator, si le champ est vide, envoyer "" explicitement
                string passwordToUse = VaultPassword ?? "";
                
                bool success = _vaultService.Connect(
                    Configuration.VaultConfig.ServerName,
                    Configuration.VaultConfig.VaultName,
                    Configuration.VaultConfig.Username,
                    passwordToUse
                );

                if (success)
                {
                    IsConnected = true;
                    StatusMessage = $"[+] Connecte: {Configuration.VaultConfig.Username}@{Configuration.VaultConfig.ServerName}/{Configuration.VaultConfig.VaultName}";
                    
                    // Charger les categories disponibles
                    LoadCategories();
                    
                    // Charger les Lifecycle Definitions disponibles
                    LoadLifecycleDefinitions();
                    
                    // NOTE: Les revisions sont gerees automatiquement par Vault via les transitions d'etat
                    
                    // Sauvegarder la configuration (avec mot de passe si SaveCredentials est coche)
                    SaveConfiguration();
                }
                else
                {
                    IsConnected = false;
                    StatusMessage = "[-] Echec de connexion a Vault. Verifiez le serveur, vault, identifiants et consultez les logs.";
                }

                OnPropertyChanged(nameof(StatusMessage));
            }
            catch (Exception ex)
            {
                IsConnected = false;
                StatusMessage = $"[-] Erreur connexion: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }
        
        // Browse folder avec OpenFileDialog moderne (grande fenêtre style Windows)
        private void BrowseFolder()
        {
            try
            {
                // Utiliser OpenFileDialog avec option de sélection de dossier via un fichier virtuel
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Sélectionner un dossier (naviguez vers le dossier puis cliquez Ouvrir)",
                    Filter = "Dossiers|*.folder|Tous les fichiers|*.*",
                    FileName = "Sélectionner ce dossier",
                    CheckFileExists = false,
                    CheckPathExists = true,
                    InitialDirectory = string.IsNullOrEmpty(ProjectPath) ? @"C:\Vault\Engineering\Projects" : ProjectPath
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    // Extraire le dossier du chemin
                    string selectedPath = System.IO.Path.GetDirectoryName(openFileDialog.FileName) ?? openFileDialog.FileName;
                    
                    if (Directory.Exists(selectedPath))
                    {
                        ProjectPath = selectedPath;
                        OnPropertyChanged(nameof(ProjectPath));
                        
                        // Scanner automatiquement le dossier selectionne
                        ScanProject();
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"[-] Erreur lors de la selection du dossier: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }
        
        /// <summary>
        /// Recupere le chemin du document actif dans Inventor
        /// Utilise P/Invoke direct vers oleaut32.dll (solution recommandee pour .NET Framework 4.8)
        /// </summary>
        private void GetPathFromInventor()
        {
            Services.Logger.Log("[>] [GetPathFromInventor] METHODE APPELEE", Services.Logger.LogLevel.INFO);
            
            try
            {
                Services.Logger.Log("[>] [GetPathFromInventor] Debut de la detection Inventor...", Services.Logger.LogLevel.INFO);
                StatusMessage = "[>] Recherche d'Inventor en cours...";
                OnPropertyChanged(nameof(StatusMessage));
                
                // Verifier si Inventor est lance
                var processes = System.Diagnostics.Process.GetProcessesByName("Inventor");
                if (processes.Length == 0)
                {
                    Services.Logger.Log("[-] [GetPathFromInventor] Aucun processus Inventor trouve", Services.Logger.LogLevel.WARNING);
                    StatusMessage = "[-] Inventor n'est pas lance. Veuillez ouvrir Inventor.";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }
                Services.Logger.Log($"[+] [GetPathFromInventor] {processes.Length} processus Inventor trouve(s)", Services.Logger.LogLevel.DEBUG);
                
                // Chercher un processus avec une fenetre visible
                bool hasVisibleWindow = false;
                foreach (var proc in processes)
                {
                    try
                    {
                        if (proc.MainWindowHandle != IntPtr.Zero && !string.IsNullOrEmpty(proc.MainWindowTitle))
                        {
                            hasVisibleWindow = true;
                            Services.Logger.Log($"[+] [GetPathFromInventor] Fenetre trouvee: '{proc.MainWindowTitle}' (PID: {proc.Id})", Services.Logger.LogLevel.DEBUG);
                            break;
                        }
                    }
                    catch { }
                }
                
                if (!hasVisibleWindow)
                {
                    Services.Logger.Log("[!] [GetPathFromInventor] Aucune fenetre Inventor visible, tentative COM quand meme...", Services.Logger.LogLevel.DEBUG);
                }
                
                // Utiliser P/Invoke pour recuperer l'objet COM
                Inventor.Application? inventorApp = null;
                
                try
                {
                    Services.Logger.Log("[>] [GetPathFromInventor] CLSIDFromProgID('Inventor.Application')...", Services.Logger.LogLevel.DEBUG);
                    int hr = CLSIDFromProgID("Inventor.Application", out Guid clsid);
                    if (hr != 0)
                    {
                        Services.Logger.Log($"[-] [GetPathFromInventor] CLSIDFromProgID echoyee: 0x{hr:X8}", Services.Logger.LogLevel.ERROR);
                        StatusMessage = "[-] Inventor non installe correctement";
                        OnPropertyChanged(nameof(StatusMessage));
                        return;
                    }
                    Services.Logger.Log($"[+] [GetPathFromInventor] CLSID: {clsid}", Services.Logger.LogLevel.DEBUG);
                    
                    Services.Logger.Log("[>] [GetPathFromInventor] GetActiveObject via oleaut32.dll...", Services.Logger.LogLevel.DEBUG);
                    hr = OleGetActiveObject(ref clsid, IntPtr.Zero, out object? inventorObj);
                    
                    Services.Logger.Log($"[>] [GetPathFromInventor] GetActiveObject retourne: hr=0x{hr:X8}, obj={(inventorObj != null ? "OK" : "null")}", Services.Logger.LogLevel.DEBUG);
                    
                    if (hr != 0)
                    {
                        Services.Logger.Log($"[-] [GetPathFromInventor] GetActiveObject echoyee: 0x{hr:X8}", Services.Logger.LogLevel.WARNING);
                        
                        if (hr == unchecked((int)0x800401E3)) // MK_E_UNAVAILABLE
                        {
                            // Inventor n'est pas enregistre dans le ROT - peut arriver si Inventor a ete lance en mode admin ou autre
                            StatusMessage = "[-] Inventor non accessible. Essayez de redemarrer Inventor normalement (sans mode admin).";
                        }
                        else
                        {
                            StatusMessage = $"[-] Erreur COM: 0x{hr:X8}";
                        }
                        OnPropertyChanged(nameof(StatusMessage));
                        return;
                    }
                    
                    if (inventorObj == null)
                    {
                        Services.Logger.Log("[-] [GetPathFromInventor] inventorObj est null", Services.Logger.LogLevel.WARNING);
                        StatusMessage = "[-] Impossible de se connecter a Inventor";
                        OnPropertyChanged(nameof(StatusMessage));
                        return;
                    }
                    
                    inventorApp = (Inventor.Application)inventorObj;
                    Services.Logger.Log($"[+] [GetPathFromInventor] Connecte a Inventor {inventorApp.SoftwareVersion?.DisplayVersion ?? "?"}", Services.Logger.LogLevel.INFO);
                }
                catch (InvalidCastException castEx)
                {
                    Services.Logger.Log($"[-] [GetPathFromInventor] Cast echoyee: {castEx.Message}", Services.Logger.LogLevel.ERROR);
                    StatusMessage = "[-] Erreur de type COM Inventor";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }
                catch (Exception comEx)
                {
                    Services.Logger.Log($"[-] [GetPathFromInventor] Exception COM: {comEx.Message}", Services.Logger.LogLevel.ERROR);
                    StatusMessage = $"[-] Erreur connexion Inventor: {comEx.Message}";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }
                
                if (inventorApp == null)
                {
                    StatusMessage = "[-] Impossible de se connecter a Inventor";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }
                
                // Recuperer le document actif
                Services.Logger.Log("[>] [GetPathFromInventor] Recuperation du document actif...", Services.Logger.LogLevel.DEBUG);
                Inventor.Document? activeDoc = inventorApp.ActiveDocument;
                
                if (activeDoc == null)
                {
                    Services.Logger.Log("[-] [GetPathFromInventor] Aucun document actif dans Inventor", Services.Logger.LogLevel.WARNING);
                    StatusMessage = "[-] Aucun document ouvert dans Inventor";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }
                
                Services.Logger.Log($"[+] [GetPathFromInventor] Document actif: {activeDoc.DisplayName}", Services.Logger.LogLevel.DEBUG);
                
                // Recuperer le FullFileName du document
                string fullFileName = activeDoc.FullFileName;
                Services.Logger.Log($"[>] [GetPathFromInventor] FullFileName: {fullFileName ?? "(null)"}", Services.Logger.LogLevel.DEBUG);
                
                if (string.IsNullOrEmpty(fullFileName))
                {
                    Services.Logger.Log("[-] [GetPathFromInventor] Document non enregistre (pas de chemin)", Services.Logger.LogLevel.WARNING);
                    StatusMessage = "[-] Le document Inventor n'a pas de chemin (document non enregistre)";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }
                
                // Extraire le dossier du fichier
                string? folderPath = System.IO.Path.GetDirectoryName(fullFileName);
                Services.Logger.Log($"[>] [GetPathFromInventor] Dossier extrait: {folderPath}", Services.Logger.LogLevel.DEBUG);
                
                if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                {
                    ProjectPath = folderPath;
                    OnPropertyChanged(nameof(ProjectPath));
                    
                    Services.Logger.Log($"[+] [GetPathFromInventor] Chemin defini: {folderPath}", Services.Logger.LogLevel.INFO);
                    StatusMessage = $"[+] Chemin recupere depuis Inventor: {System.IO.Path.GetFileName(fullFileName)}";
                    OnPropertyChanged(nameof(StatusMessage));
                    
                    // Scanner automatiquement le dossier
                    ScanProject();
                }
                else
                {
                    Services.Logger.Log($"[-] [GetPathFromInventor] Dossier non trouve: {folderPath}", Services.Logger.LogLevel.WARNING);
                    StatusMessage = $"[-] Dossier non trouve: {folderPath}";
                    OnPropertyChanged(nameof(StatusMessage));
                }
            }
            catch (Exception ex)
            {
                Services.Logger.Log($"[-] [GetPathFromInventor] Exception: {ex.Message}", Services.Logger.LogLevel.ERROR);
                StatusMessage = $"[-] Erreur lors de la detection Inventor: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
                Services.Logger.LogException("GetPathFromInventor", ex, Services.Logger.LogLevel.WARNING);
            }
        }
        
        // ═══════════════════════════════════════════════════════════════════════════════
        // P/Invoke pour accéder à COM - Solution .NET Framework 4.8
        // Source: https://docs.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.getactiveobject
        // Note: Marshal.GetActiveObject n'est pas disponible en .NET Core/5+, d'où P/Invoke
        // ═══════════════════════════════════════════════════════════════════════════════
        
        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, 
            out Guid pclsid);
        
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(
            ref Guid rclsid, 
            IntPtr pvReserved, 
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
        
        // Version avec code retour pour diagnostiquer les erreurs
        [DllImport("oleaut32.dll", EntryPoint = "GetActiveObject")]
        private static extern int OleGetActiveObject(
            ref Guid rclsid, 
            IntPtr pvReserved, 
            [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);
        
        /// <summary>
        /// Upload des fichiers sélectionnés vers Vault avec propriétés (VERSION ASYNCHRONE)
        /// </summary>
        private async Task AutoCheckInAsync()
        {
            try
            {
                // Verifier la connexion
                if (!IsConnected)
                {
                    StatusMessage = "[-] Non connecte a Vault. Veuillez vous connecter avant d'uploader des fichiers.";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                // Recuperer UNIQUEMENT les fichiers selectionnes (IsSelected = true)
                var selectedFiles = InventorFiles.Where(f => f.IsSelected)
                    .Concat(NonInventorFiles.Where(f => f.IsSelected))
                    .ToList();

                if (selectedFiles.Count == 0)
                {
                    StatusMessage = "[!] Aucun fichier selectionne. Veuillez selectionner au moins un fichier a uploader.";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                // Verifier les proprietes du projet
                if (ProjectProperties == null || 
                    string.IsNullOrWhiteSpace(ProjectProperties.ProjectNumber) ||
                    string.IsNullOrWhiteSpace(ProjectProperties.Reference) ||
                    string.IsNullOrWhiteSpace(ProjectProperties.Module))
                {
                    StatusMessage = "[-] Proprietes manquantes (Project/Reference/Module). Verifiez la structure du dossier.";
                    OnPropertyChanged(nameof(StatusMessage));
                    return;
                }

                // Confirmation moved to info bar: proceed immediately but notify user via status bar (no modal)
                StatusMessage = $"[>] Upload demarre: {selectedFiles.Count} fichier(s).";
                OnPropertyChanged(nameof(StatusMessage));

                // [+] ACTIVER LA BARRE DE PROGRESSION IMMÉDIATEMENT
                IsCheckingIn = true;
                OnPropertyChanged(nameof(IsCheckingIn));
                OnPropertyChanged(nameof(IsProcessing));

                int successCount = 0;
                int failedCount = 0;
                int currentIndex = 0;
                int totalFiles = selectedFiles.Count;

                // Initialiser la barre de progression
                ProgressMaximum = totalFiles;
                ProgressValue = 0;
                OnPropertyChanged(nameof(ProgressMaximum));
                OnPropertyChanged(nameof(ProgressValue));

                StatusMessage = $"[>] Upload en cours de {totalFiles} fichier(s)...";
                OnPropertyChanged(nameof(StatusMessage));

                // [+] ATTENDRE UN PEU POUR QUE L'UI SE METTE À JOUR
                await Task.Delay(100);

                // Upload asynchrone avec Task.Run pour ne pas bloquer l'UI
                await Task.Run(() =>
                {
                    foreach (var fileItem in selectedFiles)
                    {
                        // Vérifier pause
                        while (IsPaused)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

                        currentIndex++;
                        
                        // [+] Dispatcher pour mettre à jour l'UI depuis le thread de fond
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            ProgressText = $"{currentIndex}/{totalFiles}";
                            ProgressValue = currentIndex;
                            OnPropertyChanged(nameof(ProgressText));
                            OnPropertyChanged(nameof(ProgressValue));

                            fileItem.Status = "[~] Upload...";
                            OnPropertyChanged();

                            StatusMessage = $"[>] Upload: {fileItem.FileName} ({currentIndex}/{totalFiles})";
                            OnPropertyChanged(nameof(StatusMessage));
                        });

                        try
                        {
                            // Déterminer le chemin Vault basé sur la structure du projet
                            string vaultPath = DetermineVaultPath(fileItem.FullPath);

                            // Déterminer la catégorie à utiliser selon le type de fichier
                            long? categoryId = null;
                            string? categoryName = null;
                            long? lifecycleDefinitionId = null;
                            long? lifecycleStateId = null;
                            // NOTE: Revision gérée automatiquement par Vault via transitions d'état (WIP → Released)
                            
                            if (fileItem.IsInventorFile && SelectedCategoryInventor != null && SelectedCategoryInventor.Id > 0)
                            {
                                categoryId = SelectedCategoryInventor.Id;
                                categoryName = SelectedCategoryInventor.Name;
                                
                                // Obtenir le Lifecycle Definition ID selon la catégorie
                                lifecycleDefinitionId = _vaultService.GetLifecycleDefinitionIdByCategory(categoryName);
                                
                                // Obtenir l'état sélectionné - LOGGING DIAGNOSTIC
                                Services.Logger.Log($"   [i] [DEBUG] SelectedStateInventor est {(SelectedStateInventor != null ? $"ID={SelectedStateInventor.Id}, Name='{SelectedStateInventor.Name}'" : "NULL")}", Services.Logger.LogLevel.DEBUG);
                                
                                if (SelectedStateInventor != null)
                                {
                                    lifecycleStateId = SelectedStateInventor.Id;
                                    Services.Logger.Log($"   [i] [DEBUG] Utilisation de l'état sélectionné: ID={lifecycleStateId}", Services.Logger.LogLevel.DEBUG);
                                }
                                else if (lifecycleDefinitionId.HasValue)
                                {
                                    // Si aucun état sélectionné, utiliser "Work in Progress" par défaut
                                    lifecycleStateId = _vaultService.GetWorkInProgressStateId(lifecycleDefinitionId.Value);
                                    Services.Logger.Log($"   [i] [DEBUG] Fallback vers Work in Progress: ID={lifecycleStateId}", Services.Logger.LogLevel.DEBUG);
                                }
                                
                                // NOTE: Revision gérée automatiquement par Vault via transitions d'état
                            }
                            else if (!fileItem.IsInventorFile && SelectedCategoryNonInventor != null && SelectedCategoryNonInventor.Id > 0)
                            {
                                categoryId = SelectedCategoryNonInventor.Id;
                                categoryName = SelectedCategoryNonInventor.Name;
                                
                                // Obtenir le Lifecycle Definition ID selon la catégorie
                                lifecycleDefinitionId = _vaultService.GetLifecycleDefinitionIdByCategory(categoryName);
                                
                                // Obtenir l'état sélectionné
                                if (SelectedStateNonInventor != null)
                                {
                                    lifecycleStateId = SelectedStateNonInventor.Id;
                                }
                                else if (lifecycleDefinitionId.HasValue)
                                {
                                    // Si aucun état sélectionné, utiliser "Work in Progress" par défaut
                                    lifecycleStateId = _vaultService.GetWorkInProgressStateId(lifecycleDefinitionId.Value);
                                }
                                
                                // NOTE: Revision gérée automatiquement par Vault via transitions d'état
                            }
                            
                            // Upload via VaultSDK avec catégorie (qui détermine le FileClassification)
                            bool success = _vaultService.UploadFile(
                                fileItem.FullPath,
                                vaultPath,
                                ProjectProperties.ProjectNumber,
                                ProjectProperties.Reference,
                                ProjectProperties.Module,
                                categoryId,
                                categoryName,
                                lifecycleDefinitionId,
                                lifecycleStateId,
                                null, // revision = null, gérée par Vault
                                Comment  // Passer le commentaire depuis le champ UI
                            );

                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (success)
                                {
                                    fileItem.Status = "[+] Uploade";
                                    successCount++;
                                    StatusMessage = $"[+] {fileItem.FileName} uploade avec succes ({currentIndex}/{totalFiles})";
                                }
                                else
                                {
                                    fileItem.Status = "[-] Echec";
                                    failedCount++;
                                    StatusMessage = $"[-] Echec upload {fileItem.FileName} ({currentIndex}/{totalFiles})";
                                }
                                
                                OnPropertyChanged();
                                OnPropertyChanged(nameof(StatusMessage));
                            });
                        }
                        catch (Exception ex)
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                fileItem.Status = $"[-] Erreur: {ex.Message}";
                                failedCount++;
                                StatusMessage = $"[-] Erreur {fileItem.FileName}: {ex.Message}";
                                OnPropertyChanged();
                                OnPropertyChanged(nameof(StatusMessage));
                            });
                        }
                    }
                });

                // ═══════════════════════════════════════════════════════════════════════════════
                // [+] PHASE 2: APPLIQUER LES PROPRIETES EN BATCH (pour fichiers non-Inventor)
                // ═══════════════════════════════════════════════════════════════════════════════
                if (_vaultService.PendingPropertyCount > 0)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        StatusMessage = $"[>] Application des proprietes ({_vaultService.PendingPropertyCount} fichiers)...";
                        OnPropertyChanged(nameof(StatusMessage));
                    });
                    
                    var (propSuccess, propFailed) = _vaultService.ApplyPendingPropertyUpdates(waitBeforeStart: 0);
                    
                    if (propFailed > 0)
                    {
                        Services.Logger.Log($"[!] {propFailed} fichier(s) n'ont pas pu recevoir les proprietes (Job Processor occupe)", Services.Logger.LogLevel.WARNING);
                    }
                }

                // Fin de l'upload
                IsCheckingIn = false;
                ProgressValue = 0;
                OnPropertyChanged(nameof(IsCheckingIn));
                OnPropertyChanged(nameof(IsProcessing));
                OnPropertyChanged(nameof(ProgressValue));

                StatusMessage = $"[+] Upload termine: {successCount} reussi(s), {failedCount} echec(s) sur {totalFiles} fichier(s)";
                OnPropertyChanged(nameof(StatusMessage));
            }
            catch (Exception ex)
            {
                IsCheckingIn = false;
                ProgressValue = 0;
                OnPropertyChanged(nameof(IsCheckingIn));
                OnPropertyChanged(nameof(IsProcessing));
                OnPropertyChanged(nameof(ProgressValue));

                StatusMessage = $"[-] Erreur upload: {ex.Message}";
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        /// <summary>
        /// Determine le chemin Vault base sur le chemin local du fichier
        /// </summary>
        private string DetermineVaultPath(string localFilePath)
        {
            try
            {
                // Structure attendue: C:\Vault\Engineering\Projects\[PROJECT]\[REF]\[MODULE]\...
                // Chemin Vault: $/Engineering/Projects/[PROJECT]/[REF]/[MODULE]/...

                string relativePath = localFilePath;
                
                // Trouver le dossier du module dans le chemin
                int projectsIndex = relativePath.IndexOf(@"\Projects\", StringComparison.OrdinalIgnoreCase);
                if (projectsIndex >= 0)
                {
                    // Extraire la partie après "Projects\"
                    relativePath = relativePath.Substring(projectsIndex + @"\Projects\".Length);
                    
                    // Construire le chemin Vault
                    string vaultPath = "$/Engineering/Projects/" + relativePath.Replace('\\', '/');
                    
                    // Retirer le nom du fichier pour ne garder que le dossier
                    int lastSlash = vaultPath.LastIndexOf('/');
                    if (lastSlash > 0)
                    {
                        vaultPath = vaultPath.Substring(0, lastSlash);
                    }
                    
                    return vaultPath;
                }
                
                // Par défaut, uploader dans le dossier du module
                if (ProjectProperties == null) return "$/Engineering";
                return $"$/Engineering/Projects/{ProjectProperties.ProjectNumber}/{ProjectProperties.Reference}/{ProjectProperties.Module}";
            }
            catch
            {
                // Fallback: dossier du module
                if (ProjectProperties == null) return "$/Engineering";
                return $"$/Engineering/Projects/{ProjectProperties.ProjectNumber}/{ProjectProperties.Reference}/{ProjectProperties.Module}";
            }
        }
        
        /// <summary>
        /// Charge la configuration depuis appsettings.json
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    
                    // Parse JSON simple (sans Newtonsoft.Json pour NET Framework 4.8)
                    var lines = json.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    
                    foreach (var line in lines)
                    {
                        var trimmed = line.Trim().TrimEnd(',');
                        
                        if (trimmed.Contains("\"serverName\""))
                        {
                            Configuration.VaultConfig.ServerName = ExtractJsonValue(trimmed);
                        }
                        else if (trimmed.Contains("\"vaultName\""))
                        {
                            Configuration.VaultConfig.VaultName = ExtractJsonValue(trimmed);
                        }
                        else if (trimmed.Contains("\"username\""))
                        {
                            Configuration.VaultConfig.Username = ExtractJsonValue(trimmed);
                        }
                        else if (trimmed.Contains("\"password\""))
                        {
                            string pwd = ExtractJsonValue(trimmed);
                            if (!string.IsNullOrEmpty(pwd))
                            {
                                VaultPassword = pwd;
                                SaveCredentials = true;
                            }
                        }
                        else if (trimmed.Contains("\"connectionTimeoutSeconds\""))
                        {
                            if (int.TryParse(ExtractJsonValue(trimmed), out int timeout))
                                Configuration.VaultConfig.ConnectionTimeoutSeconds = timeout;
                        }
                    }
                    
                    OnPropertyChanged(nameof(Configuration));
                    OnPropertyChanged(nameof(VaultPassword));
                    OnPropertyChanged(nameof(SaveCredentials));
                }
            }
            catch (Exception ex)
            {
                // Pas d'erreur si le fichier n'existe pas ou est mal formaté
                System.Diagnostics.Debug.WriteLine($"LoadConfiguration error: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Sauvegarde la configuration dans appsettings.json
        /// </summary>
        private void SaveConfiguration()
        {
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                
                string json = @"{
  ""vaultConfiguration"": {
    ""serverName"": """ + Configuration.VaultConfig.ServerName + @""",
    ""vaultName"": """ + Configuration.VaultConfig.VaultName + @""",
    ""username"": """ + Configuration.VaultConfig.Username + @""",
    ""password"": """ + (SaveCredentials ? VaultPassword : "") + @""",
    ""domain"": """",
    ""connectionTimeoutSeconds"": " + Configuration.VaultConfig.ConnectionTimeoutSeconds + @",
    ""retryAttempts"": " + Configuration.VaultConfig.RetryAttempts + @",
    ""retryDelayMs"": " + Configuration.VaultConfig.RetryDelayMs + @"
  }
}";
                
                File.WriteAllText(configPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SaveConfiguration error: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Extrait une valeur d'une ligne JSON simple (sans bibliothèque)
        /// </summary>
        private string ExtractJsonValue(string line)
        {
            // Format: "key": "value" ou "key": value
            int colonIndex = line.IndexOf(':');
            if (colonIndex < 0) return "";
            
            string valuePart = line.Substring(colonIndex + 1).Trim();
            
            // Enlever les guillemets si présents
            valuePart = valuePart.Trim('"', ',', ' ');
            
            return valuePart;
        }
  
        /// <summary>
        /// Arrête le traitement en cours
        /// </summary>
        private void StopProcessing()
        {
            // Stop the processing gracefully
            IsCheckingIn = false;
            IsAddingToVault = false;
        OnPropertyChanged(nameof(IsProcessing));
      StatusMessage = "⏹️ Traitement arrêté par l'utilisateur";
            OnPropertyChanged(nameof(StatusMessage));
        }
  
        /// <summary>
 /// Annule le traitement en cours
        /// </summary>
        private void CancelProcessing()
        {
            // Cancel the processing
            IsCheckingIn = false;
            IsAddingToVault = false;
            OnPropertyChanged(nameof(IsProcessing));
            StatusMessage = "[-] Traitement annule par l'utilisateur";
            OnPropertyChanged(nameof(StatusMessage));
        }

    }
    
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        public RelayCommand(Action execute) => _execute = execute;
        public event EventHandler CanExecuteChanged { add { } remove { } }
        public bool CanExecute(object parameter) => true;
        public void Execute(object parameter) => _execute();
    }
}