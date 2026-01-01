#nullable enable
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using XnrgyEngineeringAutomationTools.Models;
using XnrgyEngineeringAutomationTools.Modules.VaultUpload.Models;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Views;

namespace XnrgyEngineeringAutomationTools.Modules.VaultUpload.Views
{
    /// <summary>
    /// Upload Module vers Vault - Module integre dans XNRGY Engineering Automation Tools
    /// Auteur: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// Version: 1.0.0 - Decembre 2025
    /// </summary>
    public partial class VaultUploadModuleWindow : Window
    {
        // ====================================================================
        // Services et connexion
        // ====================================================================
        private readonly VaultSdkService? _vaultService;
        private bool _isVaultConnected = false;

        // ====================================================================
        // Collections de fichiers
        // ====================================================================
        public ObservableCollection<VaultUploadFileItem> AllFiles { get; } = new();
        private readonly List<VaultUploadFileItem> _allFilesMaster = new();

        // Collections legacy pour compatibilite (pointent vers AllFiles filtrees)
        public ObservableCollection<VaultUploadFileItem> InventorFiles => AllFiles;
        public ObservableCollection<VaultUploadFileItem> NonInventorFiles => AllFiles;

        // ====================================================================
        // Proprietes projet
        // ====================================================================
        private VaultProjectProperties? _projectProperties;
        private string _projectPath = string.Empty;

        // ====================================================================
        // Etat du traitement
        // ====================================================================
        private bool _isProcessing = false;
        private bool _isPaused = false;
        private CancellationTokenSource? _cancellationTokenSource;
        private int _uploadedCount = 0;
        private int _failedCount = 0;

        // ====================================================================
        // Extensions a exclure
        // ====================================================================
        private static readonly HashSet<string> ExcludedExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".v", ".bak", ".old", ".tmp", ".temp", ".ipj", ".lck", ".lock", ".log", ".dwl", ".dwl2"
        };

        private static readonly string[] ExcludedPrefixes = { "~$", "._", "Backup_", ".~" };
        private static readonly string[] ExcludedFolders = { "OldVersions", "oldversions", "Backup", "backup", ".vault", ".git", ".vs" };
        private static readonly HashSet<string> InventorExtensions = new(StringComparer.OrdinalIgnoreCase) { ".ipt", ".iam", ".idw", ".ipn" };

        // ====================================================================
        // Constructeur
        // ====================================================================
        public VaultUploadModuleWindow(VaultSdkService? vaultService)
        {
            InitializeComponent();
            _vaultService = vaultService;
            
            DgFiles.ItemsSource = AllFiles;
            
            // S'abonner aux changements de theme
            MainWindow.ThemeChanged += OnThemeChanged;
            this.Closed += (s, e) => MainWindow.ThemeChanged -= OnThemeChanged;
            
            // Appliquer le theme actuel au demarrage
            ApplyTheme(MainWindow.CurrentThemeIsDark);
        }

        /// <summary>
        /// Gestionnaire de changement de theme depuis MainWindow
        /// </summary>
        private void OnThemeChanged(bool isDarkTheme)
        {
            Dispatcher.Invoke(() => ApplyTheme(isDarkTheme));
        }

        /// <summary>
        /// Applique le theme a cette fenetre
        /// </summary>
        private void ApplyTheme(bool isDarkTheme)
        {
            if (isDarkTheme)
            {
                // Theme SOMBRE
                this.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46)); // #1E1E2E
            }
            else
            {
                // Theme CLAIR
                this.Background = new SolidColorBrush(Color.FromRgb(245, 247, 250)); // Bleu-gris tres clair
            }
        }

        // ====================================================================
        // Chargement fenetre
        // ====================================================================
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Log("[i] Module Upload vers Vault initialise", LogLevel.INFO);
            
            // Verifier connexion Vault
            if (_vaultService != null && _vaultService.IsConnected)
            {
                _isVaultConnected = true;
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(16, 124, 16)); // Vert
                RunVaultName.Text = $" Vault : {_vaultService.VaultName}  /  ";
                RunUserName.Text = $" Utilisateur : {_vaultService.UserName}  /  ";
                RunStatus.Text = " Statut : Connecte";
                Log($"[+] Connexion Vault active: {_vaultService.UserName}@{_vaultService.ServerName}/{_vaultService.VaultName}", LogLevel.SUCCESS);
                
                // Charger categories
                LoadCategories();
            }
            else
            {
                _isVaultConnected = false;
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(232, 17, 35)); // Rouge
                RunVaultName.Text = " Vault : --  /  ";
                RunUserName.Text = " Utilisateur : --  /  ";
                RunStatus.Text = " Statut : Deconnecte";
                Log("[!] Vault non connecte - Upload impossible", LogLevel.WARNING);
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (_isProcessing)
            {
                var result = XnrgyMessageBox.Show(
                    "Un upload est en cours. Voulez-vous vraiment annuler et fermer?",
                    "Upload en cours",
                    XnrgyMessageBoxType.Warning,
                    XnrgyMessageBoxButtons.YesNo,
                    this);

                if (result != XnrgyMessageBoxResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
                _cancellationTokenSource?.Cancel();
            }
        }

        // ====================================================================
        // Charger categories depuis Vault
        // ====================================================================
        private void LoadCategories()
        {
            try
            {
                if (_vaultService == null || !_vaultService.IsConnected) return;

                var categories = _vaultService.GetAvailableCategories();
                
                Dispatcher.Invoke(() =>
                {
                    CmbCategory.Items.Clear();

                    foreach (var cat in categories)
                    {
                        CmbCategory.Items.Add(new VaultCategoryItem { Id = cat.Id, Name = cat.Name });
                    }

                    // Selectionner "Engineering" par defaut
                    var engineering = CmbCategory.Items.Cast<VaultCategoryItem>()
                        .FirstOrDefault(c => c.Name.Equals("Engineering", StringComparison.OrdinalIgnoreCase));
                    
                    if (engineering != null)
                    {
                        CmbCategory.SelectedItem = engineering;
                    }
                    else if (CmbCategory.Items.Count > 0)
                    {
                        CmbCategory.SelectedIndex = 0;
                    }

                    Log($"[+] {categories.Count()} categories chargees", LogLevel.INFO);
                    
                    // Charger les Lifecycle States pour la categorie selectionnee
                    LoadLifecycleStatesForCategory(CmbCategory, CmbLifecycleState);
                });
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur chargement categories: {ex.Message}", LogLevel.ERROR);
            }
        }

        // ====================================================================
        // Changement de categorie - Charger les Lifecycle States (unifie)
        // ====================================================================
        private void CmbCategoryInventor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadLifecycleStatesForCategory(CmbCategory, CmbLifecycleState);
        }

        private void CmbCategoryNonInventor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadLifecycleStatesForCategory(CmbCategory, CmbLifecycleState);
        }

        private void LoadLifecycleStatesForCategory(ComboBox categoryCombo, ComboBox stateCombo)
        {
            try
            {
                if (_vaultService == null || !_vaultService.IsConnected) return;

                var selectedCategory = categoryCombo.SelectedItem as VaultCategoryItem;
                if (selectedCategory == null)
                {
                    stateCombo.Items.Clear();
                    return;
                }

                // Obtenir le Lifecycle Definition ID pour cette categorie
                var lifecycleDefId = _vaultService.GetLifecycleDefinitionIdByCategory(selectedCategory.Name);
                
                if (!lifecycleDefId.HasValue)
                {
                    stateCombo.Items.Clear();
                    Log($"[i] Pas de Lifecycle pour la categorie '{selectedCategory.Name}'", LogLevel.INFO);
                    return;
                }

                // Obtenir les states disponibles
                var lifecycleDefs = _vaultService.GetAvailableLifecycleDefinitions();
                var lifecycleDef = lifecycleDefs?.FirstOrDefault(d => d.Id == lifecycleDefId.Value);

                if (lifecycleDef == null || lifecycleDef.States == null || !lifecycleDef.States.Any())
                {
                    stateCombo.Items.Clear();
                    return;
                }

                Dispatcher.Invoke(() =>
                {
                    stateCombo.Items.Clear();
                    
                    // Filtrer les states selon la categorie
                    var allowedStates = GetAllowedStatesForCategory(selectedCategory.Name.ToLowerInvariant());
                    
                    foreach (var state in lifecycleDef.States)
                    {
                        if (allowedStates.Count == 0 || allowedStates.Contains(state.Name.ToLowerInvariant()))
                        {
                            stateCombo.Items.Add(new VaultLifecycleStateItem { Id = state.Id, Name = state.Name });
                        }
                    }

                    // Selectionner "Work in Progress" par defaut
                    var wip = stateCombo.Items.Cast<VaultLifecycleStateItem>()
                        .FirstOrDefault(s => s.Name.ToLowerInvariant().Contains("work in progress"));
                    
                    if (wip != null)
                    {
                        stateCombo.SelectedItem = wip;
                    }
                    else if (stateCombo.Items.Count > 0)
                    {
                        stateCombo.SelectedIndex = 0;
                    }

                    Log($"[+] {stateCombo.Items.Count} states charges pour '{selectedCategory.Name}'", LogLevel.INFO);
                });
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur chargement states: {ex.Message}", LogLevel.ERROR);
            }
        }

        private HashSet<string> GetAllowedStatesForCategory(string categoryLower)
        {
            // Filtrer les States selon ce qui est vraiment disponible dans Vault Client
            return categoryLower switch
            {
                "engineering" => new HashSet<string>(StringComparer.OrdinalIgnoreCase) 
                    { "for review", "work in progress", "released", "obsolete" },
                "office" => new HashSet<string>(StringComparer.OrdinalIgnoreCase) 
                    { "work in progress", "released", "obsolete" },
                "design representation" => new HashSet<string>(StringComparer.OrdinalIgnoreCase) 
                    { "released", "work in progress", "obsolete" },
                _ => new HashSet<string>() // Tous les states pour les autres categories
            };
        }

        // ====================================================================
        // Selection Module
        // ====================================================================
        private void SelectModule_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("[?] Scan des modules disponibles...", LogLevel.INFO);
                const string basePath = @"C:\Vault\Engineering\Projects";

                if (!Directory.Exists(basePath))
                {
                    Log($"[-] Dossier Vault introuvable: {basePath}", LogLevel.ERROR);
                    XnrgyMessageBox.ShowError($"Dossier Vault introuvable:\n{basePath}", "Erreur", this);
                    return;
                }

                var modules = ScanAvailableModules(basePath);

                if (modules.Count == 0)
                {
                    Log("[!] Aucun module trouve", LogLevel.WARNING);
                    XnrgyMessageBox.ShowInfo($"Aucun module trouve dans:\n{basePath}", "Information", this);
                    return;
                }

                Log($"[+] {modules.Count} modules trouves", LogLevel.SUCCESS);

                // Creer fenetre de selection
                var window = new ModuleSelectionWindow(modules.Select(m => new ModuleInfo
                {
                    FullPath = m.FullPath,
                    DisplayName = m.DisplayName,
                    ProjectNumber = m.ProjectNumber,
                    Reference = m.Reference,
                    Module = m.Module
                }).ToList());
                window.Owner = this;

                if (window.ShowDialog() == true && window.SelectedModule != null)
                {
                    var selected = window.SelectedModule;
                    _projectPath = selected.FullPath;
                    TxtProjectPath.Text = _projectPath;
                    _projectProperties = new VaultProjectProperties
                    {
                        ProjectNumber = selected.ProjectNumber,
                        Reference = selected.Reference,
                        Module = selected.Module
                    };
                    UpdatePropertiesDisplay();
                    Log($"[+] Module selectionne: {selected.DisplayName}", LogLevel.SUCCESS);
                    
                    // Scanner automatiquement
                    ScanProject();
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur selection module: {ex.Message}", LogLevel.ERROR);
            }
        }

        private List<VaultModuleInfo> ScanAvailableModules(string basePath)
        {
            var modules = new List<VaultModuleInfo>();

            try
            {
                var projectDirs = Directory.GetDirectories(basePath)
                    .Where(d => Regex.IsMatch(Path.GetFileName(d), @"^\d+$"))
                    .OrderBy(d => d);

                foreach (var projectDir in projectDirs)
                {
                    string projectNum = Path.GetFileName(projectDir);
                    var refDirs = Directory.GetDirectories(projectDir)
                        .Where(d => Regex.IsMatch(Path.GetFileName(d), @"^REF\d+$", RegexOptions.IgnoreCase))
                        .OrderBy(d => d);

                    foreach (var refDir in refDirs)
                    {
                        string refFull = Path.GetFileName(refDir);
                        string refNum = refFull.Substring(3);
                        var moduleDirs = Directory.GetDirectories(refDir)
                            .Where(d => Regex.IsMatch(Path.GetFileName(d), @"^M\d+$", RegexOptions.IgnoreCase))
                            .OrderBy(d => d);

                        foreach (var moduleDir in moduleDirs)
                        {
                            string moduleFull = Path.GetFileName(moduleDir);
                            string moduleNum = moduleFull.Substring(1);
                            int fileCount = 0;
                            try { fileCount = Directory.GetFiles(moduleDir, "*.*", SearchOption.AllDirectories).Length; }
                            catch { }

                            if (fileCount > 0)
                            {
                                modules.Add(new VaultModuleInfo
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
                Log($"[-] Erreur scan: {ex.Message}", LogLevel.ERROR);
            }

            return modules;
        }

        // ====================================================================
        // Parcourir dossier
        // ====================================================================
        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Selectionner un dossier (naviguez vers le dossier puis cliquez Ouvrir)",
                    Filter = "Dossiers|*.folder|Tous les fichiers|*.*",
                    FileName = "Selectionner ce dossier",
                    CheckFileExists = false,
                    CheckPathExists = true,
                    InitialDirectory = string.IsNullOrEmpty(_projectPath) ? @"C:\Vault\Engineering\Projects" : _projectPath
                };

                if (dialog.ShowDialog() == true)
                {
                    string selectedPath = Path.GetDirectoryName(dialog.FileName) ?? dialog.FileName;
                    if (Directory.Exists(selectedPath))
                    {
                        _projectPath = selectedPath;
                        TxtProjectPath.Text = _projectPath;
                        _projectProperties = ExtractPropertiesFromPath(_projectPath);
                        UpdatePropertiesDisplay();
                        ScanProject();
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur parcourir: {ex.Message}", LogLevel.ERROR);
            }
        }

        // ====================================================================
        // Depuis Inventor
        // ====================================================================
        private void GetFromInventor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("[?] Recherche Inventor en cours...", LogLevel.INFO);

                var processes = System.Diagnostics.Process.GetProcessesByName("Inventor");
                if (processes.Length == 0)
                {
                    Log("[-] Inventor n'est pas lance", LogLevel.ERROR);
                    XnrgyMessageBox.ShowInfo("Inventor n'est pas lance.\nVeuillez ouvrir Inventor avec un document.", "Information", this);
                    return;
                }

                // Tenter de recuperer via COM
                try
                {
                    var inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType != null)
                    {
                        dynamic? inventorApp = Marshal.GetActiveObject("Inventor.Application");
                        if (inventorApp != null)
                        {
                            dynamic? activeDoc = inventorApp.ActiveDocument;
                            if (activeDoc != null)
                            {
                                string fullPath = activeDoc.FullFileName;
                                if (!string.IsNullOrEmpty(fullPath))
                                {
                                    string? folder = Path.GetDirectoryName(fullPath);
                                    if (!string.IsNullOrEmpty(folder))
                                    {
                                        _projectPath = folder;
                                        TxtProjectPath.Text = _projectPath;
                                        _projectProperties = ExtractPropertiesFromPath(_projectPath);
                                        UpdatePropertiesDisplay();
                                        Log($"[+] Chemin Inventor: {_projectPath}", LogLevel.SUCCESS);
                                        ScanProject();
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (COMException)
                {
                    Log("[!] Impossible de communiquer avec Inventor via COM", LogLevel.WARNING);
                }

                XnrgyMessageBox.ShowInfo("Aucun document actif dans Inventor.\nOuvrez un fichier et reessayez.", "Information", this);
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur Inventor: {ex.Message}", LogLevel.ERROR);
            }
        }

        // ====================================================================
        // Scanner projet
        // ====================================================================
        private void ScanProject_Click(object sender, RoutedEventArgs e)
        {
            _projectPath = TxtProjectPath.Text;
            _projectProperties = ExtractPropertiesFromPath(_projectPath);
            UpdatePropertiesDisplay();
            ScanProject();
        }

        private void ScanProject()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_projectPath) || !Directory.Exists(_projectPath))
                {
                    Log($"[-] Dossier introuvable: {_projectPath}", LogLevel.ERROR);
                    return;
                }

                Log($"[?] Scan du dossier: {_projectPath}", LogLevel.INFO);

                AllFiles.Clear();
                _allFilesMaster.Clear();

                var allFilesOnDisk = Directory.GetFiles(_projectPath, "*.*", SearchOption.AllDirectories);
                int excludedCount = 0;

                foreach (var f in allFilesOnDisk.OrderBy(p => Path.GetFileName(p), StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var fi = new FileInfo(f);

                        // Exclure fichiers backup/temporaires
                        if (ExcludedExtensions.Contains(fi.Extension)) { excludedCount++; continue; }
                        if (ExcludedPrefixes.Any(p => fi.Name.StartsWith(p, StringComparison.OrdinalIgnoreCase))) { excludedCount++; continue; }
                        var dirName = fi.Directory?.Name ?? "";
                        if (ExcludedFolders.Any(ex => dirName.Equals(ex, StringComparison.OrdinalIgnoreCase))) { excludedCount++; continue; }

                        // Calculer chemin relatif
                        var relativePath = f.StartsWith(_projectPath) 
                            ? f.Substring(_projectPath.Length).TrimStart('\\') 
                            : fi.Directory?.Name ?? "";

                        var item = new VaultUploadFileItem
                        {
                            FileName = fi.Name,
                            FullPath = fi.FullName,
                            RelativePath = relativePath,
                            FileType = fi.Extension,
                            FileExtension = fi.Extension,
                            FileSizeFormatted = FormatSize(fi.Length),
                            IsInventorFile = InventorExtensions.Contains(fi.Extension),
                            IsSelected = true,
                            Status = "En attente"
                        };

                        _allFilesMaster.Add(item);
                        AllFiles.Add(item);
                    }
                    catch { }
                }

                UpdateStatistics();
                
                // Peupler le filtre d'extension dynamiquement
                PopulateExtensionFilter();
                
                int inventorCount = _allFilesMaster.Count(f => f.IsInventorFile);
                int otherCount = _allFilesMaster.Count - inventorCount;
                var statusMsg = $"[+] Scanne: {_allFilesMaster.Count} fichiers ({inventorCount} inventor, {otherCount} autres)";
                if (excludedCount > 0) statusMsg += $" | {excludedCount} exclus";
                Log(statusMsg, LogLevel.SUCCESS);
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur scan: {ex.Message}", LogLevel.ERROR);
            }
        }

        // ====================================================================
        // Selection fichiers (unifie)
        // ====================================================================
        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in AllFiles) f.IsSelected = true;
            UpdateStatistics();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in AllFiles) f.IsSelected = false;
            UpdateStatistics();
        }

        // Legacy handlers pour compatibilite
        private void SelectAllInventor_Click(object sender, RoutedEventArgs e) => SelectAll_Click(sender, e);
        private void DeselectAllInventor_Click(object sender, RoutedEventArgs e) => DeselectAll_Click(sender, e);
        private void SelectAllNonInventor_Click(object sender, RoutedEventArgs e) => SelectAll_Click(sender, e);
        private void DeselectAllNonInventor_Click(object sender, RoutedEventArgs e) => DeselectAll_Click(sender, e);

        // ====================================================================
        // Recherche unifiee
        // ====================================================================
        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyAllFilters();
        }

        // Legacy handlers
        private void SearchInventor_TextChanged(object sender, TextChangedEventArgs e) => ApplyAllFilters();
        private void SearchNonInventor_TextChanged(object sender, TextChangedEventArgs e) => ApplyAllFilters();

        // ====================================================================
        // Filtres Extension et Etat (unifie)
        // ====================================================================
        private void CmbExtension_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyAllFilters();
        }

        private void CmbState_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyAllFilters();
        }

        private void CmbCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Charger les Lifecycle States pour la categorie selectionnee
            LoadLifecycleStatesForCategory(CmbCategory, CmbLifecycleState);
        }

        // Legacy handlers pour compatibilite
        private void CmbExtensionInventor_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyAllFilters();
        private void CmbExtensionNonInventor_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyAllFilters();
        private void CmbStateInventor_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyAllFilters();
        private void CmbStateNonInventor_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyAllFilters();

        /// <summary>
        /// Peuple le filtre d'extension dynamiquement avec les extensions presentes dans les fichiers charges
        /// </summary>
        private void PopulateExtensionFilter()
        {
            if (_allFilesMaster == null || _allFilesMaster.Count == 0 || CmbExtension == null)
                return;

            // Extraire les extensions uniques et les trier
            var uniqueExtensions = _allFilesMaster
                .Select(f => f.FileExtension?.ToLowerInvariant() ?? "")
                .Where(ext => !string.IsNullOrWhiteSpace(ext))
                .Distinct()
                .OrderBy(ext => ext)
                .ToList();

            // Garder "Tous" et ajouter les extensions dynamiques
            CmbExtension.Items.Clear();
            CmbExtension.Items.Add(new ComboBoxItem { Content = "Tous", IsSelected = true });

            foreach (var ext in uniqueExtensions)
            {
                CmbExtension.Items.Add(new ComboBoxItem { Content = ext });
            }

            CmbExtension.SelectedIndex = 0;
            Log($"[i] Filtre extension: {uniqueExtensions.Count} extensions detectees", LogLevel.INFO);
        }

        /// <summary>
        /// Applique tous les filtres au DataGrid unifie
        /// </summary>
        private void ApplyAllFilters()
        {
            if (_allFilesMaster == null || _allFilesMaster.Count == 0) return;

            var source = _allFilesMaster.AsEnumerable();

            // Filtre recherche
            var searchText = TxtSearch?.Text?.ToLowerInvariant() ?? "";
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                source = source.Where(f => 
                    f.FileName.ToLowerInvariant().Contains(searchText) ||
                    f.FullPath.ToLowerInvariant().Contains(searchText));
            }

            // Filtre extension
            var extItem = CmbExtension?.SelectedItem as ComboBoxItem;
            var extFilter = extItem?.Content?.ToString() ?? "Tous";
            if (extFilter != "Tous" && !string.IsNullOrEmpty(extFilter))
            {
                source = source.Where(f => f.FileExtension?.ToLowerInvariant() == extFilter.ToLowerInvariant());
            }

            // Filtre etat
            var stateItem = CmbState?.SelectedItem as ComboBoxItem;
            var stateFilter = stateItem?.Content?.ToString() ?? "Tous";
            switch (stateFilter)
            {
                case "En attente":
                    source = source.Where(f => f.Status == "En attente");
                    break;
                case "Uploade":
                    source = source.Where(f => f.Status?.Contains("Uploade") == true || f.Status?.Contains("[+]") == true);
                    break;
                case "Ignore":
                    source = source.Where(f => f.Status?.Contains("Ignore") == true || f.Status?.Contains("Existe") == true);
                    break;
                case "Erreur":
                    source = source.Where(f => f.Status?.Contains("Erreur") == true || f.Status?.Contains("[-]") == true);
                    break;
            }

            // Appliquer
            AllFiles.Clear();
            foreach (var f in source)
            {
                AllFiles.Add(f);
            }
        }

        // ====================================================================
        // Upload vers Vault
        // ====================================================================
        private async void CheckIn_Click(object sender, RoutedEventArgs e)
        {
            if (!_isVaultConnected || _vaultService == null)
            {
                XnrgyMessageBox.ShowError("Vault non connecte.\nVeuillez vous connecter depuis la fenetre principale.", "Erreur", this);
                return;
            }

            var selectedFiles = _allFilesMaster.Where(f => f.IsSelected).ToList();
            if (selectedFiles.Count == 0)
            {
                XnrgyMessageBox.ShowInfo("Aucun fichier selectionne pour l'upload.", "Information", this);
                return;
            }

            var confirm = XnrgyMessageBox.Show(
                $"Vous allez uploader {selectedFiles.Count} fichiers vers Vault.\n\n" +
                $"Projet: {_projectProperties?.ProjectNumber ?? "N/A"}\n" +
                $"Reference: {_projectProperties?.Reference ?? "N/A"}\n" +
                $"Module: {_projectProperties?.Module ?? "N/A"}\n\n" +
                "Continuer?",
                "Confirmation Upload",
                XnrgyMessageBoxType.Info,
                XnrgyMessageBoxButtons.YesNo,
                this);

            if (confirm != XnrgyMessageBoxResult.Yes) return;

            await StartUploadAsync(selectedFiles);
        }

        private async Task StartUploadAsync(List<VaultUploadFileItem> files)
        {
            _isProcessing = true;
            _isPaused = false;
            _uploadedCount = 0;
            _failedCount = 0;
            _cancellationTokenSource = new CancellationTokenSource();

            // Capturer le commentaire AVANT le Task.Run (sur le thread UI)
            string baseComment = TxtComment.Text;
            string uploadComment = baseComment;
            if (_projectProperties != null)
            {
                uploadComment = $"{baseComment} | Project: {_projectProperties.ProjectNumber}, Ref: {_projectProperties.Reference}, Module: {_projectProperties.Module}";
            }

            SetProcessingState(true);

            Log($"[>] Debut upload de {files.Count} fichiers...", LogLevel.INFO);
            Log($"[i] Commentaire: {uploadComment}", LogLevel.INFO);

            try
            {
                int total = files.Count;
                int current = 0;

                foreach (var file in files)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        Log("[!] Upload annule par l'utilisateur", LogLevel.WARNING);
                        break;
                    }

                    // Pause
                    while (_isPaused && !_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        await Task.Delay(100);
                    }

                    current++;
                    UpdateProgress(current, total, file.FileName);

                    try
                    {
                        file.Status = "Upload en cours...";
                        
                        // Upload vers Vault - passer le commentaire en parametre
                        bool success = await Task.Run(() => UploadFileToVault(file, uploadComment));

                        if (success)
                        {
                            file.Status = "[+] Uploade";
                            _uploadedCount++;
                            Log($"[+] {file.FileName}", LogLevel.SUCCESS);
                        }
                        else
                        {
                            file.Status = "[-] Echec";
                            _failedCount++;
                            Log($"[-] Echec: {file.FileName}", LogLevel.ERROR);
                        }
                    }
                    catch (Exception ex)
                    {
                        file.Status = $"[-] Erreur: {ex.Message}";
                        _failedCount++;
                        Log($"[-] Erreur {file.FileName}: {ex.Message}", LogLevel.ERROR);
                    }

                    UpdateStatistics();
                }

                // Resultat final
                if (_uploadedCount > 0)
                {
                    Log($"[+] Upload termine: {_uploadedCount} fichiers uploades, {_failedCount} echecs", LogLevel.SUCCESS);
                    XnrgyMessageBox.ShowSuccess(
                        $"Upload termine!\n\n" +
                        $"Fichiers uploades: {_uploadedCount}\n" +
                        $"Echecs: {_failedCount}",
                        "Succes", this);
                }
                else
                {
                    Log($"[-] Aucun fichier uploade", LogLevel.ERROR);
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur upload: {ex.Message}", LogLevel.ERROR);
            }
            finally
            {
                _isProcessing = false;
                SetProcessingState(false);
                UpdateProgress(0, 100, "Termine");
            }
        }

        private bool UploadFileToVault(VaultUploadFileItem file, string comment)
        {
            if (_vaultService == null) return false;

            try
            {
                // Calculer le chemin Vault relatif
                string localRoot = @"C:\Vault";
                string vaultPath = "$/";

                if (file.FullPath.StartsWith(localRoot, StringComparison.OrdinalIgnoreCase))
                {
                    string relativePath = file.FullPath.Substring(localRoot.Length).TrimStart('\\');
                    string? folder = Path.GetDirectoryName(relativePath);
                    vaultPath = "$/" + (folder?.Replace("\\", "/") ?? "");
                }

                // Extraire les proprietes du fichier (Project, Reference, Module)
                string? projectNumber = _projectProperties?.ProjectNumber;
                string? reference = _projectProperties?.Reference;
                string? module = _projectProperties?.Module;

                // Upload fichier vers Vault avec proprietes et commentaire
                // Utilise UploadFile qui applique les proprietes (Check-in/Lifecycle/Properties)
                return _vaultService.UploadFile(
                    file.FullPath,
                    vaultPath,
                    projectNumber,
                    reference,
                    module,
                    categoryId: null,
                    categoryName: null,
                    lifecycleDefinitionId: null,
                    lifecycleStateId: null,
                    revision: null,
                    checkInComment: comment
                );
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur upload {file.FileName}: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        // ====================================================================
        // Controles traitement
        // ====================================================================
        private void Pause_Click(object sender, RoutedEventArgs e)
        {
            _isPaused = !_isPaused;
            BtnPause.Content = _isPaused ? "[>] REPRENDRE" : "[~] PAUSE";
            Log(_isPaused ? "[~] Upload en pause" : "[>] Upload repris", LogLevel.INFO);
        }

        private void Stop_Click(object sender, RoutedEventArgs e)
        {
            Log("[!] Arret demande - Finalisation du fichier en cours...", LogLevel.WARNING);
            _cancellationTokenSource?.Cancel();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            var confirm = XnrgyMessageBox.Show(
                "Voulez-vous vraiment annuler l'upload?",
                "Confirmation",
                XnrgyMessageBoxType.Warning,
                XnrgyMessageBoxButtons.YesNo,
                this);

            if (confirm == XnrgyMessageBoxResult.Yes)
            {
                Log("[-] Upload annule", LogLevel.WARNING);
                _cancellationTokenSource?.Cancel();
            }
        }

        // ====================================================================
        // Utilitaires
        // ====================================================================
        private void SetProcessingState(bool isProcessing)
        {
            Dispatcher.Invoke(() =>
            {
                BtnCheckIn.IsEnabled = !isProcessing;
                BtnPause.IsEnabled = isProcessing;
                BtnStop.IsEnabled = isProcessing;
                BtnCancel.IsEnabled = isProcessing;
            });
        }

        private void UpdateProgress(int current, int total, string fileName)
        {
            Dispatcher.Invoke(() =>
            {
                int percent = total > 0 ? (current * 100 / total) : 0;
                ProgressBar.Value = percent;
                TxtProgress.Text = $"{percent}% - {current}/{total} - {fileName}";
            });
        }

        private void UpdateStatistics()
        {
            Dispatcher.Invoke(() =>
            {
                // Mise à jour des statistiques dans le header
                int totalCount = _allFilesMaster.Count;
                int inventorCount = _allFilesMaster.Count(f => f.IsInventorFile);
                int nonInventorCount = totalCount - inventorCount;
                int selectedCount = _allFilesMaster.Count(f => f.IsSelected);
                
                if (TxtStatsTotal != null) TxtStatsTotal.Text = totalCount.ToString();
                TxtStatsInventor.Text = inventorCount.ToString();
                TxtStatsNonInventor.Text = nonInventorCount.ToString();
                TxtStatsSelected.Text = selectedCount.ToString();
            });
        }

        private void UpdatePropertiesDisplay()
        {
            Dispatcher.Invoke(() =>
            {
                if (_projectProperties != null && !string.IsNullOrWhiteSpace(_projectProperties.ProjectNumber))
                {
                    PropertiesPanel.Visibility = Visibility.Visible;
                    TxtProjectNumber.Text = _projectProperties.ProjectNumber;
                    TxtReference.Text = _projectProperties.Reference;
                    TxtModule.Text = _projectProperties.Module;
                }
                else
                {
                    PropertiesPanel.Visibility = Visibility.Collapsed;
                }
            });
        }

        private VaultProjectProperties? ExtractPropertiesFromPath(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path)) return null;

                path = path.TrimEnd('\\', '/');
                var parts = path.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);

                // Chercher "Projects" et extraire les 3 niveaux suivants
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
                    string projectNumber = parts[projectsIndex + 1];
                    string refFolder = parts[projectsIndex + 2];
                    string moduleFolder = parts[projectsIndex + 3];

                    if (Regex.IsMatch(projectNumber, @"^\d+$") &&
                        Regex.IsMatch(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase) &&
                        Regex.IsMatch(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase))
                    {
                        var refMatch = Regex.Match(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase);
                        var moduleMatch = Regex.Match(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase);

                        return new VaultProjectProperties
                        {
                            ProjectNumber = projectNumber,
                            Reference = refMatch.Groups[1].Value,
                            Module = moduleMatch.Groups[1].Value
                        };
                    }
                }

                // Alternative: 3 derniers dossiers
                if (parts.Length >= 3)
                {
                    string projectNumber = parts[parts.Length - 3];
                    string refFolder = parts[parts.Length - 2];
                    string moduleFolder = parts[parts.Length - 1];

                    if (Regex.IsMatch(projectNumber, @"^\d+$") &&
                        Regex.IsMatch(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase) &&
                        Regex.IsMatch(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase))
                    {
                        var refMatch = Regex.Match(refFolder, @"^REF(\d+)$", RegexOptions.IgnoreCase);
                        var moduleMatch = Regex.Match(moduleFolder, @"^M(\d+)$", RegexOptions.IgnoreCase);

                        return new VaultProjectProperties
                        {
                            ProjectNumber = projectNumber,
                            Reference = refMatch.Groups[1].Value,
                            Module = moduleMatch.Groups[1].Value
                        };
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            double kb = bytes / 1024.0;
            if (kb < 1024) return kb.ToString("0.0") + " KB";
            double mb = kb / 1024.0;
            if (mb < 1024) return mb.ToString("0.0") + " MB";
            double gb = mb / 1024.0;
            return gb.ToString("0.00") + " GB";
        }

        // ====================================================================
        // Journal - Utilise JournalColorService pour uniformite
        // ====================================================================
        private enum LogLevel { INFO, SUCCESS, WARNING, ERROR }

        private void Log(string message, LogLevel level)
        {
            Dispatcher.Invoke(() =>
            {
                var paragraph = new Paragraph();
                var run = new Run($"[{DateTime.Now:HH:mm:ss}] {message}");

                // Utilise JournalColorService pour les couleurs uniformisees
                var serviceLevel = level switch
                {
                    LogLevel.SUCCESS => Services.JournalColorService.LogLevel.SUCCESS,
                    LogLevel.ERROR => Services.JournalColorService.LogLevel.ERROR,
                    LogLevel.WARNING => Services.JournalColorService.LogLevel.WARNING,
                    _ => Services.JournalColorService.LogLevel.INFO
                };
                run.Foreground = Services.JournalColorService.GetBrushForLevel(serviceLevel);

                paragraph.Inlines.Add(run);
                paragraph.Margin = new Thickness(0, 2, 0, 2);
                LogBox.Document.Blocks.Add(paragraph);
                LogBox.ScrollToEnd();
            });

            // Logger aussi vers fichier
            Services.Logger.Log(message, level == LogLevel.ERROR ? Services.Logger.LogLevel.ERROR :
                                         level == LogLevel.WARNING ? Services.Logger.LogLevel.WARNING :
                                         Services.Logger.LogLevel.INFO);
        }

        /// <summary>
        /// Toggle la checkbox IsSelected quand on clique sur une ligne du DataGrid
        /// </summary>
        private void DataGrid_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                if (sender is System.Windows.Controls.DataGrid dataGrid)
                {
                    // Verifier si on a clique directement sur la checkbox (ne pas toggle deux fois)
                    var originalSource = e.OriginalSource as System.Windows.FrameworkElement;
                    if (originalSource != null)
                    {
                        // Si c'est un CheckBox, ne pas faire le toggle (deja gere par le CheckBox)
                        var parent = originalSource;
                        while (parent != null)
                        {
                            if (parent is System.Windows.Controls.CheckBox)
                            {
                                return;
                            }
                            parent = System.Windows.Media.VisualTreeHelper.GetParent(parent) as System.Windows.FrameworkElement;
                        }
                    }

                    // Toggle la selection de l'item clique - utiliser VaultUploadFileItem
                    if (dataGrid.SelectedItem is VaultUploadFileItem fileItem)
                    {
                        fileItem.IsSelected = !fileItem.IsSelected;
                        UpdateStatistics();
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur toggle selection: {ex.Message}", LogLevel.ERROR);
            }
        }

        private void ClearLog_Click(object sender, RoutedEventArgs e)
        {
            LogBox.Document.Blocks.Clear();
            Log("[i] Journal efface", LogLevel.INFO);
        }
    }
}
