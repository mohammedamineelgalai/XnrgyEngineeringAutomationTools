using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using XnrgyEngineeringAutomationTools.Models;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Views
{
    /// <summary>
    /// Fenêtre de création de module XNRGY - Pack & Go / Copy Design
    /// </summary>
    public partial class CreateModuleWindow : Window
    {
        private readonly CreateModuleRequest _request;
        private readonly ObservableCollection<FileRenameItem> _files;
        private readonly string _defaultTemplatePath = @"C:\Vault\Engineering\Library\Xnrgy_Module";
        private readonly string _projectsBasePath = @"C:\Vault\Engineering\Projects";
        private readonly string _defaultDestinationBase = @"C:\Vault\Engineering\Projects";

        // Liste des initiales dessinateurs XNRGY
        private readonly List<string> _designerInitials = new List<string>
        {
            "N/A", "AC", "AM", "AR", "CC", "DC", "DL", "DM", "FL", "IM",
            "KB", "KJ", "MAE", "MC", "NJ", "TG", "TV", "SB", "VK", "YS"
        };

        public CreateModuleWindow()
        {
            // IMPORTANT: Initialiser _request et _files AVANT InitializeComponent()
            // car les événements TextChanged du XAML sont déclenchés pendant l'initialisation
            _request = new CreateModuleRequest();
            _files = new ObservableCollection<FileRenameItem>();
            
            InitializeComponent();
            
            // Attendre que la fenêtre soit chargée pour initialiser les contrôles
            this.Loaded += CreateModuleWindow_Loaded;
        }

        #region Journal des Opérations

        /// <summary>
        /// Ajoute un message au journal des opérations avec style coloré
        /// </summary>
        private void AddLog(string message, string level = "INFO")
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string icon = level switch
            {
                "ERROR" => "✕",
                "WARN" => "⚠",
                "SUCCESS" => "✓",
                "START" => "▶",
                "STOP" => "■",
                "CRITICAL" => "✕✕",
                _ => "ℹ"
            };
            
            string text = $"[{timestamp}] {icon} {message}";
            
            Dispatcher.Invoke(() =>
            {
                var textBlock = new TextBlock
                {
                    Text = text,
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = 13,
                    Padding = new Thickness(10, 4, 10, 4),
                    TextWrapping = TextWrapping.Wrap
                };
                
                switch (level)
                {
                    case "ERROR":
                    case "CRITICAL":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(255, 80, 80));
                        textBlock.FontWeight = FontWeights.Bold;
                        StartBlinkAnimation(textBlock, Color.FromRgb(255, 80, 80), Color.FromRgb(255, 200, 200));
                        break;
                    case "ACTION":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(255, 100, 100));
                        textBlock.FontWeight = FontWeights.Bold;
                        StartBlinkAnimation(textBlock, Color.FromRgb(255, 100, 100), Color.FromRgb(255, 50, 50));
                        break;
                    case "READY":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(50, 255, 100));
                        textBlock.FontWeight = FontWeights.Bold;
                        StartBlinkAnimation(textBlock, Color.FromRgb(50, 255, 100), Color.FromRgb(150, 255, 180));
                        break;
                    case "WARN":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(255, 180, 0));
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    case "SUCCESS":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(80, 200, 80));
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    case "START":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(100, 180, 255));
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    case "STOP":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(180, 100, 255));
                        break;
                    default:
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(220, 220, 220));
                        break;
                }
                
                LogListBox.Items.Add(textBlock);
                LogListBox.ScrollIntoView(textBlock);
                
                // Limiter à 100 entrées
                while (LogListBox.Items.Count > 100)
                {
                    LogListBox.Items.RemoveAt(0);
                }
            });
        }

        private void StartBlinkAnimation(TextBlock textBlock, Color fromColor, Color toColor)
        {
            var animation = new ColorAnimation
            {
                From = fromColor,
                To = toColor,
                Duration = TimeSpan.FromMilliseconds(300),
                AutoReverse = true,
                RepeatBehavior = new RepeatBehavior(5)
            };
            
            var brush = new SolidColorBrush(fromColor);
            textBlock.Foreground = brush;
            brush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            LogListBox.Items.Clear();
            AddLog("Journal effacé", "INFO");
        }

        /// <summary>
        /// Met à jour la barre de progression avec animation
        /// </summary>
        private void UpdateProgress(int percent, string statusText, bool isError = false)
        {
            Dispatcher.Invoke(() =>
            {
                // Calculer la largeur de la barre (basée sur la largeur du conteneur parent)
                var container = ProgressBarFill.Parent as FrameworkElement;
                if (container != null)
                {
                    double maxWidth = container.ActualWidth > 0 ? container.ActualWidth : 400;
                    double targetWidth = (percent / 100.0) * maxWidth;

                    // Animation de la largeur
                    var widthAnimation = new DoubleAnimation
                    {
                        To = targetWidth,
                        Duration = TimeSpan.FromMilliseconds(300),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                    };

                    ProgressBarFill.BeginAnimation(WidthProperty, widthAnimation);
                    ProgressBarShine.BeginAnimation(WidthProperty, widthAnimation);
                }

                // Couleur de la barre selon l'état
                if (isError)
                {
                    ProgressBarFill.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC3545"));
                }
                else if (percent >= 100)
                {
                    ProgressBarFill.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#28A745"));
                }
                else
                {
                    ProgressBarFill.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#007ACC"));
                }

                // Mise à jour du texte
                TxtStatus.Text = statusText;
                TxtProgressPercent.Text = percent > 0 ? $"{percent}%" : "";
            });
        }

        /// <summary>
        /// Réinitialise la barre de progression
        /// </summary>
        private void ResetProgress()
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBarFill.Width = 0;
                ProgressBarShine.Width = 0;
                ProgressBarFill.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#007ACC"));
                TxtStatus.Text = "Prêt - Remplissez les informations du projet";
                TxtProgressPercent.Text = "";
            });
        }

        #endregion

        private void CreateModuleWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                DgFiles.ItemsSource = _files;
                
                // Message de bienvenue dans le journal
                AddLog("Fenêtre Créer Module initialisée", "START");
                
                // Initialiser les ComboBox Référence et Module (01-50)
                InitializeReferenceModuleComboBoxes();
                AddLog("ComboBox Référence/Module chargées (01-50)", "INFO");
                
                // Initialiser les ComboBox Initiales Dessinateurs
                InitializeDesignerComboBoxes();
                AddLog($"Liste des dessinateurs chargée ({_designerInitials.Count} initiales)", "INFO");
                
                // Initialiser la date de création avec DatePicker
                var today = DateTime.Now;
                DpCreationDate.SelectedDate = today;
                _request.CreationDate = today;
                
                // Charger automatiquement le template si disponible
                if (Directory.Exists(_defaultTemplatePath))
                {
                    TxtSourcePath.Text = _defaultTemplatePath;
                    LoadFilesFromPath(_defaultTemplatePath);
                    AddLog($"Template chargé: {_defaultTemplatePath}", "SUCCESS");
                    TxtStatus.Text = "✓ Template Xnrgy_Module chargé automatiquement";
                }
                else
                {
                    AddLog($"Template non trouvé: {_defaultTemplatePath}", "WARN");
                }
                
                UpdateDestinationPreview();
                UpdateStatistics();
                AddLog("Prêt - Remplissez les informations du projet", "INFO");
            }
            catch (Exception ex)
            {
                AddLog($"Erreur d'initialisation: {ex.Message}", "ERROR");
            }
        }

        private void InitializeReferenceModuleComboBoxes()
        {
            // Remplir Référence et Module avec 01 à 50
            for (int i = 1; i <= 50; i++)
            {
                string value = i.ToString("D2"); // Format 01, 02, ... 50
                CmbReference.Items.Add(value);
                CmbModule.Items.Add(value);
            }
            
            // Sélectionner 01 par défaut
            CmbReference.SelectedIndex = 0;
            CmbModule.SelectedIndex = 0;
        }

        private void InitializeDesignerComboBoxes()
        {
            // Remplir les ComboBox avec les initiales dessinateurs
            foreach (var initial in _designerInitials)
            {
                CmbInitialeDessinateur.Items.Add(initial);
                CmbInitialeCoDessinateur.Items.Add(initial);
            }
            
            // Sélectionner N/A par défaut pour le co-dessinateur
            CmbInitialeDessinateur.SelectedIndex = 0;
            CmbInitialeCoDessinateur.SelectedIndex = 0; // N/A
        }

        private void CmbReference_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDestinationPreview();
            UpdateFullProjectNumber();
            ValidateForm();
        }

        private void CmbModule_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDestinationPreview();
            UpdateFullProjectNumber();
            ValidateForm();
        }

        private void CmbInitialeDessinateur_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ValidateForm();
        }

        private void CmbInitialeCoDessinateur_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Pas de validation requise car optionnel
        }

        #region Event Handlers - Project Info

        private void ProjectInfo_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateDestinationPreview();
            UpdateFullProjectNumber();
            ValidateForm();
        }

        private void TxtJobTitle_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Vérifier que la fenêtre est chargée avant de valider
            if (!IsLoaded || BtnCreateModule == null) return;
            ValidateForm();
        }

        private void UpdateDestinationPreview()
        {
            if (TxtProject == null || CmbReference == null || CmbModule == null) return;

            var project = TxtProject.Text?.Trim() ?? "";
            var reference = CmbReference.SelectedItem?.ToString() ?? "01";
            var module = CmbModule.SelectedItem?.ToString() ?? "01";

            string destPath;
            if (string.IsNullOrEmpty(project))
            {
                destPath = $"{_defaultDestinationBase}\\[PROJET]\\REF{reference}\\M{module}";
            }
            else
            {
                destPath = Path.Combine(_defaultDestinationBase, project, $"REF{reference}", $"M{module}");
            }

            if (TxtDestinationPath != null)
                TxtDestinationPath.Text = destPath;

            if (TxtDestinationPreview != null)
                TxtDestinationPreview.Text = $"Destination: {destPath}";

            // Mettre à jour les propriétés du request (DestinationPath est calculé automatiquement)
            _request.Project = project;
            _request.Reference = reference;
            _request.Module = module;

            // Mettre à jour les chemins de destination pour chaque fichier
            UpdateFileDestinationPaths(destPath);
        }

        private void UpdateFileDestinationPaths(string destinationBase)
        {
            if (_files == null || _files.Count == 0) return;

            bool isFromExistingProject = _request.Source == CreateModuleSource.FromExistingProject;

            foreach (var file in _files)
            {
                // Calculer le chemin relatif une seule fois
                var relativeDir = Path.GetDirectoryName(file.RelativePath) ?? "";
                var isAtRoot = string.IsNullOrEmpty(relativeDir);
                
                // Renommer le Top Assembly (.iam) avec le numéro de projet formaté
                // Pour templates: fichier marqué IsTopAssembly (Module_.iam)
                // Pour projets existants: premier .iam à la racine
                if (!string.IsNullOrEmpty(_request.FullProjectNumber))
                {
                    if (file.IsTopAssembly)
                    {
                        file.NewFileName = $"{_request.FullProjectNumber}.iam";
                    }
                    else if (isFromExistingProject && isAtRoot && file.FileType == "IAM")
                    {
                        // Pour projets existants: renommer le .iam à la racine (un seul)
                        // Vérifier qu'on n'a pas déjà renommé un autre .iam
                        var alreadyRenamed = _files.Any(f => f != file && f.FileType == "IAM" && 
                            string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath)) &&
                            f.NewFileName == $"{_request.FullProjectNumber}.iam");
                        if (!alreadyRenamed)
                        {
                            file.NewFileName = $"{_request.FullProjectNumber}.iam";
                        }
                    }
                }
                
                // Renommer le fichier projet principal (.ipj)
                // Pour templates: pattern XXXXX-XX-XX_2026.ipj
                // Pour projets existants: tout .ipj à la racine
                if (file.FileType == "IPJ" && !string.IsNullOrEmpty(_request.FullProjectNumber) && isAtRoot)
                {
                    if (isFromExistingProject || IsMainProjectFilePattern(file.OriginalFileName))
                    {
                        file.NewFileName = $"{_request.FullProjectNumber}.ipj";
                    }
                }
                
                // Construire le chemin de destination en conservant la structure relative
                var fileName = !string.IsNullOrEmpty(file.NewFileName) ? file.NewFileName : file.OriginalFileName;
                
                if (isAtRoot)
                {
                    file.DestinationPath = Path.Combine(destinationBase, fileName);
                }
                else
                {
                    file.DestinationPath = Path.Combine(destinationBase, relativeDir, fileName);
                }
            }

            // Rafraîchir l'affichage
            DgFiles?.Items.Refresh();
        }

        private void UpdateFullProjectNumber()
        {
            if (TxtFullProjectNumber == null) return;

            var project = (TxtProject?.Text?.Trim() ?? "").PadLeft(5, '0');
            var reference = CmbReference?.SelectedItem?.ToString() ?? "01";
            var module = CmbModule?.SelectedItem?.ToString() ?? "01";

            TxtFullProjectNumber.Text = $"{project}{reference}{module}";
        }

        #endregion

        #region Event Handlers - Source Options

        private void SourceOption_Changed(object sender, RoutedEventArgs e)
        {
            if (RbFromTemplate == null || RbFromExisting == null) return;
            if (PnlProjectSelector == null) return;

            if (RbFromTemplate.IsChecked == true)
            {
                _request.Source = CreateModuleSource.FromTemplate;
                TxtSourcePath.Text = _defaultTemplatePath;
                TxtSourcePath.IsReadOnly = true;
                PnlProjectSelector.Visibility = Visibility.Collapsed;
                TxtStatus.Text = "Mode: Création depuis Template";
                
                // Charger le template automatiquement
                if (Directory.Exists(_defaultTemplatePath))
                {
                    LoadFilesFromPath(_defaultTemplatePath);
                }
            }
            else
            {
                _request.Source = CreateModuleSource.FromExistingProject;
                TxtSourcePath.Text = "";
                TxtSourcePath.IsReadOnly = true;
                PnlProjectSelector.Visibility = Visibility.Visible;
                TxtStatus.Text = "Mode: Création depuis Projet Existant - Sélectionnez un projet";
                
                // Charger la liste des projets
                LoadProjectsList();
            }

            _files.Clear();
            UpdateStatistics();
            ValidateForm();
        }

        private void LoadProjectsList()
        {
            if (CmbProjects == null) return;
            
            CmbProjects.Items.Clear();
            
            if (!Directory.Exists(_projectsBasePath))
            {
                TxtStatus.Text = "[!] Dossier Projects non trouve";
                return;
            }

            try
            {
                // Trouver tous les modules dans les projets
                var modules = new List<string>();
                
                foreach (var projectDir in Directory.GetDirectories(_projectsBasePath))
                {
                    var projectName = Path.GetFileName(projectDir);
                    
                    // Chercher les REF
                    foreach (var refDir in Directory.GetDirectories(projectDir, "REF*"))
                    {
                        var refName = Path.GetFileName(refDir);
                        
                        // Chercher les Modules
                        foreach (var moduleDir in Directory.GetDirectories(refDir, "M*"))
                        {
                            var moduleName = Path.GetFileName(moduleDir);
                            var displayName = $"{projectName} / {refName} / {moduleName}";
                            CmbProjects.Items.Add(new ComboBoxItem 
                            { 
                                Content = displayName, 
                                Tag = moduleDir 
                            });
                        }
                    }
                }
                
                TxtStatus.Text = $"[+] {CmbProjects.Items.Count} modules trouves";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"[!] Erreur chargement projets: {ex.Message}";
            }
        }

        private void CmbProjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbProjects?.SelectedItem is ComboBoxItem item && item.Tag is string path)
            {
                TxtSourcePath.Text = path;
                _request.SourceExistingProjectPath = path;
                LoadFilesFromPath(path);
            }
        }

        private void BtnRefreshProjects_Click(object sender, RoutedEventArgs e)
        {
            LoadProjectsList();
        }

        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Sélectionner le dossier source",
                ShowNewFolderButton = false
            };

            // Si mode template, commencer dans Library
            if (RbFromTemplate?.IsChecked == true && Directory.Exists(_defaultTemplatePath))
            {
                dialog.SelectedPath = _defaultTemplatePath;
            }
            else if (Directory.Exists(_defaultDestinationBase))
            {
                dialog.SelectedPath = _defaultDestinationBase;
            }

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TxtSourcePath.Text = dialog.SelectedPath;
                _request.SourceExistingProjectPath = dialog.SelectedPath;
            }
        }

        private void BtnBrowseDestination_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Sélectionner le dossier de destination de base",
                ShowNewFolderButton = true
            };

            if (Directory.Exists(_defaultDestinationBase))
            {
                dialog.SelectedPath = _defaultDestinationBase;
            }

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _request.DestinationBasePath = dialog.SelectedPath;
                UpdateDestinationPreview();
            }
        }

        #endregion

        #region Event Handlers - Load Files

        private void BtnLoadFiles_Click(object sender, RoutedEventArgs e)
        {
            LoadFilesFromPath(TxtSourcePath?.Text);
        }

        /// <summary>
        /// Charge tous les fichiers depuis un dossier source (copie complète du dossier)
        /// </summary>
        private void LoadFilesFromPath(string sourcePath)
        {
            try
            {
                _files.Clear();
                if (TxtStatus != null) TxtStatus.Text = "Chargement des fichiers...";

                if (string.IsNullOrEmpty(sourcePath) || !Directory.Exists(sourcePath))
                {
                    if (TxtStatus != null) TxtStatus.Text = "[!] Chemin source invalide";
                    return;
                }

                // Determiner si on est en mode "Projet Existant"
                bool isFromExistingProject = _request.Source == CreateModuleSource.FromExistingProject;

                // Scanner TOUS les fichiers du dossier (copie design complète)
                // Exclure: fichiers temporaires, Vault, .bak, dossiers _V et OldVersions
                var vaultTempExtensions = new[] { ".v", ".v1", ".v2", ".v3", ".v4", ".v5", ".vbak", ".bak" };
                var excludedFolders = new[] { "_V", "OldVersions", "oldversions" };
                
                var allFiles = Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories)
                    .Where(f => 
                    {
                        var fileName = Path.GetFileName(f);
                        var ext = Path.GetExtension(f).ToLower();
                        var dirPath = Path.GetDirectoryName(f) ?? "";
                        
                        // Exclure les fichiers temporaires
                        if (fileName.StartsWith(".") || fileName.StartsWith("~"))
                            return false;
                        
                        // Exclure les fichiers .bak
                        if (ext == ".bak")
                            return false;
                        
                        // Exclure les fichiers temporaires Vault (.v, .v1, .v2, etc.)
                        if (vaultTempExtensions.Any(ve => ext.StartsWith(ve)))
                            return false;
                        
                        // Exclure les dossiers _V et OldVersions
                        if (excludedFolders.Any(ef => dirPath.Contains($"\\{ef}\\") || dirPath.EndsWith($"\\{ef}")))
                            return false;
                        
                        // Pour les projets existants: NE PAS exclure l'IPJ, il sera copié et renommé
                        // Pour les templates: tout inclure aussi
                        
                        return true;
                    })
                    .ToList();

                // Extensions Inventor pour identifier le type
                var inventorExtensions = new[] { ".iam", ".ipt", ".idw", ".dwg", ".ipn" };

                // Trier: IAM en premier, puis IPT, IDW, puis autres
                var sortedFiles = allFiles
                    .OrderBy(f => 
                    {
                        var ext = Path.GetExtension(f).ToLower();
                        if (ext == ".iam") return 0;
                        if (ext == ".ipt") return 1;
                        if (ext == ".idw") return 2;
                        if (ext == ".dwg") return 3;
                        return 4;
                    })
                    .ThenBy(f => Path.GetFileName(f))
                    .ToList();

                foreach (var file in sortedFiles)
                {
                    var fileName = Path.GetFileName(file);
                    var extension = Path.GetExtension(file).ToUpper().TrimStart('.');
                    var relativePath = file.Substring(sourcePath.Length).TrimStart('\\');
                    var isInventorFile = inventorExtensions.Contains(Path.GetExtension(file).ToLower());
                    var isTopAssembly = fileName.Equals("Module_.iam", StringComparison.OrdinalIgnoreCase);
                    var isProjectFile = extension == "IPJ";
                    // Fichier projet principal: pattern XXXXX-XX-XX_2026.ipj (à la racine du module)
                    var isMainProjectFile = isProjectFile && 
                                           string.IsNullOrEmpty(Path.GetDirectoryName(relativePath)) &&
                                           IsMainProjectFilePattern(fileName);

                    var item = new FileRenameItem
                    {
                        IsSelected = true,
                        OriginalPath = file,
                        RelativePath = relativePath,
                        NewFileName = fileName,
                        FileType = extension,
                        Status = isInventorFile ? "Inventor" : (isProjectFile ? "Projet" : "Copie simple"),
                        IsTopAssembly = isTopAssembly,
                        IsInventorFile = isInventorFile
                    };

                    // Pour templates: renommer Module_.iam et fichier IPJ principal (pattern XXXXX-XX-XX_2026.ipj)
                    // Pour projets existants: renommer le premier .iam à la racine et le premier .ipj à la racine
                    if (!isFromExistingProject)
                    {
                        // Renommage automatique du Top Assembly (.iam) - template Module_.iam
                        if (isTopAssembly && !string.IsNullOrEmpty(TxtProject?.Text))
                        {
                            item.NewFileName = $"{_request.FullProjectNumber}.iam";
                        }
                        
                        // Renommage automatique UNIQUEMENT du fichier projet principal (.ipj)
                        // Pattern: XXXXX-XX-XX_2026.ipj (à la racine)
                        if (isMainProjectFile && !string.IsNullOrEmpty(TxtProject?.Text))
                        {
                            item.NewFileName = $"{_request.FullProjectNumber}.ipj";
                        }
                    }
                    else
                    {
                        // PROJET EXISTANT: Renommer le fichier Top Assembly (.iam à la racine) avec le numéro de projet
                        // Note: Pour les projets existants, isTopAssembly est false car le fichier ne s'appelle pas "Module_.iam"
                        // On détecte le premier .iam à la racine comme Top Assembly
                        bool isRootIam = extension == "IAM" && string.IsNullOrEmpty(Path.GetDirectoryName(relativePath));
                        bool isRootIpj = isProjectFile && string.IsNullOrEmpty(Path.GetDirectoryName(relativePath));
                        
                        if (isRootIam && !string.IsNullOrEmpty(TxtProject?.Text))
                        {
                            // Vérifier si c'est le premier .iam trouvé à la racine (Top Assembly)
                            bool alreadyHasTopAssembly = _files.Any(f => f.IsTopAssembly || 
                                (f.FileType == "IAM" && string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath))));
                            
                            if (!alreadyHasTopAssembly)
                            {
                                item.IsTopAssembly = true;
                                item.NewFileName = $"{_request.FullProjectNumber}.iam";
                            }
                        }
                        
                        if (isRootIpj && !string.IsNullOrEmpty(TxtProject?.Text))
                        {
                            // Vérifier si c'est le premier .ipj trouvé à la racine (fichier projet)
                            bool alreadyHasProjectFile = _files.Any(f => f.FileType == "IPJ" && 
                                string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath)));
                            
                            if (!alreadyHasProjectFile)
                            {
                                item.NewFileName = $"{_request.FullProjectNumber}.ipj";
                            }
                        }
                    }
                    
                    // Calculer le chemin de destination
                    UpdateFileDestination(item);

                    _files.Add(item);
                }

                // Mettre à jour le filtre des types avec les extensions réellement trouvées
                UpdateFileTypeFilter();

                UpdateStatistics();
                UpdateFileCount();
                AddLog($"{_files.Count} fichiers chargés depuis {Path.GetFileName(sourcePath)}", "SUCCESS");
                if (TxtStatus != null) TxtStatus.Text = $"✓ {_files.Count} fichiers chargés depuis {Path.GetFileName(sourcePath)}";
                ValidateForm();
            }
            catch (Exception ex)
            {
                AddLog($"Erreur lors du chargement des fichiers: {ex.Message}", "ERROR");
                if (TxtStatus != null) TxtStatus.Text = $"Erreur: {ex.Message}";
            }
        }

        #endregion

        #region Event Handlers - Rename Options

        private void RenameOptions_Changed(object sender, TextChangedEventArgs e)
        {
            // Mise à jour en temps réel désactivée pour performance
            // L'utilisateur doit cliquer sur "Appliquer"
        }

        private void BtnApplyRename_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var search = TxtSearch?.Text ?? "";
                var replace = TxtReplace?.Text ?? "";
                var prefix = TxtPrefix?.Text ?? "";
                var suffix = TxtSuffix?.Text ?? "";

                foreach (var file in _files.Where(f => f.IsSelected))
                {
                    var newName = file.OriginalFileName;

                    // Appliquer rechercher/remplacer (case-insensitive)
                    if (!string.IsNullOrEmpty(search))
                    {
                        // .NET Framework compatible replace
                        var index = newName.IndexOf(search, StringComparison.OrdinalIgnoreCase);
                        while (index >= 0)
                        {
                            newName = newName.Substring(0, index) + replace + newName.Substring(index + search.Length);
                            index = newName.IndexOf(search, index + replace.Length, StringComparison.OrdinalIgnoreCase);
                        }
                    }

                    // Appliquer préfixe/suffixe
                    if (!string.IsNullOrEmpty(prefix) || !string.IsNullOrEmpty(suffix))
                    {
                        var nameWithoutExt = Path.GetFileNameWithoutExtension(newName);
                        var ext = Path.GetExtension(newName);
                        newName = $"{prefix}{nameWithoutExt}{suffix}{ext}";
                    }

                    file.NewFileName = newName;
                }

                // Toujours renommer le Top Assembly et IPJ avec le numéro de projet
                bool isFromExistingProject = _request.Source == CreateModuleSource.FromExistingProject;
                
                // Top Assembly
                var topAssembly = _files.FirstOrDefault(f => f.IsTopAssembly);
                if (topAssembly == null && isFromExistingProject)
                {
                    // Pour projets existants: premier .iam à la racine
                    topAssembly = _files.FirstOrDefault(f => f.FileType == "IAM" && 
                        string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath)));
                }
                if (topAssembly != null && !string.IsNullOrEmpty(TxtProject?.Text))
                {
                    topAssembly.NewFileName = $"{_request.FullProjectNumber}.iam";
                }
                
                // IPJ principal
                var mainIpj = _files.FirstOrDefault(f => f.FileType == "IPJ" && 
                    string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath)) &&
                    (isFromExistingProject || IsMainProjectFilePattern(f.OriginalFileName)));
                if (mainIpj != null && !string.IsNullOrEmpty(TxtProject?.Text))
                {
                    mainIpj.NewFileName = $"{_request.FullProjectNumber}.ipj";
                }

                AddLog("Renommage appliqué aux fichiers sélectionnés", "SUCCESS");
                TxtStatus.Text = "✓ Renommage appliqué";
            }
            catch (Exception ex)
            {
                AddLog($"Erreur lors du renommage: {ex.Message}", "ERROR");
            }
        }

        private void BtnResetNames_Click(object sender, RoutedEventArgs e)
        {
            foreach (var file in _files)
            {
                file.NewFileName = file.OriginalFileName;
            }

            TxtSearch.Text = "";
            TxtReplace.Text = "";
            TxtPrefix.Text = "";
            TxtSuffix.Text = "";

            TxtStatus.Text = "✓ Noms réinitialisés";
        }

        #endregion

        #region Event Handlers - Selection & Filtering

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var file in _files)
            {
                file.IsSelected = true;
            }
            UpdateFileCount();
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var file in _files)
            {
                file.IsSelected = false;
            }
            UpdateFileCount();
        }

        private void BtnInvertSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (var file in _files)
            {
                file.IsSelected = !file.IsSelected;
            }
            UpdateFileCount();
            AddLog("Sélection inversée", "INFO");
        }

        private void DgFiles_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Toggle selection on click - bascule la checkbox quand on clique sur la ligne
            if (DgFiles.SelectedItem is FileRenameItem selectedFile)
            {
                selectedFile.IsSelected = !selectedFile.IsSelected;
                DgFiles.Items.Refresh();
                UpdateFileCount();
            }
        }

        private void TxtSearchFiles_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilters();
            BtnClearSearch.Visibility = string.IsNullOrEmpty(TxtSearchFiles.Text) ? Visibility.Collapsed : Visibility.Visible;
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            TxtSearchFiles.Text = "";
        }

        private void CmbFilterType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void CmbFilterSelection_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            if (_files == null || DgFiles == null) return;

            var filteredFiles = _files.AsEnumerable();

            // Filtre par recherche texte
            if (!string.IsNullOrWhiteSpace(TxtSearchFiles?.Text))
            {
                var searchTerm = TxtSearchFiles.Text.ToLower();
                filteredFiles = filteredFiles.Where(f => 
                    f.OriginalFileName.ToLower().Contains(searchTerm) ||
                    f.NewFileName.ToLower().Contains(searchTerm));
            }

            // Filtre par type de fichier (supporte ComboBoxItem et string)
            if (CmbFilterType?.SelectedItem != null)
            {
                string? selectedType = null;
                
                if (CmbFilterType.SelectedItem is ComboBoxItem typeItem)
                {
                    selectedType = typeItem.Content?.ToString();
                }
                else if (CmbFilterType.SelectedItem is string typeString)
                {
                    selectedType = typeString;
                }
                
                if (!string.IsNullOrEmpty(selectedType) && selectedType != "Tous")
                {
                    filteredFiles = filteredFiles.Where(f => f.FileType.Equals(selectedType, StringComparison.OrdinalIgnoreCase));
                }
            }

            // Filtre par état de sélection
            if (CmbFilterSelection?.SelectedItem is ComboBoxItem selItem)
            {
                var selection = selItem.Content.ToString();
                if (selection == "Sélectionnés")
                    filteredFiles = filteredFiles.Where(f => f.IsSelected);
                else if (selection == "Non sélectionnés")
                    filteredFiles = filteredFiles.Where(f => !f.IsSelected);
            }

            DgFiles.ItemsSource = filteredFiles.ToList();
            UpdateFileCount();
        }

        private void UpdateFileCount()
        {
            if (TxtFileCount == null || _files == null) return;

            var displayed = (DgFiles.ItemsSource as IEnumerable<FileItem>)?.Count() ?? _files.Count;
            var selected = _files.Count(f => f.IsSelected);
            TxtFileCount.Text = $"{selected}/{_files.Count} sélectionnés";
        }

        #endregion

        #region Event Handlers - Actions

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            var validation = ValidateAllInputs();
            if (!validation.IsValid)
            {
                AddLog($"Validation échouée: {validation.ErrorMessage}", "WARN");
                return;
            }

            var selectedFiles = _files.Where(f => f.IsSelected).ToList();
            AddLog($"Prévisualisation demandée - {selectedFiles.Count} fichiers sélectionnés", "INFO");
            
            // Mettre à jour la date depuis le DatePicker
            _request.CreationDate = DpCreationDate.SelectedDate ?? DateTime.Now;
            
            // Ouvrir la fenêtre de prévisualisation moderne
            var previewWindow = new PreviewWindow();
            previewWindow.Owner = this;
            previewWindow.SetPreviewData(
                TxtProject?.Text ?? "",
                CmbReference?.SelectedItem?.ToString() ?? "01",
                CmbModule?.SelectedItem?.ToString() ?? "01",
                _request.FullProjectNumber,
                TxtDestinationPath?.Text ?? "",
                CmbInitialeDessinateur?.SelectedItem?.ToString() ?? "",
                CmbInitialeCoDessinateur?.SelectedItem?.ToString() ?? "",
                _request.CreationDate,
                TxtJobTitle?.Text ?? "",
                selectedFiles,
                selectedFiles.Count
            );

            // Si l'utilisateur confirme, procéder à la création
            if (previewWindow.ShowDialog() == true && previewWindow.IsConfirmed)
            {
                AddLog("Création confirmée via prévisualisation", "START");
                ExecuteModuleCreation();
            }
            else
            {
                AddLog("Création annulée par l'utilisateur", "INFO");
            }
        }

        private void BtnCreateModule_Click(object sender, RoutedEventArgs e)
        {
            var validation = ValidateAllInputs();
            if (!validation.IsValid)
            {
                AddLog($"Validation échouée: {validation.ErrorMessage}", "WARN");
                return;
            }

            var selectedFiles = _files.Where(f => f.IsSelected).ToList();
            if (selectedFiles.Count == 0)
            {
                AddLog("Aucun fichier sélectionné pour la copie", "WARN");
                return;
            }

            // Confirmation simple avant création
            var reference = CmbReference?.SelectedItem?.ToString() ?? "01";
            var module = CmbModule?.SelectedItem?.ToString() ?? "01";
            var moduleName = $"{TxtProject.Text?.Trim()}-REF{reference}-M{module}";
            
            AddLog($"Demande de création du module: {moduleName}", "START");
            AddLog($"Fichiers à copier: {selectedFiles.Count}", "INFO");
            AddLog($"Destination: {_request.DestinationPath}", "INFO");
            
            // Pas de MessageBox - l'utilisateur a confirmé en cliquant sur le bouton
            AddLog("Création en cours...", "INFO");
            ExecuteModuleCreation();
        }

        private async void ExecuteModuleCreation()
        {
            try
            {
                BtnCreateModule.IsEnabled = false;
                BtnCancel.IsEnabled = false;
                BtnPreview.IsEnabled = false;
                UpdateProgress(0, "Initialisation du Copy Design...");
                AddLog("Début de la création du module...", "START");

                // Préparer les données depuis les ComboBox
                var initialeDessinateur = CmbInitialeDessinateur?.SelectedItem?.ToString() ?? "";
                var initialeCoDessinateur = CmbInitialeCoDessinateur?.SelectedItem?.ToString() ?? "";
                
                // Si N/A est sélectionné, on le traite comme vide
                _request.InitialeDessinateur = initialeDessinateur == "N/A" ? "" : initialeDessinateur;
                _request.InitialeCoDessinateur = initialeCoDessinateur == "N/A" ? "" : initialeCoDessinateur;
                _request.JobTitle = TxtJobTitle?.Text ?? "";
                _request.CreationDate = DpCreationDate.SelectedDate ?? DateTime.Now;
                _request.FilesToCopy = _files;

                AddLog($"Dessinateur: {_request.InitialeDessinateur}", "INFO");
                if (!string.IsNullOrEmpty(_request.InitialeCoDessinateur))
                    AddLog($"Co-Dessinateur: {_request.InitialeCoDessinateur}", "INFO");
                AddLog($"Date de création: {_request.CreationDate:yyyy-MM-dd}", "INFO");
                AddLog($"Destination: {_request.DestinationPath}", "INFO");

                // Utiliser le service Copy Design avec Inventor API et callbacks
                using (var copyDesignService = new InventorCopyDesignService(
                    // Callback pour les logs
                    (message, level) =>
                    {
                        Dispatcher.Invoke(() => AddLog(message, level));
                    },
                    // Callback pour la progression
                    (percent, statusText) =>
                    {
                        UpdateProgress(percent, statusText);
                    }))
                {
                    // Initialiser la connexion à Inventor
                    UpdateProgress(2, "Connexion à Inventor...");
                    AddLog("Connexion à Inventor...", "INFO");
                    if (!copyDesignService.Initialize())
                    {
                        UpdateProgress(0, "✗ Erreur: Inventor non disponible", isError: true);
                        throw new Exception("Impossible de se connecter à Inventor. Assurez-vous qu'Inventor 2026 est installé.");
                    }

                    // Exécuter le Copy Design complet
                    AddLog("Exécution du Copy Design...", "INFO");
                    var result = await copyDesignService.ExecuteCopyDesignAsync(_request);

                    if (result.Success)
                    {
                        UpdateProgress(100, $"✓ Module {_request.FullProjectNumber} créé avec succès!");
                        AddLog($"Module {_request.FullProjectNumber} créé avec succès!", "SUCCESS");
                        AddLog($"Fichiers copiés: {result.FilesCopied}", "SUCCESS");
                        AddLog($"iProperties mis à jour: {result.PropertiesUpdated}", "INFO");
                        AddLog($"Durée: {(result.EndTime - result.StartTime).TotalSeconds:F1} secondes", "INFO");

                        // Rafraîchir l'affichage des fichiers pour montrer le statut
                        DgFiles.Items.Refresh();

                        // Ouvrir automatiquement le dossier de destination
                        AddLog($"Ouverture du dossier: {_request.DestinationPath}", "INFO");
                        System.Diagnostics.Process.Start("explorer.exe", _request.DestinationPath);
                        
                        AddLog("✓ Note: Utilisez l'outil principal pour uploader vers Vault", "INFO");
                        
                        // Attendre un peu pour que l'utilisateur voie le message de succès
                        await Task.Delay(1500);
                        DialogResult = true;
                    }
                    else
                    {
                        UpdateProgress(0, $"✗ Erreur: {result.ErrorMessage}", isError: true);
                        AddLog($"ERREUR: {result.ErrorMessage}", "ERROR");
                        AddLog("ACTION", "⚠ Vérifiez les paramètres et réessayez");
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog($"ERREUR CRITIQUE: {ex.Message}", "ERROR");
                UpdateProgress(0, $"✗ Erreur critique: {ex.Message}", isError: true);
                AddLog("ACTION", "⚠ Une erreur inattendue s'est produite. Vérifiez Inventor et réessayez.");
            }
            finally
            {
                BtnCreateModule.IsEnabled = true;
                BtnCancel.IsEnabled = true;
                BtnPreview.IsEnabled = true;
                AddLog("Opération terminée", "STOP");
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            // Fermer sans confirmation - l'utilisateur sait ce qu'il fait
            AddLog("Fermeture de la fenêtre", "STOP");
            DialogResult = false;
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            BtnCancel_Click(sender, e);
        }

        #endregion

        #region Helpers

        private void UpdateStatistics()
        {
            if (TxtCountIam == null) return;

            var iamCount = _files.Count(f => f.FileType == "IAM");
            var iptCount = _files.Count(f => f.FileType == "IPT");
            var idwCount = _files.Count(f => f.FileType == "IDW");
            var otherCount = _files.Count(f => f.FileType != "IAM" && f.FileType != "IPT" && f.FileType != "IDW");

            TxtCountIam.Text = iamCount.ToString();
            TxtCountIpt.Text = iptCount.ToString();
            TxtCountIdw.Text = idwCount.ToString();
            if (TxtCountOther != null) TxtCountOther.Text = otherCount.ToString();
            TxtCountTotal.Text = _files.Count.ToString();
        }

        private void ValidateForm()
        {
            if (BtnCreateModule == null) return;

            var validation = ValidateAllInputs();
            BtnCreateModule.IsEnabled = validation.IsValid;
            
            // Afficher message READY ou ACTION selon la validation
            if (validation.IsValid)
            {
                AddLog("READY", "✓ Prêt pour la création du module");
            }
            else if (!string.IsNullOrEmpty(validation.ErrorMessage))
            {
                AddLog("ACTION", $"⚠ {validation.ErrorMessage}");
            }
        }

        private (bool IsValid, string ErrorMessage) ValidateAllInputs()
        {
            if (string.IsNullOrWhiteSpace(TxtProject?.Text))
                return (false, "Le numéro de projet est requis");

            if (TxtProject.Text.Length < 4)
                return (false, "Le numéro de projet doit contenir au moins 4 chiffres");

            // Vérifier que les initiales du dessinateur sont sélectionnées (et pas N/A)
            var initialeDessinateur = CmbInitialeDessinateur?.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(initialeDessinateur) || initialeDessinateur == "N/A")
                return (false, "Les initiales du dessinateur sont requises (ne peut pas être N/A)");

            // Job Title est requis
            if (string.IsNullOrWhiteSpace(TxtJobTitle?.Text))
                return (false, "Le titre du projet (Job Title) est requis");

            if (!_files.Any(f => f.IsSelected))
                return (false, "Aucun fichier sélectionné");

            return (true, string.Empty);
        }

        /// <summary>
        /// Vérifie si un fichier .ipj correspond au pattern du fichier projet principal
        /// Pattern: XXXXX-XX-XX_2026.ipj ou similaire (contient _2026 ou _202X)
        /// </summary>
        private bool IsMainProjectFilePattern(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return false;
            
            // Le fichier projet principal contient généralement "_2026" ou "_202" dans le nom
            // Exemples: XXXXX-XX-XX_2026.ipj, Module_2026.ipj, etc.
            var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            
            // Pattern 1: Contient _202X (année)
            if (nameWithoutExt.Contains("_202"))
                return true;
            
            // Pattern 2: Format XXXXX-XX-XX (numéro de projet avec tirets)
            if (System.Text.RegularExpressions.Regex.IsMatch(nameWithoutExt, @"^\d{5}-\d{2}-\d{2}"))
                return true;
            
            // Pattern 3: Le nom contient "Module" (fichier projet du module)
            if (nameWithoutExt.IndexOf("Module", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
                
            return false;
        }

        /// <summary>
        /// Met à jour le chemin de destination et le nouveau nom pour un fichier
        /// </summary>
        private void UpdateFileDestination(FileRenameItem item)
        {
            if (item == null) return;
            
            var destBase = _request.DestinationPath;
            if (string.IsNullOrEmpty(destBase)) return;
            
            // Calculer le chemin relatif pour la destination
            var relativePath = item.RelativePath;
            var fileName = item.NewFileName;
            
            // Pour les fichiers à la racine (Module_.iam, .ipj), pas de sous-dossier
            var destDir = Path.GetDirectoryName(relativePath);
            if (string.IsNullOrEmpty(destDir))
            {
                item.DestinationPath = Path.Combine(destBase, fileName);
            }
            else
            {
                item.DestinationPath = Path.Combine(destBase, destDir, fileName);
            }
        }
        
        /// <summary>
        /// Met à jour le filtre des types de fichiers avec les extensions réellement trouvées
        /// </summary>
        private void UpdateFileTypeFilter()
        {
            if (CmbFilterType == null) return;
            
            // Collecter toutes les extensions uniques
            var extensions = _files
                .Select(f => f.FileType)
                .Distinct()
                .OrderBy(e => e)
                .ToList();
            
            // Sauvegarder la sélection actuelle
            var currentSelection = CmbFilterType.SelectedItem?.ToString();
            
            // Mettre à jour les items
            CmbFilterType.Items.Clear();
            CmbFilterType.Items.Add("Tous");
            foreach (var ext in extensions)
            {
                CmbFilterType.Items.Add(ext);
            }
            
            // Restaurer la sélection ou mettre "Tous"
            if (!string.IsNullOrEmpty(currentSelection) && CmbFilterType.Items.Contains(currentSelection))
            {
                CmbFilterType.SelectedItem = currentSelection;
            }
            else
            {
                CmbFilterType.SelectedIndex = 0;
            }
        }
        
        /// <summary>
        /// Met à jour tous les fichiers avec les nouveaux noms et destinations
        /// Appelé quand les champs Projet/Reference/Module changent
        /// </summary>
        private void UpdateAllFileNamesAndDestinations()
        {
            if (_files == null || _files.Count == 0) return;
            
            foreach (var item in _files)
            {
                // Renommer le Top Assembly (.iam)
                if (item.IsTopAssembly)
                {
                    item.NewFileName = $"{_request.FullProjectNumber}.iam";
                }
                
                // Renommer le fichier projet (.ipj)
                if (item.FileType == "IPJ")
                {
                    item.NewFileName = $"{_request.FullProjectNumber}.ipj";
                }
                
                // Mettre à jour le chemin de destination
                UpdateFileDestination(item);
            }
            
            // Rafraîchir l'affichage
            DgFiles?.Items.Refresh();
        }

        #endregion
    }}
