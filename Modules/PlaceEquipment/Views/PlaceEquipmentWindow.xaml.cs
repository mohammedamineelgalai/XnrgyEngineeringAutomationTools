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
using ACW = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using XnrgyEngineeringAutomationTools.Models;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Shared.Views;
using XnrgyEngineeringAutomationTools.Modules.CreateModule.Models;
using XnrgyEngineeringAutomationTools.Modules.CreateModule.Services;
using XnrgyEngineeringAutomationTools.Modules.CreateModule.Views;
using XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models;
using XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Services;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Views
{
    /// <summary>
    /// Fenetre Place Equipment - Copy Design pour equipements XNRGY
    /// </summary>
    public partial class PlaceEquipmentWindow : Window
    {
        private readonly CreateModuleRequest _request;
        private readonly ObservableCollection<FileRenameItem> _files;
        private readonly string _defaultTemplatePath = @"C:\Vault\Engineering\Library\Equipment";
        private readonly string _projectsBasePath = @"C:\Vault\Engineering\Projects";
        
        // Service Place Equipment et equipement selectionne
        private readonly EquipmentPlacementService _equipmentService;
        private EquipmentItem? _selectedEquipment;
        private readonly string _defaultDestinationBase = @"C:\Vault\Engineering\Projects";
        
        // Instance d'equipement selectionnee (suffixe _01, _02, etc.)
        private string _selectedInstanceSuffix = "_01";
        
        // Service Vault pour vérification admin (optionnel)
        private readonly VaultSdkService? _vaultService;
        private bool _isVaultAdmin = false;
        
        // Service Inventor pour vérification de connexion
        private readonly InventorService _inventorService;
        private System.Windows.Threading.DispatcherTimer? _inventorStatusTimer;
        
        // Suivi du temps pour la progression
        private DateTime _startTime;
        private TimeSpan _pausedTime = TimeSpan.Zero;
        private DateTime _pauseStartTime;

        // Liste des initiales dessinateurs XNRGY (mise à jour 2026-01-09)
        private readonly List<string> _designerInitials = new List<string>
        {
            "N/A", "AC", "AM", "AR", "CC", "DC", "DL", "FL", 
            "IM", "KB", "KJ", "MAE", "MC", "NJ", "RO", "SB", "TG", "TV", "VK", "YS",
            "Autre..."
        };

        /// <summary>
        /// Constructeur par défaut (sans vérification admin)
        /// </summary>
        public PlaceEquipmentWindow() : this(null, null)
        {
        }

        /// <summary>
        /// Constructeur avec service Vault pour vérification admin
        /// </summary>
        /// <param name="vaultService">Service Vault connecté (optionnel)</param>
        public PlaceEquipmentWindow(VaultSdkService? vaultService) : this(vaultService, null)
        {
        }

        /// <summary>
        /// Constructeur avec services Vault et Inventor pour héritage du statut de connexion
        /// </summary>
        /// <param name="vaultService">Service Vault connecté (optionnel)</param>
        /// <param name="inventorService">Service Inventor du formulaire principal (optionnel)</param>
        public PlaceEquipmentWindow(VaultSdkService? vaultService, InventorService? inventorService)
        {
            // IMPORTANT: Initialiser _request et _files AVANT InitializeComponent()
            // car les evenements TextChanged du XAML sont declenches pendant l'initialisation
            _request = new CreateModuleRequest();
            _files = new ObservableCollection<FileRenameItem>();
            _vaultService = vaultService;
            // Utiliser le service Inventor du formulaire principal (heritage du statut)
            _inventorService = inventorService ?? new InventorService();
            
            // [+] Initialiser le service Equipment avec callback de log
            _equipmentService = new EquipmentPlacementService(_vaultService, (msg, level) => AddLog(msg, level));
            
            // [+] Forcer la reconnexion COM a chaque ouverture de Place Equipment
            // Evite les problemes de connexion COM obsolete
            _inventorService.ForceReconnect();
            
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                XnrgyMessageBox.ShowError(
                    $"Erreur lors de l'initialisation de la fenetre Place Equipment:\n\n{ex.Message}\n\nDetails:\n{ex.StackTrace}",
                    "Erreur d'initialisation");
                throw;
            }
            
            // S'abonner aux changements de theme
            MainWindow.ThemeChanged += OnThemeChanged;
            
            // Appliquer le theme actuel au demarrage
            ApplyTheme(MainWindow.CurrentThemeIsDark);
            
            // Attendre que la fenetre soit chargee pour initialiser les controles
            this.Loaded += PlaceEquipmentWindow_Loaded;
            this.Closed += (s, e) =>
            {
                MainWindow.ThemeChanged -= OnThemeChanged;
                _inventorStatusTimer?.Stop();
                // Nettoyer le dossier temporaire Vault si present
                CleanupTempVaultFolder();
            };
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
            // Elements avec fond noir FIXE (ne changent jamais)
            StatisticsBorder.Background = new SolidColorBrush(Color.FromRgb(26, 26, 40)); // #1A1A28 - Header stats
            
            if (isDarkTheme)
            {
                // Theme SOMBRE
                this.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46)); // #1E1E2E
                InputsSectionBorder.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46)); // #1E1E2E
            }
            else
            {
                // Theme CLAIR
                this.Background = new SolidColorBrush(Color.FromRgb(245, 247, 250)); // Bleu-gris tres clair
                InputsSectionBorder.Background = new SolidColorBrush(Color.FromRgb(245, 247, 250)); // Meme fond clair
            }
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
                    FontSize = 14,
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
        private void UpdateProgress(int percent, string statusText, bool isError = false, string currentFile = "")
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
                }

                // Gradient brillant et cristallisé selon l'état
                LinearGradientBrush gradientBrush;
                if (isError)
                {
                    // Rouge brillant pour erreur
                    gradientBrush = new LinearGradientBrush
                    {
                        StartPoint = new System.Windows.Point(0, 0),
                        EndPoint = new System.Windows.Point(1, 0)
                    };
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#FF4444"), 0));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#FF6B6B"), 0.5));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#FF4444"), 1));
                }
                else if (percent >= 100)
                {
                    // Vert néon brillant pour succès
                    gradientBrush = new LinearGradientBrush
                    {
                        StartPoint = new System.Windows.Point(0, 0),
                        EndPoint = new System.Windows.Point(1, 0)
                    };
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FF88"), 0));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FFAA"), 0.3));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FFCC"), 0.5));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FFAA"), 0.7));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FF88"), 1));
                }
                else
                {
                    // Cyan/bleu électrique brillant pour progression
                    gradientBrush = new LinearGradientBrush
                    {
                        StartPoint = new System.Windows.Point(0, 0),
                        EndPoint = new System.Windows.Point(1, 0)
                    };
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FFAA"), 0));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00D4FF"), 0.3));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00BFFF"), 0.5));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00D4FF"), 0.7));
                    gradientBrush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FFAA"), 1));
                }
                ProgressBarFill.Background = gradientBrush;

                // Afficher le fichier actuellement en traitement
                TxtCurrentFile.Text = currentFile ?? "";
                
                // Calcul du temps écoulé et estimé
                TimeSpan elapsed = DateTime.Now - _startTime - _pausedTime;
                TimeSpan? estimatedTotal = null;
                if (percent > 0 && percent < 100)
                {
                    double estimatedSeconds = elapsed.TotalSeconds * 100 / percent;
                    estimatedTotal = TimeSpan.FromSeconds(estimatedSeconds);
                }
                
                // Formatage du temps
                string elapsedStr = FormatTimeSpan(elapsed);
                string estimatedStr = estimatedTotal.HasValue 
                    ? FormatTimeSpan(estimatedTotal.Value)
                    : "00:00";
                    
                TxtProgressTimeElapsed.Text = elapsedStr;
                TxtProgressTimeEstimated.Text = estimatedStr;

                // Mise à jour du texte de statut
                TxtStatus.Text = statusText;
                
                // Mise a jour du pourcentage
                TxtProgressPercent.Text = percent > 0 ? $"{percent}%" : "";
            });
        }

        
        private string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalHours >= 1)
                return $"{(int)ts.TotalHours}h {ts.Minutes:D2}m {ts.Seconds:D2}s";
            else if (ts.TotalMinutes >= 1)
                return $"{ts.Minutes}m {ts.Seconds:D2}s";
            else
                return $"{ts.Seconds}s";
        }

        /// <summary>
        /// Réinitialise la barre de progression
        /// </summary>
        private void ResetProgress()
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBarFill.Width = 0;
                
                // Gradient par défaut (cyan/bleu électrique brillant)
                var defaultGradient = new LinearGradientBrush
                {
                    StartPoint = new System.Windows.Point(0, 0),
                    EndPoint = new System.Windows.Point(1, 0)
                };
                defaultGradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FFAA"), 0));
                defaultGradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00D4FF"), 0.3));
                defaultGradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00BFFF"), 0.5));
                defaultGradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00D4FF"), 0.7));
                defaultGradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00FFAA"), 1));
                ProgressBarFill.Background = defaultGradient;
                
                TxtStatus.Text = "Pret - Selectionnez un equipement";
                TxtProgressPercent.Text = "";
                TxtProgressTimeElapsed.Text = "00:00";
                TxtProgressTimeEstimated.Text = "00:00";
                TxtCurrentFile.Text = "";
            });
        }

        #endregion

        private void PlaceEquipmentWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                DgFiles.ItemsSource = _files;
                
                // Message de bienvenue dans le journal
                AddLog("Fenetre Place Equipment initialisee", "START");
                
                // Verifier si l'utilisateur est administrateur Vault
                CheckVaultAdminPermissions();
                
                // Initialiser le statut Inventor
                UpdateInventorStatus();
                
                // Timer pour mettre a jour le statut Inventor periodiquement
                _inventorStatusTimer = new System.Windows.Threading.DispatcherTimer();
                _inventorStatusTimer.Interval = TimeSpan.FromSeconds(3);
                _inventorStatusTimer.Tick += (s, args) => UpdateInventorStatus();
                _inventorStatusTimer.Start();
                
                // [+] Charger la liste des equipements depuis Vault
                LoadEquipmentListFromVault();
                
                // Initialiser les ComboBox Reference et Module (01-50)
                InitializeReferenceModuleComboBoxes();
                AddLog("ComboBox Reference/Module chargees (01-50)", "INFO");
                
                // Initialiser les ComboBox Initiales Dessinateurs
                InitializeDesignerComboBoxes();
                AddLog($"Liste des dessinateurs chargee ({_designerInitials.Count} initiales)", "INFO");
                
                // Initialiser le ComboBox Instance Equipement (1er par defaut)
                CmbEquipmentInstance.SelectedIndex = 0;
                AddLog("ComboBox Instance Equipement initialisee (1er par defaut)", "INFO");
                
                // Initialiser la date de creation avec DatePicker
                var today = DateTime.Now;
                DpCreationDate.SelectedDate = today;
                _request.CreationDate = today;
                
                // Ne pas charger automatiquement de template - l'utilisateur doit selectionner un equipement
                AddLog("Selectionnez un equipement depuis Vault pour commencer", "INFO");
                TxtStatus.Text = "Pret - Selectionnez un equipement depuis Vault";
                
                UpdateDestinationPreview();
                UpdateStatistics();
            }
            catch (Exception ex)
            {
                AddLog($"Erreur d'initialisation: {ex.Message}", "ERROR");
            }
        }

        #region Vault Equipment Loading

        /// <summary>
        /// Classe pour representer un equipement dans Vault
        /// </summary>
        private class VaultEquipment
        {
            public string Name { get; set; } = string.Empty;
            public string DisplayName { get; set; } = string.Empty;
            public string VaultPath { get; set; } = string.Empty;
            public string LocalPath { get; set; } = string.Empty;
            public string ProjectFileName { get; set; } = string.Empty;
            public string AssemblyFileName { get; set; } = string.Empty;
            
            /// <summary>
            /// Override ToString pour afficher le DisplayName dans le ComboBox
            /// </summary>
            public override string ToString() => DisplayName;
        }

        /// <summary>
        /// Charge la liste des equipements depuis Vault ($/Engineering/Library/Equipment)
        /// </summary>
        private void LoadEquipmentListFromVault()
        {
            if (CmbEquipment == null) return;
            
            CmbEquipment.Items.Clear();
            
            if (_vaultService == null || !_vaultService.IsConnected)
            {
                TxtStatus.Text = "[!] Non connecte a Vault - Impossible de charger les equipements";
                AddLog("[-] Non connecte a Vault", "WARN");
                return;
            }

            try
            {
                TxtStatus.Text = "Chargement des equipements depuis Vault...";
                AddLog("[>] Chargement des equipements depuis Vault...", "INFO");
                
                var equipmentBasePath = "$/Engineering/Library/Equipment";
                
                var connection = _vaultService.Connection;
                if (connection == null)
                {
                    TxtStatus.Text = "[!] Connexion Vault non disponible";
                    return;
                }

                // Obtenir le dossier Equipment
                var equipmentFolder = connection.WebServiceManager.DocumentService.GetFolderByPath(equipmentBasePath);
                if (equipmentFolder == null)
                {
                    TxtStatus.Text = "[!] Dossier Equipment non trouve dans Vault";
                    AddLog($"[-] Dossier non trouve: {equipmentBasePath}", "ERROR");
                    return;
                }

                // Obtenir tous les sous-dossiers (chaque sous-dossier = un equipement)
                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(equipmentFolder.Id, false);
                if (subFolders == null || subFolders.Length == 0)
                {
                    TxtStatus.Text = "[!] Aucun equipement trouve dans Vault";
                    AddLog("[-] Aucun sous-dossier dans Equipment", "WARN");
                    return;
                }

                var equipments = new List<VaultEquipment>();
                foreach (var folder in subFolders)
                {
                    // Chercher les infos IPJ/IAM dans la liste AvailableEquipment
                    var knownEquipment = EquipmentPlacementService.AvailableEquipment
                        .FirstOrDefault(e => e.Name.Equals(folder.Name, StringComparison.OrdinalIgnoreCase) 
                                          || e.VaultPath.EndsWith("/" + folder.Name, StringComparison.OrdinalIgnoreCase));
                    
                    // [DEBUG] Log si equipement connu trouve ou non
                    if (knownEquipment != null)
                    {
                        Logger.Debug($"[+] Equipement '{folder.Name}' trouve dans liste: IPJ='{knownEquipment.ProjectFileName}', IAM='{knownEquipment.AssemblyFileName}'");
                    }
                    else
                    {
                        Logger.Warning($"[!] Equipement '{folder.Name}' non trouve dans AvailableEquipment - IPJ/IAM seront vides");
                    }
                    
                    var equipment = new VaultEquipment
                    {
                        Name = folder.Name,
                        DisplayName = knownEquipment?.DisplayName ?? folder.Name.Replace("_", " "),
                        VaultPath = $"{equipmentBasePath}/{folder.Name}",
                        LocalPath = Path.Combine(_defaultTemplatePath, folder.Name),
                        ProjectFileName = knownEquipment?.ProjectFileName ?? "",
                        AssemblyFileName = knownEquipment?.AssemblyFileName ?? ""
                    };
                    equipments.Add(equipment);
                }
                
                // Trier par nom
                equipments = equipments.OrderBy(e => e.DisplayName).ToList();
                
                foreach (var equipment in equipments)
                {
                    CmbEquipment.Items.Add(equipment);
                }
                
                TxtStatus.Text = $"[+] {CmbEquipment.Items.Count} equipements trouves dans Vault";
                AddLog($"[+] {CmbEquipment.Items.Count} equipements charges depuis Vault", "SUCCESS");
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"[!] Erreur chargement equipements: {ex.Message}";
                AddLog($"[-] Erreur chargement equipements Vault: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// Bouton Actualiser la liste des equipements
        /// </summary>
        private void BtnRefreshEquipments_Click(object sender, RoutedEventArgs e)
        {
            LoadEquipmentListFromVault();
        }

        /// <summary>
        /// Gestionnaire de selection d'equipement
        /// </summary>
        private void CmbEquipment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbEquipment.SelectedItem is VaultEquipment equipment)
            {
                _selectedEquipment = new EquipmentItem
                {
                    Name = equipment.Name,
                    DisplayName = equipment.DisplayName,
                    VaultPath = equipment.VaultPath,
                    LocalTempPath = equipment.LocalPath,
                    ProjectFileName = equipment.ProjectFileName,
                    AssemblyFileName = equipment.AssemblyFileName
                };
                
                // Afficher le chemin Vault
                TxtEquipmentVaultPath.Text = equipment.VaultPath;
                TxtSourcePath.Text = equipment.LocalPath;
                
                AddLog($"[+] Equipement selectionne: {equipment.DisplayName}", "INFO");
                AddLog($"    Vault: {equipment.VaultPath}", "INFO");
                AddLog($"    IPJ: {equipment.ProjectFileName}, IAM: {equipment.AssemblyFileName}", "DEBUG");
                TxtStatus.Text = $"[+] Equipement selectionne: {equipment.DisplayName} - Cliquez 'Charger depuis Vault'";
                
                // Vider les fichiers - ils seront charges apres le telechargement
                _files.Clear();
                UpdateStatistics();
                UpdateDestinationPreview();
                
                // [+] Detecter les instances existantes dans le module cible
                DetectExistingEquipmentInstances();
            }
        }

        /// <summary>
        /// Gestionnaire de selection d'instance d'equipement (1er, 2e, 3e, 4e, Sans suffixe)
        /// </summary>
        private void CmbEquipmentInstance_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbEquipmentInstance.SelectedItem is ComboBoxItem selectedItem)
            {
                // Tag peut etre null ou vide pour "Sans suffixe"
                _selectedInstanceSuffix = selectedItem.Tag?.ToString() ?? "";
                
                if (string.IsNullOrEmpty(_selectedInstanceSuffix))
                {
                    AddLog($"[i] Instance selectionnee: {selectedItem.Content} (pas de suffixe)", "INFO");
                }
                else
                {
                    AddLog($"[i] Instance selectionnee: {selectedItem.Content} (suffixe: {_selectedInstanceSuffix})", "INFO");
                }
                
                UpdateDestinationPreview();
                
                // [+] Mettre a jour les noms des fichiers dans le DataGrid avec le nouveau suffixe
                ApplyInstanceSuffixToFileNames();
            }
        }

        /// <summary>
        /// Applique le suffixe d'instance (_01, _02, etc.) aux noms des fichiers dans le DataGrid
        /// </summary>
        private void ApplyInstanceSuffixToFileNames()
        {
            if (_files == null || _files.Count == 0) return;
            
            foreach (var file in _files)
            {
                // Recuperer le nom original sans extension (utiliser OriginalFileName, pas OriginalName)
                var originalNameWithoutExt = Path.GetFileNameWithoutExtension(file.OriginalFileName);
                var extension = Path.GetExtension(file.OriginalFileName);
                
                // Supprimer tout suffixe existant (_01, _02, _03, _04)
                var cleanName = RemoveInstanceSuffix(originalNameWithoutExt);
                
                // Appliquer le nouveau suffixe (ou rien si "Sans suffixe")
                if (!string.IsNullOrEmpty(_selectedInstanceSuffix))
                {
                    file.NewFileName = $"{cleanName}{_selectedInstanceSuffix}{extension}";
                }
                else
                {
                    file.NewFileName = $"{cleanName}{extension}";
                }
            }
            
            // Rafraichir le DataGrid
            DgFiles?.Items.Refresh();
            
            AddLog($"[+] Noms mis a jour avec suffixe: {(_selectedInstanceSuffix == "" ? "(aucun)" : _selectedInstanceSuffix)}", "INFO");
        }

        /// <summary>
        /// Supprime les suffixes d'instance existants (_01, _02, _03, _04) d'un nom de fichier
        /// </summary>
        private string RemoveInstanceSuffix(string name)
        {
            var suffixes = new[] { "_01", "_02", "_03", "_04" };
            foreach (var suffix in suffixes)
            {
                if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    return name.Substring(0, name.Length - suffix.Length);
                }
            }
            return name;
        }

        /// <summary>
        /// Detecte les instances d'equipement existantes dans le dossier 1-Equipment du module cible
        /// et selectionne automatiquement la prochaine instance disponible
        /// </summary>
        private void DetectExistingEquipmentInstances()
        {
            try
            {
                if (_selectedEquipment == null)
                {
                    TxtInstancesDetected.Text = "";
                    return;
                }
                
                // Construire le chemin du dossier 1-Equipment du module cible
                var project = TxtProject?.Text?.Trim() ?? "";
                var reference = CmbReference?.SelectedItem?.ToString() ?? "";
                var module = CmbModule?.SelectedItem?.ToString() ?? "";
                
                if (string.IsNullOrEmpty(project) || string.IsNullOrEmpty(reference) || string.IsNullOrEmpty(module))
                {
                    TxtInstancesDetected.Text = "Selectionnez un module cible pour detecter les instances existantes";
                    return;
                }
                
                var modulePath = Path.Combine(_projectsBasePath, project, $"REF{reference}", $"M{module}");
                var equipmentBasePath = Path.Combine(modulePath, "1-Equipment");
                
                if (!Directory.Exists(equipmentBasePath))
                {
                    TxtInstancesDetected.Text = "Dossier 1-Equipment inexistant - 1er equipement par defaut";
                    CmbEquipmentInstance.SelectedIndex = 0; // 1er equipement (_01)
                    return;
                }
                
                // Chercher les dossiers existants avec le nom de l'equipement
                var equipmentName = _selectedEquipment.Name;
                var existingDirs = Directory.GetDirectories(equipmentBasePath)
                    .Select(d => Path.GetFileName(d))
                    .Where(name => name.StartsWith(equipmentName, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                
                if (existingDirs.Count == 0)
                {
                    TxtInstancesDetected.Text = $"Aucun '{equipmentName}' existant - 1er equipement par defaut";
                    CmbEquipmentInstance.SelectedIndex = 0; // 1er equipement (_01)
                    return;
                }
                
                // Identifier les suffixes existants
                var existingSuffixes = new HashSet<string>();
                foreach (var dir in existingDirs)
                {
                    // Ex: "Angular Filter_01" -> "_01"
                    if (dir.Length > equipmentName.Length)
                    {
                        var suffix = dir.Substring(equipmentName.Length);
                        existingSuffixes.Add(suffix);
                    }
                    else if (dir.Equals(equipmentName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Dossier sans suffixe = considere comme _01
                        existingSuffixes.Add("_01");
                    }
                }
                
                // Determiner la prochaine instance disponible
                var allSuffixes = new[] { "_01", "_02", "_03", "_04" };
                var nextIndex = 0;
                for (int i = 0; i < allSuffixes.Length; i++)
                {
                    if (!existingSuffixes.Contains(allSuffixes[i]))
                    {
                        nextIndex = i;
                        break;
                    }
                    // Si tous sont pris, prendre le dernier (ecrasement)
                    nextIndex = allSuffixes.Length - 1;
                }
                
                // Afficher les instances detectees
                TxtInstancesDetected.Text = $"Existants: {string.Join(", ", existingDirs)} - Suggestion: {(nextIndex + 1)}e";
                
                // Selectionner automatiquement la prochaine instance
                CmbEquipmentInstance.SelectedIndex = nextIndex;
                
                AddLog($"[i] Instances detectees: {string.Join(", ", existingDirs)}", "DEBUG");
                AddLog($"[i] Prochaine instance suggeree: {allSuffixes[nextIndex]}", "DEBUG");
            }
            catch (Exception ex)
            {
                Logger.Warning($"[!] Erreur detection instances: {ex.Message}");
                TxtInstancesDetected.Text = "Erreur detection - 1er equipement par defaut";
                CmbEquipmentInstance.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Bouton Charger depuis Vault - Telecharge l'equipement et charge les fichiers
        /// </summary>
        private async void BtnLoadFromVault_Click(object sender, RoutedEventArgs e)
        {
            if (CmbEquipment.SelectedItem is not VaultEquipment equipment)
            {
                TxtStatus.Text = "[!] Veuillez selectionner un equipement";
                XnrgyMessageBox.ShowWarning(
                    "Veuillez d'abord selectionner un equipement dans la liste.",
                    "Equipement requis",
                    this);
                return;
            }

            if (_vaultService == null || !_vaultService.IsConnected)
            {
                TxtStatus.Text = "[!] Non connecte a Vault";
                XnrgyMessageBox.ShowWarning(
                    "Vous n'etes pas connecte a Vault.\nConnectez-vous d'abord depuis la fenetre principale.",
                    "Connexion Vault requise",
                    this);
                return;
            }

            await DownloadEquipmentFromVault(equipment);
        }

        /// <summary>
        /// Telecharge un equipement depuis Vault vers le dossier local
        /// Utilise la barre de progression principale (0-100%)
        /// </summary>
        private async Task DownloadEquipmentFromVault(VaultEquipment equipment)
        {
            try
            {
                Logger.Info("═══════════════════════════════════════════════════════");
                Logger.Info("[>] TELECHARGEMENT EQUIPEMENT DEPUIS VAULT");
                Logger.Info("═══════════════════════════════════════════════════════");
                
                // Desactiver les controles pendant le telechargement
                BtnLoadFromVault.IsEnabled = false;
                CmbEquipment.IsEnabled = false;
                
                // [+] RESET complet de la progression (temps a 00:00)
                ResetProgress();
                _startTime = DateTime.Now;
                _pausedTime = TimeSpan.Zero;
                
                // Initialiser la progression principale
                UpdateProgress(0, $"Preparation du telechargement de {equipment.DisplayName}...");
                
                AddLog($"[>] Demarrage telechargement: {equipment.DisplayName}", "INFO");
                Logger.Info($"[>] Equipment: {equipment.DisplayName}");
                Logger.Info($"[i] Source Vault: {equipment.VaultPath}");
                Logger.Info($"[i] Destination locale: {equipment.LocalPath}");

                // ══════════════════════════════════════════════════════════════════
                // ETAPE 0: NETTOYAGE COMPLET du dossier Equipment AVANT TOUT
                // Nettoyer C:\Vault\Engineering\Library\Equipment AU COMPLET
                // ══════════════════════════════════════════════════════════════════
                UpdateProgress(1, "NETTOYAGE COMPLET du dossier Equipment...");
                AddLog("[>] NETTOYAGE COMPLET de C:\\Vault\\Engineering\\Library\\Equipment...", "INFO");
                Logger.Info("[>] NETTOYAGE COMPLET du dossier Equipment de base");
                
                bool cleanSuccess = CleanEntireEquipmentFolder();
                if (!cleanSuccess)
                {
                    AddLog("[!] Nettoyage partiel - certains fichiers n'ont pas pu etre supprimes", "WARN");
                    Logger.Warning("[!] Nettoyage partiel du dossier Equipment");
                }
                else
                {
                    AddLog("[+] Dossier Equipment nettoye completement", "SUCCESS");
                    Logger.Info("[+] Dossier Equipment nettoye completement");
                }

                // Creer le dossier de destination
                Directory.CreateDirectory(equipment.LocalPath);

                var connection = _vaultService.Connection;
                if (connection == null)
                {
                    throw new Exception("Connexion Vault perdue");
                }

                // Obtenir le dossier Vault (5%)
                UpdateProgress(5, "Connexion au dossier Vault...");
                
                var vaultFolder = connection.WebServiceManager.DocumentService.GetFolderByPath(equipment.VaultPath);
                if (vaultFolder == null)
                {
                    throw new Exception($"Dossier Vault introuvable: {equipment.VaultPath}");
                }
                Logger.Info($"[+] Dossier Vault trouve: {vaultFolder.FullName} (ID: {vaultFolder.Id})");

                // Obtenir TOUS les fichiers RECURSIVEMENT (10%)
                UpdateProgress(10, "Enumeration recursive des fichiers...");
                AddLog($"[>] Enumeration recursive depuis: {equipment.VaultPath}", "INFO");
                Logger.Info($"[>] Enumeration recursive depuis: {equipment.VaultPath}");
                
                var allFiles = new List<ACW.File>();
                var allFolders = new List<ACW.Folder>();
                await Task.Run(() => GetAllFilesRecursiveForEquipment(connection, vaultFolder, allFiles, allFolders));
                
                AddLog($"[+] {allFolders.Count} dossiers trouves", "INFO");
                Logger.Info($"[+] {allFolders.Count} dossiers trouves");
                foreach (var folder in allFolders.Take(10))
                {
                    Logger.Debug($"    [i] {folder.FullName}");
                }
                if (allFolders.Count > 10)
                {
                    Logger.Debug($"    ... et {allFolders.Count - 10} autres dossiers");
                }
                
                if (allFiles.Count == 0)
                {
                    UpdateProgress(0, "[!] Aucun fichier trouve dans l'equipement", isError: true);
                    AddLog("[-] Aucun fichier trouve (meme recursivement)", "ERROR");
                    Logger.Error("[-] Aucun fichier trouve (meme recursivement)");
                    return;
                }
                
                int totalFiles = allFiles.Count;
                AddLog($"[+] {totalFiles} fichiers trouves au total", "SUCCESS");
                Logger.Info($"[+] {totalFiles} fichiers trouves au total (recursif)");

                // Obtenir le working folder
                var workingFolderObj = connection.WorkingFoldersManager.GetWorkingFolder("$");
                if (workingFolderObj == null || string.IsNullOrEmpty(workingFolderObj.FullPath))
                {
                    throw new Exception("Working folder non configure dans Vault");
                }

                var workingFolder = workingFolderObj.FullPath;
                var relativePath = equipment.VaultPath.TrimStart('$', '/').Replace("/", "\\");
                var localFolder = Path.Combine(workingFolder, relativePath);
                Logger.Debug($"[i] Working folder: {workingFolder}");
                Logger.Debug($"[i] Chemin local Vault: {localFolder}");

                // Preparer le telechargement batch (15%)
                UpdateProgress(15, $"Preparation de {totalFiles} fichiers...");
                var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection, false);
                
                int fileIndex = 0;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                
                foreach (var file in allFiles)
                {
                    try
                    {
                        var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(connection, file);
                        downloadSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                        fileIndex++;
                        
                        // Mise a jour progression pendant la preparation (15-20%)
                        if (fileIndex % 10 == 0 || fileIndex == totalFiles)
                        {
                            int prepProgress = 15 + (int)((fileIndex / (double)totalFiles) * 5);
                            UpdateProgress(prepProgress, $"Preparation {fileIndex}/{totalFiles} fichiers...", currentFile: file.Name);
                        }
                    }
                    catch (Exception fileEx)
                    {
                        AddLog($"[!] Erreur preparation {file.Name}: {fileEx.Message}", "WARN");
                        Logger.Warning($"[!] Erreur preparation {file.Name}: {fileEx.Message}");
                    }
                }
                sw.Stop();
                AddLog($"[+] {fileIndex} fichiers prepares en {sw.ElapsedMilliseconds}ms", "INFO");
                Logger.Info($"[+] {fileIndex} fichiers prepares en {sw.ElapsedMilliseconds}ms");

                // ══════════════════════════════════════════════════════════════════
                // TELECHARGEMENT PAR LOTS avec progression precise (20-70%)
                // Telecharge par lots de 10 fichiers pour permettre mise a jour UI
                // ══════════════════════════════════════════════════════════════════
                const int BATCH_SIZE = 10;
                int downloadedCount = 0;
                int totalToDownload = allFiles.Count;
                var allDownloadResults = new List<VDF.Vault.Results.FileAcquisitionResult>();
                
                UpdateProgress(20, $"Telechargement de {totalToDownload} fichiers...");
                AddLog($"[>] Lancement telechargement par lots de {BATCH_SIZE}...", "INFO");
                Logger.Info($"[>] Lancement telechargement par lots de {BATCH_SIZE} fichiers...");
                
                sw.Restart();
                
                // Diviser en lots pour mise a jour progressive
                var fileBatches = allFiles
                    .Select((file, index) => new { file, index })
                    .GroupBy(x => x.index / BATCH_SIZE)
                    .Select(g => g.Select(x => x.file).ToList())
                    .ToList();
                
                foreach (var batch in fileBatches)
                {
                    try
                    {
                        // Preparer le lot
                        var batchSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection, false);
                        foreach (var file in batch)
                        {
                            var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(connection, file);
                            batchSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                        }
                        
                        // Telecharger le lot
                        var batchResult = await Task.Run(() => connection.FileManager.AcquireFiles(batchSettings));
                        
                        if (batchResult?.FileResults != null)
                        {
                            allDownloadResults.AddRange(batchResult.FileResults);
                        }
                        
                        downloadedCount += batch.Count;
                        
                        // Mise a jour progression (20-70% = 50% de plage)
                        int progressPercent = 20 + (int)((downloadedCount / (double)totalToDownload) * 50);
                        string currentFileName = batch.LastOrDefault()?.Name ?? "";
                        UpdateProgress(progressPercent, $"Telechargement {downloadedCount}/{totalToDownload}...", currentFile: currentFileName);
                        
                        // Permettre a l'UI de se rafraichir
                        await Task.Delay(10);
                    }
                    catch (Exception batchEx)
                    {
                        Logger.Warning($"[!] Erreur lot: {batchEx.Message}");
                        downloadedCount += batch.Count; // Continuer meme en cas d'erreur
                    }
                }
                
                sw.Stop();

                if (allDownloadResults.Count == 0)
                {
                    UpdateProgress(0, "[!] Aucun fichier telecharge", isError: true);
                    AddLog("[-] AcquireFiles n'a retourne aucun resultat", "ERROR");
                    Logger.Error("[-] AcquireFiles n'a retourne aucun resultat");
                    return;
                }

                var fileResultsList = allDownloadResults;
                int successCount = fileResultsList.Count(r => r.LocalPath?.FullPath != null && File.Exists(r.LocalPath.FullPath));
                
                // Progression apres telechargement batch (70%)
                UpdateProgress(70, $"[+] {successCount}/{fileResultsList.Count} fichiers telecharges");
                AddLog($"[+] {successCount}/{fileResultsList.Count} fichiers telecharges en {sw.ElapsedMilliseconds}ms", "SUCCESS");
                Logger.Info($"[+] {successCount}/{fileResultsList.Count} fichiers telecharges en {sw.ElapsedMilliseconds}ms ({sw.ElapsedMilliseconds / Math.Max(1, successCount)}ms/fichier)");

                // [+] PAS DE COPIE NECESSAIRE - Vault telecharge directement dans le working folder
                // qui est deja equipment.LocalPath (C:\Vault\Engineering\Library\Equipment\XXX)
                // La copie precedente causait des erreurs "Access denied" car les fichiers etaient deja la
                
                // Verifier si des fichiers ont ete telecharges (90%)
                UpdateProgress(90, "Verification des fichiers telecharges...");
                
                // Enlever les attributs ReadOnly sur les fichiers telecharges pour pouvoir les utiliser
                RemoveReadOnlyAttributes(equipment.LocalPath);
                
                var downloadedFiles = Directory.GetFiles(equipment.LocalPath, "*.*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\_V\\"))  // Exclure les fichiers de version _V
                    .ToArray();
                AddLog($"[+] {downloadedFiles.Length} fichiers telecharges dans le dossier", "SUCCESS");
                Logger.Info($"[+] {downloadedFiles.Length} fichiers dans le dossier destination (excl. _V)");

                // Charger les fichiers depuis le dossier local (95%)
                UpdateProgress(95, "Chargement de la liste des fichiers...");
                await Task.Delay(300);
                LoadFilesFromPath(equipment.LocalPath);
                
                // [+] Appliquer le suffixe d'instance aux noms des fichiers
                ApplyInstanceSuffixToFileNames();
                
                // Termine! (100%)
                UpdateProgress(100, $"[+] {_files.Count} fichiers charges depuis {equipment.DisplayName}");
                AddLog($"[+] Telechargement termine: {equipment.DisplayName}", "SUCCESS");
                AddLog($"[+] {_files.Count} fichiers charges dans la liste", "SUCCESS");
                Logger.Info($"[+] Telechargement termine: {equipment.DisplayName}");
                Logger.Info($"[+] {_files.Count} fichiers charges dans la liste");
                Logger.Info("═══════════════════════════════════════════════════════");
            }
            catch (Exception ex)
            {
                AddLog($"[-] Erreur telechargement: {ex.Message}", "ERROR");
                Logger.Error($"[-] Erreur telechargement: {ex.Message}");
                Logger.Debug($"    StackTrace: {ex.StackTrace}");
                UpdateProgress(0, $"[-] Erreur: {ex.Message}", isError: true);
                
                XnrgyMessageBox.ShowError(
                    $"Erreur lors du telechargement:\n{ex.Message}",
                    "Erreur de telechargement",
                    this);
            }
            finally
            {
                // Reactiver les controles
                BtnLoadFromVault.IsEnabled = true;
                CmbEquipment.IsEnabled = true;
            }
        }

        /// <summary>
        /// Copie recursivement un dossier avec mise a jour de la progression
        /// NOTE: Cette methode n'est plus utilisee pour le telechargement Vault
        /// car Vault telecharge directement dans le working folder
        /// </summary>
        private async Task CopyDirectoryRecursiveWithProgress(string sourceDir, string destDir, int startProgress, int endProgress)
        {
            // Enumerer tous les fichiers a copier
            var allSourceFiles = Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories);
            int totalFiles = allSourceFiles.Length;
            int copiedCount = 0;
            
            foreach (var sourceFile in allSourceFiles)
            {
                try
                {
                    // Calculer le chemin relatif
                    var relativePath = sourceFile.Substring(sourceDir.Length).TrimStart('\\', '/');
                    var destFile = Path.Combine(destDir, relativePath);
                    var destDirectory = Path.GetDirectoryName(destFile);
                    
                    if (!string.IsNullOrEmpty(destDirectory))
                    {
                        Directory.CreateDirectory(destDirectory);
                    }
                    
                    File.Copy(sourceFile, destFile, true);
                    copiedCount++;
                    
                    // Mise a jour progression
                    if (copiedCount % 5 == 0 || copiedCount == totalFiles)
                    {
                        int progress = startProgress + (int)((copiedCount / (double)totalFiles) * (endProgress - startProgress));
                        UpdateProgress(progress, $"Copie {copiedCount}/{totalFiles}...", currentFile: Path.GetFileName(sourceFile));
                        await Task.Yield(); // Permettre a l'UI de se rafraichir
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"[!] Erreur copie {Path.GetFileName(sourceFile)}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Obtient tous les fichiers recursivement depuis un dossier Vault (pour equipements)
        /// IMPORTANT: Recupere aussi les associations/dependances des fichiers Inventor
        /// Tous les fichiers sont dans le meme dossier (et sous-dossiers) dans Vault
        /// </summary>
        private void GetAllFilesRecursiveForEquipment(VDF.Vault.Currency.Connections.Connection connection, ACW.Folder folder, List<ACW.File> allFiles, List<ACW.Folder> allFolders)
        {
            try
            {
                // Ajouter ce dossier a la liste
                allFolders.Add(folder);
                
                // Obtenir les fichiers de ce dossier
                var files = connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, false);
                if (files != null && files.Length > 0)
                {
                    foreach (var file in files)
                    {
                        // Ajouter le fichier s'il n'est pas deja dans la liste
                        if (!allFiles.Any(f => f.Id == file.Id))
                        {
                            allFiles.Add(file);
                        }
                    }
                }
                
                // Obtenir les sous-dossiers
                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(folder.Id, false);
                if (subFolders != null && subFolders.Length > 0)
                {
                    foreach (var subFolder in subFolders)
                    {
                        // Recursion dans chaque sous-dossier
                        GetAllFilesRecursiveForEquipment(connection, subFolder, allFiles, allFolders);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"[!] Erreur enumeration {folder.FullName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Copie recursivement un dossier vers une destination
        /// </summary>
        private void CopyDirectoryRecursive(string sourceDir, string destDir)
        {
            // Creer le dossier de destination s'il n'existe pas
            Directory.CreateDirectory(destDir);

            // Copier tous les fichiers
            foreach (var file in Directory.GetFiles(sourceDir))
            {
                var destFile = Path.Combine(destDir, Path.GetFileName(file));
                try
                {
                    File.Copy(file, destFile, true);
                }
                catch (Exception ex)
                {
                    Logger.Warning($"[!] Erreur copie {Path.GetFileName(file)}: {ex.Message}");
                }
            }

            // Copier recursivement les sous-dossiers
            foreach (var subDir in Directory.GetDirectories(sourceDir))
            {
                var destSubDir = Path.Combine(destDir, Path.GetFileName(subDir));
                CopyDirectoryRecursive(subDir, destSubDir);
            }
        }

        #endregion

        /// <summary>
        /// Ancienne methode - gardee pour compatibilite mais non utilisee
        /// </summary>
        private void InitializeEquipmentComboBox()
        {
            // Ne fait plus rien - les equipements sont charges depuis Vault
            // via LoadEquipmentListFromVault()
        }

        /// <summary>
        /// Detecte le module actif dans Inventor et pre-remplit les champs
        /// </summary>
        private void BtnDetectModule_Click(object sender, RoutedEventArgs e)
        {
            Logger.Info("═══════════════════════════════════════════════════════");
            Logger.Info("[>] DETECTION MODULE ET CHARGEMENT iProperties");
            Logger.Info("═══════════════════════════════════════════════════════");
            AddLog("[>] Detection du module actif dans Inventor...", "INFO");
            
            try
            {
                if (!_inventorService.IsConnected)
                {
                    Logger.Error("[-] Inventor non connecte - Detection impossible");
                    AddLog("[-] Inventor non connecte - Impossible de detecter le module actif", "ERROR");
                    XnrgyMessageBox.ShowWarning(
                        "Inventor n'est pas connecte.\nVeuillez ouvrir Inventor et un assembly de module.",
                        "Inventor non connecte",
                        this);
                    return;
                }

                Logger.Info("[+] Inventor connecte");

                // Obtenir le chemin du document actif via InventorService
                var docPath = _inventorService.GetActiveDocumentPath();
                if (string.IsNullOrEmpty(docPath))
                {
                    Logger.Warning("[-] Aucun document actif dans Inventor");
                    AddLog("[-] Aucun document actif dans Inventor", "WARN");
                    XnrgyMessageBox.ShowWarning(
                        "Aucun document n'est ouvert dans Inventor.\nVeuillez ouvrir un assembly de module.",
                        "Aucun document actif",
                        this);
                    return;
                }

                Logger.Info($"[i] Document actif: {docPath}");
                AddLog($"[>] Document actif detecte: {docPath}", "INFO");

                // Extraire Project/Reference/Module du chemin
                // Pattern: C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]\...
                var match = System.Text.RegularExpressions.Regex.Match(
                    docPath, 
                    @"Projects\\(\d{5,6})\\REF(\d{2})\\M(\d{2})\\",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                if (match.Success)
                {
                    var project = match.Groups[1].Value;
                    var reference = match.Groups[2].Value;
                    var module = match.Groups[3].Value;

                    Logger.Info($"[+] Pattern reconnu: Projet={project}, REF={reference}, Module={module}");

                    // Pre-remplir les champs
                    TxtProject.Text = project;
                    Logger.Debug($"[i] TxtProject.Text = {project}");
                    
                    // Selectionner la reference dans la ComboBox
                    bool refFound = false;
                    foreach (var item in CmbReference.Items)
                    {
                        if (item.ToString() == reference)
                        {
                            CmbReference.SelectedItem = item;
                            refFound = true;
                            Logger.Debug($"[+] CmbReference selectionne: {reference}");
                            break;
                        }
                    }
                    if (!refFound) Logger.Warning($"[!] Reference '{reference}' non trouvee dans CmbReference");
                    
                    // Selectionner le module dans la ComboBox
                    bool modFound = false;
                    foreach (var item in CmbModule.Items)
                    {
                        if (item.ToString() == module)
                        {
                            CmbModule.SelectedItem = item;
                            modFound = true;
                            Logger.Debug($"[+] CmbModule selectionne: {module}");
                            break;
                        }
                    }
                    if (!modFound) Logger.Warning($"[!] Module '{module}' non trouve dans CmbModule");

                    AddLog($"[+] Module detecte: Projet={project}, REF{reference}, M{module}", "SUCCESS");
                    Logger.Info($"[+] Module detecte avec succes: {project}-REF{reference}-M{module}");
                    TxtStatus.Text = $"[+] Module detecte: {project}-REF{reference}-M{module}";

                    // Essayer de lire les iProperties du document actif
                    Logger.Info("[>] Chargement des iProperties...");
                    TryReadIPropertiesFromActiveDocument();
                }
                else
                {
                    Logger.Warning($"[!] Pattern non reconnu dans le chemin: {docPath}");
                    Logger.Debug("[i] Pattern attendu: ...\\Projects\\[PROJECT]\\REF[XX]\\M[XX]\\...");
                    AddLog($"[!] Impossible d'extraire les infos projet du chemin: {docPath}", "WARN");
                    XnrgyMessageBox.ShowWarning(
                        $"Le chemin du document actif ne correspond pas au pattern attendu:\n{docPath}\n\nPattern attendu: ...\\Projects\\[PROJECT]\\REF[XX]\\M[XX]\\...",
                        "Pattern non reconnu",
                        this);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"[-] Erreur detection module: {ex.Message}");
                Logger.Debug($"    StackTrace: {ex.StackTrace}");
                AddLog($"[-] Erreur detection module: {ex.Message}", "ERROR");
                XnrgyMessageBox.ShowError(
                    $"Erreur lors de la detection du module:\n{ex.Message}",
                    "Erreur",
                    this);
            }
            
            Logger.Info("═══════════════════════════════════════════════════════");
        }

        /// <summary>
        /// Tente de lire les iProperties du document actif pour pre-remplir les champs
        /// </summary>
        private void TryReadIPropertiesFromActiveDocument()
        {
            Logger.Info("[>] Lecture des iProperties depuis le document actif...");
            AddLog("[>] Lecture des iProperties depuis le document actif...", "INFO");
            
            try
            {
                // Obtenir l'application Inventor et le document actif
                var inventorApp = _inventorService.GetInventorApplication();
                if (inventorApp == null)
                {
                    Logger.Warning("[!] Impossible d'obtenir l'application Inventor");
                    AddLog("[!] Impossible d'obtenir l'application Inventor", "WARN");
                    return;
                }
                
                dynamic activeDoc = inventorApp.ActiveDocument;
                if (activeDoc == null)
                {
                    Logger.Warning("[!] Aucun document actif dans Inventor");
                    AddLog("[!] Aucun document actif dans Inventor", "WARN");
                    return;
                }

                string docName = activeDoc.DisplayName;
                Logger.Info($"[i] Document actif: {docName}");
                AddLog($"[i] Document actif: {docName}", "INFO");
                
                // Acceder aux PropertySets du document
                var propertySets = activeDoc.PropertySets;
                Logger.Debug($"[i] PropertySets count: {propertySets.Count}");
                
                // Chercher dans les Custom Properties
                dynamic customProps = null;
                try
                {
                    customProps = propertySets["Inventor User Defined Properties"];
                    Logger.Debug("[+] PropertySet 'Inventor User Defined Properties' trouve");
                }
                catch (Exception ex)
                {
                    Logger.Warning($"[!] PropertySet 'Inventor User Defined Properties' non trouve: {ex.Message}");
                    AddLog("[!] Pas de proprietes custom dans ce document", "WARN");
                }

                if (customProps != null)
                {
                    int propsFound = 0;
                    
                    // Lister toutes les proprietes custom pour debug
                    try
                    {
                        Logger.Debug("[i] Liste des proprietes custom disponibles:");
                        foreach (dynamic prop in customProps)
                        {
                            try
                            {
                                string propName = prop.Name;
                                string propValue = prop.Value?.ToString() ?? "(null)";
                                Logger.Debug($"    - '{propName}' = '{propValue}'");
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[!] Erreur lors de l'enumeration des proprietes: {ex.Message}");
                    }
                    
                    // Chercher "Initiale_du_Dessinateur" (nom avec underscores dans Inventor)
                    try
                    {
                        dynamic prop = customProps["Initiale_du_Dessinateur"];
                        string initDessinateur = prop.Value?.ToString() ?? "";
                        Logger.Info($"[i] Propriete 'Initiale_du_Dessinateur' trouvee: '{initDessinateur}'");
                        
                        if (!string.IsNullOrEmpty(initDessinateur))
                        {
                            bool found = false;
                            foreach (var item in CmbInitialeDessinateur.Items)
                            {
                                if (item.ToString() == initDessinateur)
                                {
                                    CmbInitialeDessinateur.SelectedItem = item;
                                    AddLog($"[+] Initiales Dessinateur: {initDessinateur}", "SUCCESS");
                                    Logger.Info($"[+] Initiales Dessinateur selectionnees: {initDessinateur}");
                                    found = true;
                                    propsFound++;
                                    break;
                                }
                            }
                            if (!found)
                            {
                                Logger.Warning($"[!] Valeur '{initDessinateur}' non trouvee dans la ComboBox CmbInitialeDessinateur");
                                AddLog($"[!] Initiales '{initDessinateur}' non trouvees dans la liste", "WARN");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[i] Propriete 'Initiale_du_Dessinateur' non trouvee: {ex.Message}");
                    }

                    // Chercher "Initiale_du_Co_Dessinateur" (nom avec underscores dans Inventor)
                    try
                    {
                        dynamic prop = customProps["Initiale_du_Co_Dessinateur"];
                        string initCoDessinateur = prop.Value?.ToString() ?? "";
                        Logger.Info($"[i] Propriete 'Initiale_du_Co_Dessinateur' trouvee: '{initCoDessinateur}'");
                        
                        if (!string.IsNullOrEmpty(initCoDessinateur))
                        {
                            bool found = false;
                            foreach (var item in CmbInitialeCoDessinateur.Items)
                            {
                                if (item.ToString() == initCoDessinateur)
                                {
                                    CmbInitialeCoDessinateur.SelectedItem = item;
                                    AddLog($"[+] Initiales Co-Dessinateur: {initCoDessinateur}", "SUCCESS");
                                    Logger.Info($"[+] Initiales Co-Dessinateur selectionnees: {initCoDessinateur}");
                                    found = true;
                                    propsFound++;
                                    break;
                                }
                            }
                            if (!found)
                            {
                                Logger.Warning($"[!] Valeur '{initCoDessinateur}' non trouvee dans la ComboBox CmbInitialeCoDessinateur");
                                AddLog($"[!] Initiales Co '{initCoDessinateur}' non trouvees dans la liste", "WARN");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[i] Propriete 'Initiale_du_Co_Dessinateur' non trouvee: {ex.Message}");
                    }

                    // Chercher "Concepteur Lead CAD" (NOUVEAU)
                    try
                    {
                        dynamic prop = customProps["Concepteur Lead CAD"];
                        string leadCAD = prop.Value?.ToString() ?? "";
                        Logger.Info($"[i] Propriete 'Concepteur Lead CAD' trouvee: '{leadCAD}'");
                        
                        if (!string.IsNullOrEmpty(leadCAD))
                        {
                            bool found = false;
                            foreach (var item in CmbInitialeLeadCAD.Items)
                            {
                                if (item.ToString() == leadCAD)
                                {
                                    CmbInitialeLeadCAD.SelectedItem = item;
                                    AddLog($"[+] Lead CAD: {leadCAD}", "SUCCESS");
                                    Logger.Info($"[+] Lead CAD selectionne: {leadCAD}");
                                    found = true;
                                    propsFound++;
                                    break;
                                }
                            }
                            if (!found)
                            {
                                Logger.Warning($"[!] Valeur '{leadCAD}' non trouvee dans la ComboBox CmbInitialeLeadCAD");
                                AddLog($"[!] Lead CAD '{leadCAD}' non trouve dans la liste", "WARN");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[i] Propriete 'Concepteur Lead CAD' non trouvee: {ex.Message}");
                    }

                    // Chercher "Job_Title" (nom avec underscore dans Inventor)
                    try
                    {
                        dynamic prop = customProps["Job_Title"];
                        string jobTitle = prop.Value?.ToString() ?? "";
                        Logger.Info($"[i] Propriete 'Job_Title' trouvee: '{jobTitle}'");
                        
                        if (!string.IsNullOrEmpty(jobTitle))
                        {
                            TxtJobTitle.Text = jobTitle;
                            AddLog($"[+] Job Title: {jobTitle}", "SUCCESS");
                            Logger.Info($"[+] Job Title applique: {jobTitle}");
                            propsFound++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[i] Propriete 'Job_Title' non trouvee: {ex.Message}");
                    }

                    // Chercher "Creation_Date" (nom avec underscore dans Inventor)
                    try
                    {
                        dynamic prop = customProps["Creation_Date"];
                        var dateValue = prop.Value;
                        Logger.Info($"[i] Propriete 'Creation_Date' trouvee: '{dateValue}' (Type: {dateValue?.GetType().Name ?? "null"})");
                        
                        if (dateValue is DateTime dt)
                        {
                            DpCreationDate.SelectedDate = dt;
                            AddLog($"[+] Date: {dt:yyyy-MM-dd}", "SUCCESS");
                            Logger.Info($"[+] Date appliquee: {dt:yyyy-MM-dd}");
                            propsFound++;
                        }
                        else if (dateValue != null)
                        {
                            // Essayer de parser la date
                            if (DateTime.TryParse(dateValue.ToString(), out DateTime parsedDate))
                            {
                                DpCreationDate.SelectedDate = parsedDate;
                                AddLog($"[+] Date: {parsedDate:yyyy-MM-dd}", "SUCCESS");
                                Logger.Info($"[+] Date parsee et appliquee: {parsedDate:yyyy-MM-dd}");
                                propsFound++;
                            }
                            else
                            {
                                Logger.Warning($"[!] Impossible de parser la date: '{dateValue}'");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[i] Propriete 'Creation_Date' non trouvee: {ex.Message}");
                    }

                    if (propsFound > 0)
                    {
                        AddLog($"[+] {propsFound} iProperties chargees depuis le document actif", "SUCCESS");
                        Logger.Info($"[+] {propsFound} iProperties chargees avec succes");
                    }
                    else
                    {
                        AddLog("[!] Aucune iProperty correspondante trouvee", "WARN");
                        Logger.Warning("[!] Aucune iProperty correspondante trouvee dans le document");
                    }
                }
                else
                {
                    AddLog("[!] Pas de proprietes custom dans ce document", "WARN");
                    Logger.Warning("[!] PropertySet 'Inventor User Defined Properties' est null");
                }
            }
            catch (Exception ex)
            {
                AddLog($"[-] Erreur lecture iProperties: {ex.Message}", "ERROR");
                Logger.Error($"[-] Erreur lors de la lecture des iProperties: {ex.Message}");
                Logger.Debug($"    StackTrace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Vérifie si l'utilisateur connecté à Vault est administrateur
        /// et affiche/masque le bouton Réglages en conséquence
        /// </summary>
        private void CheckVaultAdminPermissions()
        {
            try
            {
                if (_vaultService != null && _vaultService.IsConnected)
                {
                    // Mettre à jour le statut de connexion dans l'en-tête
                    UpdateVaultConnectionStatus(true, _vaultService.VaultName, _vaultService.UserName);
                    
                    _isVaultAdmin = _vaultService.IsCurrentUserAdmin();
                    
                    if (_isVaultAdmin)
                    {
                        BtnSettings.Visibility = Visibility.Visible;
                        AddLog($"[+] Mode Admin activé ({_vaultService.UserName}) - Réglages accessibles", "SUCCESS");
                    }
                    else
                    {
                        BtnSettings.Visibility = Visibility.Collapsed;
                        AddLog($"[i] Utilisateur standard ({_vaultService.UserName}) - Réglages masqués", "INFO");
                    }
                }
                else
                {
                    // Pas de connexion Vault - masquer le bouton par défaut
                    UpdateVaultConnectionStatus(false, null, null);
                    BtnSettings.Visibility = Visibility.Collapsed;
                    AddLog("[i] Non connecté à Vault - Réglages non disponibles", "INFO");
                }
            }
            catch (Exception ex)
            {
                UpdateVaultConnectionStatus(false, null, null);
                BtnSettings.Visibility = Visibility.Collapsed;
                AddLog($"[!] Erreur vérification admin: {ex.Message}", "WARN");
            }
        }

        /// <summary>
        /// Met à jour l'indicateur de connexion Vault dans l'en-tête
        /// </summary>
        private void UpdateVaultConnectionStatus(bool isConnected, string? vaultName, string? userName)
        {
            Dispatcher.Invoke(() =>
            {
                if (VaultStatusIndicator != null)
                {
                    VaultStatusIndicator.Fill = new SolidColorBrush(
                        isConnected ? (Color)ColorConverter.ConvertFromString("#107C10") : (Color)ColorConverter.ConvertFromString("#E81123"));
                }
                
                if (RunVaultName != null && RunUserName != null && RunStatus != null)
                {
                    RunVaultName.Text = isConnected ? $" Vault : {vaultName}  /  " : " Vault : --  /  ";
                    RunUserName.Text = isConnected ? $" Utilisateur : {userName}  /  " : " Utilisateur : --  /  ";
                    RunStatus.Text = isConnected ? " Statut : Connecte" : " Statut : Deconnecte";
                }
            });
        }

        /// <summary>
        /// Met à jour l'indicateur de connexion Inventor dans l'en-tête
        /// </summary>
        private void UpdateInventorStatus()
        {
            Dispatcher.Invoke(() =>
            {
                bool isConnected = _inventorService.IsConnected;
                
                if (InventorStatusIndicator != null)
                {
                    InventorStatusIndicator.Fill = new SolidColorBrush(
                        isConnected ? (Color)ColorConverter.ConvertFromString("#107C10") : (Color)ColorConverter.ConvertFromString("#E81123"));
                }
                
                if (RunInventorStatus != null)
                {
                    RunInventorStatus.Text = isConnected ? " Inventor : Connecte" : " Inventor : Deconnecte";
                }
            });
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
                CmbInitialeLeadCAD.Items.Add(initial);
            }
            
            // Sélectionner N/A par défaut pour le co-dessinateur et Lead CAD
            CmbInitialeDessinateur.SelectedIndex = 0;
            CmbInitialeCoDessinateur.SelectedIndex = 0; // N/A
            CmbInitialeLeadCAD.SelectedIndex = 0; // N/A
        }

        private void CmbReference_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDestinationPreview();
            UpdateFullProjectNumber();
            ValidateForm();
            // [+] Re-detecter les instances quand reference/module change
            DetectExistingEquipmentInstances();
        }

        private void CmbModule_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDestinationPreview();
            UpdateFullProjectNumber();
            ValidateForm();
            // [+] Re-detecter les instances quand reference/module change
            DetectExistingEquipmentInstances();
        }

        private void CmbInitialeDessinateur_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Gérer l'option "Autre..." pour saisie personnalisée
            if (CmbInitialeDessinateur.SelectedItem?.ToString() == "Autre...")
            {
                string customValue = ShowCustomInitialDialog("Dessinateur");
                if (!string.IsNullOrWhiteSpace(customValue))
                {
                    // Ajouter la valeur custom avant "Autre..." si elle n'existe pas déjà
                    if (!CmbInitialeDessinateur.Items.Contains(customValue))
                    {
                        int autreIndex = CmbInitialeDessinateur.Items.IndexOf("Autre...");
                        CmbInitialeDessinateur.Items.Insert(autreIndex, customValue);
                    }
                    CmbInitialeDessinateur.SelectedItem = customValue;
                }
                else
                {
                    // Annulé - revenir à N/A
                    CmbInitialeDessinateur.SelectedIndex = 0;
                }
            }
            ValidateForm();
        }

        private void CmbInitialeCoDessinateur_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Gérer l'option "Autre..." pour saisie personnalisée
            if (CmbInitialeCoDessinateur.SelectedItem?.ToString() == "Autre...")
            {
                string customValue = ShowCustomInitialDialog("Co-Dessinateur");
                if (!string.IsNullOrWhiteSpace(customValue))
                {
                    // Ajouter la valeur custom avant "Autre..." si elle n'existe pas déjà
                    if (!CmbInitialeCoDessinateur.Items.Contains(customValue))
                    {
                        int autreIndex = CmbInitialeCoDessinateur.Items.IndexOf("Autre...");
                        CmbInitialeCoDessinateur.Items.Insert(autreIndex, customValue);
                    }
                    CmbInitialeCoDessinateur.SelectedItem = customValue;
                }
                else
                {
                    // Annulé - revenir à N/A
                    CmbInitialeCoDessinateur.SelectedIndex = 0;
                }
            }
            // Pas de validation requise car optionnel
        }

        private void CmbInitialeLeadCAD_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Gérer l'option "Autre..." pour saisie personnalisée
            if (CmbInitialeLeadCAD.SelectedItem?.ToString() == "Autre...")
            {
                string customValue = ShowCustomInitialDialog("Lead CAD");
                if (!string.IsNullOrWhiteSpace(customValue))
                {
                    // Ajouter la valeur custom avant "Autre..." si elle n'existe pas déjà
                    if (!CmbInitialeLeadCAD.Items.Contains(customValue))
                    {
                        int autreIndex = CmbInitialeLeadCAD.Items.IndexOf("Autre...");
                        CmbInitialeLeadCAD.Items.Insert(autreIndex, customValue);
                    }
                    CmbInitialeLeadCAD.SelectedItem = customValue;
                }
                else
                {
                    // Annulé - revenir à N/A
                    CmbInitialeLeadCAD.SelectedIndex = 0;
                }
            }
            // Pas de validation requise car optionnel
        }

        /// <summary>
        /// Affiche une boîte de dialogue pour saisir des initiales personnalisées
        /// </summary>
        private string ShowCustomInitialDialog(string type)
        {
            // Créer une fenêtre de dialogue simple
            var dialog = new Window
            {
                Title = $"Initiales {type} personnalisées",
                Width = 350,
                Height = 180,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(Color.FromRgb(30, 30, 45)),
                WindowStyle = WindowStyle.ToolWindow
            };

            var stack = new StackPanel { Margin = new Thickness(20) };
            
            var label = new TextBlock
            {
                Text = $"Entrez les initiales du {type} (2-4 caractères):",
                Foreground = Brushes.White,
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 10)
            };
            
            var textBox = new TextBox
            {
                MaxLength = 4,
                FontSize = 14,
                Padding = new Thickness(8, 6, 8, 6),
                Background = new SolidColorBrush(Color.FromRgb(45, 45, 65)),
                Foreground = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(80, 80, 120)),
                CaretBrush = Brushes.White
            };
            textBox.Focus();
            
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 15, 0, 0)
            };
            
            var okButton = new Button
            {
                Content = "OK",
                Width = 80,
                Padding = new Thickness(0, 6, 0, 6),
                Margin = new Thickness(0, 0, 10, 0),
                IsDefault = true
            };
            okButton.Click += (s, e) => { dialog.DialogResult = true; dialog.Close(); };
            
            var cancelButton = new Button
            {
                Content = "Annuler",
                Width = 80,
                Padding = new Thickness(0, 6, 0, 6),
                IsCancel = true
            };
            cancelButton.Click += (s, e) => { dialog.DialogResult = false; dialog.Close(); };
            
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            
            stack.Children.Add(label);
            stack.Children.Add(textBox);
            stack.Children.Add(buttonPanel);
            
            dialog.Content = stack;
            
            if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(textBox.Text))
            {
                return textBox.Text.Trim().ToUpper();
            }
            return null;
        }

        #region Event Handlers - Project Info

        private void ProjectInfo_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateDestinationPreview();
            UpdateFullProjectNumber();
            UpdateRenamePreviews();
            ValidateForm();
        }

        private void TxtJobTitle_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Verifier que la fenetre est chargee avant de valider
            if (!IsLoaded || BtnPlaceEquipment == null) return;
            ValidateForm();
        }

        private void UpdateDestinationPreview()
        {
            if (TxtProject == null || CmbReference == null || CmbModule == null) return;

            var project = TxtProject.Text?.Trim() ?? "";
            var reference = CmbReference.SelectedItem?.ToString() ?? "01";
            var module = CmbModule.SelectedItem?.ToString() ?? "01";
            
            // [+] Pour Place Equipment: destination dans 1-Equipment\[EquipmentName]
            var equipmentName = _selectedEquipment?.Name ?? "[EQUIPMENT]";

            string destPath;
            if (string.IsNullOrEmpty(project))
            {
                destPath = $"{_defaultDestinationBase}\\[PROJET]\\REF{reference}\\M{module}\\1-Equipment\\{equipmentName}";
            }
            else
            {
                destPath = Path.Combine(_defaultDestinationBase, project, $"REF{reference}", $"M{module}", "1-Equipment", equipmentName);
            }

            if (TxtDestinationPath != null)
                TxtDestinationPath.Text = destPath;

            if (TxtDestinationPreview != null)
                TxtDestinationPreview.Text = $"Destination: {destPath}";

            // Mettre a jour les proprietes du request (DestinationPath est calcule automatiquement)
            _request.Project = project;
            _request.Reference = reference;
            _request.Module = module;

            // Mettre a jour les chemins de destination pour chaque fichier
            UpdateFileDestinationPaths(destPath);
        }

        private void UpdateFileDestinationPaths(string destinationBase)
        {
            if (_files == null || _files.Count == 0) return;

            // NOTE: Pour Place Equipment, on ne renomme PAS automatiquement les fichiers .iam et .ipj
            // L'equipement doit garder son nom d'origine (ex: "Angular Filter.iam", "Angular Filter.ipj")
            // L'utilisateur peut renommer manuellement s'il le souhaite via la colonne "Nouveau nom"

            foreach (var file in _files)
            {
                // Calculer le chemin relatif une seule fois
                var relativeDir = Path.GetDirectoryName(file.RelativePath) ?? "";
                var isAtRoot = string.IsNullOrEmpty(relativeDir);
                
                // PAS DE RENOMMAGE AUTOMATIQUE pour Place Equipment
                // Les fichiers .iam et .ipj gardent leur nom d'origine
                // Seuls les fichiers Excel sont renommes automatiquement (via RenameSpecialExcelFiles)
                
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

            // Mettre à jour les propriétés du request (FullProjectNumber sera calculé automatiquement)
            _request.Project = TxtProject?.Text?.Trim() ?? "";
            _request.Reference = reference;
            _request.Module = module;

            // Afficher le numéro complet calculé
            var fullNumber = _request.FullProjectNumber;
            TxtFullProjectNumber.Text = fullNumber;
            
            // Renommer automatiquement les fichiers Excel si le numéro de projet est défini
            if (!string.IsNullOrEmpty(fullNumber) && _files.Count > 0)
            {
                RenameSpecialExcelFiles();
                DgFiles?.Items.Refresh();
            }
            
            // Mettre à jour les prévisualisations
            UpdateRenamePreviews();
        }

        #endregion

        #region Event Handlers - Source Options (Simplified for Equipment)

        /// <summary>
        /// Chemin temporaire pour le telechargement Vault
        /// </summary>
        private string? _tempVaultDownloadPath = null;

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

        private void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);

            foreach (var file in Directory.GetFiles(sourceDir))
            {
                var fileName = Path.GetFileName(file);
                var destFile = Path.Combine(destDir, fileName);
                File.Copy(file, destFile, true);
            }

            foreach (var dir in Directory.GetDirectories(sourceDir))
            {
                var dirName = Path.GetFileName(dir);
                var destSubDir = Path.Combine(destDir, dirName);
                CopyDirectory(dir, destSubDir);
            }
        }

        /// <summary>
        /// Obtient tous les fichiers recursivement depuis un dossier Vault
        /// </summary>
        private void GetAllFilesRecursive(VDF.Vault.Currency.Connections.Connection connection, ACW.Folder folder, List<ACW.File> allFiles, List<ACW.Folder> allFolders)
        {
            try
            {
                // Ajouter ce dossier a la liste
                allFolders.Add(folder);
                
                // Obtenir les fichiers de ce dossier
                var files = connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, false);
                if (files != null && files.Length > 0)
                {
                    allFiles.AddRange(files);
                }
                
                // Obtenir les sous-dossiers
                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(folder.Id, false);
                if (subFolders != null && subFolders.Length > 0)
                {
                    foreach (var subFolder in subFolders)
                    {
                        // Recursion dans chaque sous-dossier
                        GetAllFilesRecursive(connection, subFolder, allFiles, allFolders);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log l'erreur mais continue avec les autres dossiers
                System.Diagnostics.Debug.WriteLine($"[!] Erreur enumeration {folder.FullName}: {ex.Message}");
            }
        }

        private void CleanupTempVaultFolder()
        {
            if (_tempVaultDownloadPath != null && Directory.Exists(_tempVaultDownloadPath))
            {
                try
                {
                    AddLog($"Nettoyage du dossier temporaire: {_tempVaultDownloadPath}", "INFO");
                    Directory.Delete(_tempVaultDownloadPath, true);
                    _tempVaultDownloadPath = null;
                    AddLog("[+] Dossier temporaire supprime", "SUCCESS");
                }
                catch (Exception ex)
                {
                    AddLog($"[!] Impossible de supprimer le dossier temporaire: {ex.Message}", "WARNING");
                }
            }
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
                    if (TxtStatus != null) TxtStatus.Text = "⚠️ Chemin source invalide";
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
                    // Top Assembly detection: Module_.iam (ancien template) ou 000000000.iam (nouveau template)
                    var isTopAssembly = fileName.Equals("Module_.iam", StringComparison.OrdinalIgnoreCase) ||
                                        fileName.Equals("000000000.iam", StringComparison.OrdinalIgnoreCase);
                    var isProjectFile = extension == "IPJ";
                    // Fichier projet principal: pattern XXXXX-XX-XX_2026.ipj ou 000000000.ipj (à la racine du module)
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

                // Mettre à jour les filtres avec les valeurs réellement trouvées
                UpdateFileTypeFilter();
                UpdateStatusFilter();

                // Renommer automatiquement les fichiers Excel spécifiques si le numéro de projet est défini
                if (!string.IsNullOrEmpty(_request.FullProjectNumber))
                {
                    RenameSpecialExcelFiles();
                }

                UpdateStatistics();
                UpdateFileCount();
                UpdateRenamePreviews();
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
                
                // Nouvelles options de renommage
                bool applyProjectPrefix = ChkApplyProjectPrefix?.IsChecked == true;
                string fixedSuffix = "";
                if (CmbFixedSuffix?.SelectedItem is ComboBoxItem fixedSuffixItem && 
                    fixedSuffixItem.Content?.ToString() != "Aucun" && 
                    fixedSuffixItem.Content?.ToString() != "Autre...")
                {
                    fixedSuffix = fixedSuffixItem.Content.ToString();
                }
                bool applyIncrementalSuffix = ChkApplyIncrementalSuffix?.IsChecked == true;
                
                // Si la checkbox n'est pas cochée, ne renommer que les fichiers Inventor
                bool includeNonInventor = ChkIncludeNonInventor?.IsChecked == true;

                // Compteur pour suffixe incrémentatif
                int incrementalCounter = 1;

                // Liste des fichiers sélectionnés pour traitement
                var selectedFiles = _files.Where(f => f.IsSelected).ToList();

                foreach (var file in selectedFiles)
                {
                    // Skip fichiers non-Inventor si checkbox non cochée
                    if (!includeNonInventor && !file.IsInventorFile)
                    {
                        continue;
                    }
                    
                    // STRATÉGIE DE RENOMMAGE CUMULATIVE:
                    // 1. Partir du nom actuel (NewFileName) ou original si pas encore modifié
                    var currentName = file.NewFileName != file.OriginalFileName 
                        ? file.NewFileName 
                        : file.OriginalFileName;
                    
                    var baseName = Path.GetFileNameWithoutExtension(currentName);
                    var ext = Path.GetExtension(currentName);
                    var newName = currentName;

                    // 2. Appliquer préfixe numéro projet si activé (vérifier qu'il n'est pas déjà présent)
                    if (applyProjectPrefix && !string.IsNullOrEmpty(_request.FullProjectNumber))
                    {
                        var projectPrefix = $"{_request.FullProjectNumber}_";
                        if (!baseName.StartsWith(projectPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            baseName = $"{projectPrefix}{baseName}";
                        }
                    }

                    // 3. Appliquer préfixe manuel (vérifier qu'il n'est pas déjà présent)
                    if (!string.IsNullOrEmpty(prefix) && !baseName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        baseName = $"{prefix}{baseName}";
                    }

                    // 4. Appliquer suffixe manuel (vérifier qu'il n'est pas déjà présent)
                    if (!string.IsNullOrEmpty(suffix) && !baseName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                    {
                        baseName = $"{baseName}{suffix}";
                    }

                    // 5. Appliquer suffixe fixe (liste déroulante) - vérifier qu'il n'est pas déjà présent
                    if (!string.IsNullOrEmpty(fixedSuffix) && !baseName.EndsWith(fixedSuffix, StringComparison.OrdinalIgnoreCase))
                    {
                        baseName = $"{baseName}{fixedSuffix}";
                    }

                    // 6. Appliquer suffixe incrémentatif (toujours ajouter car il est unique)
                    if (applyIncrementalSuffix)
                    {
                        // Vérifier si un suffixe incrémentatif existe déjà (format _XX où XX est un nombre)
                        var incrementalPattern = new System.Text.RegularExpressions.Regex(@"_\d{2}$");
                        if (!incrementalPattern.IsMatch(baseName))
                        {
                            baseName = $"{baseName}_{incrementalCounter:D2}";
                            incrementalCounter++;
                        }
                        else
                        {
                            // Remplacer le suffixe incrémentatif existant
                            baseName = incrementalPattern.Replace(baseName, $"_{incrementalCounter:D2}");
                            incrementalCounter++;
                        }
                    }

                    // 7. Reconstruire le nom avec extension
                    newName = $"{baseName}{ext}";

                    // 8. Appliquer rechercher/remplacer (sur le nom complet)
                    if (!string.IsNullOrEmpty(search))
                    {
                        var index = newName.IndexOf(search, StringComparison.OrdinalIgnoreCase);
                        while (index >= 0)
                        {
                            newName = newName.Substring(0, index) + replace + newName.Substring(index + search.Length);
                            index = newName.IndexOf(search, index + replace.Length, StringComparison.OrdinalIgnoreCase);
                        }
                    }

                    file.NewFileName = newName;
                }

                // Renommage automatique des fichiers Excel spécifiques
                RenameSpecialExcelFiles();

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

                // Rafraîchir l'affichage
                DgFiles?.Items.Refresh();

                AddLog("Renommage appliqué aux fichiers sélectionnés", "SUCCESS");
                TxtStatus.Text = "✓ Renommage appliqué";
            }
            catch (Exception ex)
            {
                AddLog($"Erreur lors du renommage: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// Renomme automatiquement les fichiers Excel spécifiques avec le numéro de projet
        /// Recherche les fichiers: XXXXXXXXX_Décompte de DXF_DXF Count.xlsx et XXXXXXXXX_Liste de vérification_Check List.xlsm
        /// </summary>
        private void RenameSpecialExcelFiles()
        {
            if (string.IsNullOrEmpty(_request.FullProjectNumber)) return;

            // Noms des fichiers Excel à renommer automatiquement (sans le préfixe XXXXXXXXX_)
            var excelFileSuffixes = new[]
            {
                "_Décompte de DXF_DXF Count.xlsx",
                "_Liste de vérification_Check List.xlsm"
            };

            foreach (var file in _files)
            {
                var fileName = file.OriginalFileName;
                
                // Vérifier si le fichier correspond à un des patterns Excel
                foreach (var suffix in excelFileSuffixes)
                {
                    // Chercher les fichiers qui se terminent par le suffixe (peu importe le préfixe)
                    if (fileName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                    {
                        // Vérifier si le fichier commence déjà par le numéro de projet
                        if (!file.NewFileName.StartsWith(_request.FullProjectNumber, StringComparison.OrdinalIgnoreCase))
                        {
                            // Renommer avec le numéro de projet
                            file.NewFileName = $"{_request.FullProjectNumber}{suffix}";
                            AddLog($"Fichier Excel renommé automatiquement: {file.OriginalFileName} → {file.NewFileName}", "INFO");
                        }
                        break; // Un seul suffixe peut correspondre
                    }
                }
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
            ChkApplyProjectPrefix.IsChecked = false;
            CmbFixedSuffix.SelectedIndex = 0;
            ChkApplyIncrementalSuffix.IsChecked = false;

            TxtStatus.Text = "✓ Noms réinitialisés";
        }

        private void ChkApplyProjectPrefix_Changed(object sender, RoutedEventArgs e)
        {
            UpdateRenamePreviews();
        }

        private void ChkApplyIncrementalSuffix_Changed(object sender, RoutedEventArgs e)
        {
            UpdateRenamePreviews();
        }

        private void CmbFixedSuffix_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Gérer l'option "Autre..." pour saisie personnalisée
            if (CmbFixedSuffix?.SelectedItem is ComboBoxItem item && item.Content?.ToString() == "Autre...")
            {
                string customValue = ShowCustomSuffixDialog();
                if (!string.IsNullOrWhiteSpace(customValue))
                {
                    // S'assurer que le suffixe commence par _
                    if (!customValue.StartsWith("_"))
                    {
                        customValue = "_" + customValue;
                    }
                    
                    // Ajouter la valeur custom avant "Autre..." si elle n'existe pas déjà
                    if (!CmbFixedSuffix.Items.Cast<ComboBoxItem>().Any(i => i.Content?.ToString() == customValue))
                    {
                        int autreIndex = CmbFixedSuffix.Items.Cast<ComboBoxItem>()
                            .ToList()
                            .FindIndex(i => i.Content?.ToString() == "Autre...");
                        if (autreIndex >= 0)
                        {
                            var newItem = new ComboBoxItem { Content = customValue };
                            CmbFixedSuffix.Items.Insert(autreIndex, newItem);
                        }
                    }
                    CmbFixedSuffix.SelectedItem = CmbFixedSuffix.Items.Cast<ComboBoxItem>()
                        .FirstOrDefault(i => i.Content?.ToString() == customValue);
                }
                else
                {
                    // Annulé - revenir à "Aucun"
                    CmbFixedSuffix.SelectedIndex = 0;
                }
            }
        }

        private string ShowCustomSuffixDialog()
        {
            // Créer une fenêtre de dialogue simple
            var dialog = new Window
            {
                Title = "Suffixe personnalisé",
                Width = 350,
                Height = 180,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(Color.FromRgb(30, 30, 45)),
                WindowStyle = WindowStyle.ToolWindow
            };

            var stack = new StackPanel { Margin = new Thickness(20) };

            var label = new TextBlock
            {
                Text = "Entrez le suffixe personnalisé (ex: _11, _A, _TEST):",
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 15),
                TextWrapping = TextWrapping.Wrap
            };

            var textBox = new TextBox
            {
                FontSize = 14,
                Padding = new Thickness(8),
                Margin = new Thickness(0, 0, 0, 15),
                Background = new SolidColorBrush(Color.FromRgb(45, 45, 60)),
                Foreground = new SolidColorBrush(Colors.White),
                BorderBrush = new SolidColorBrush(Color.FromRgb(74, 127, 191)),
                BorderThickness = new Thickness(2)
            };

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var okButton = new Button
            {
                Content = "OK",
                Width = 80,
                Height = 35,
                Margin = new Thickness(0, 0, 10, 0),
                Background = new SolidColorBrush(Color.FromRgb(74, 127, 191)),
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 13,
                FontWeight = FontWeights.Bold
            };

            var cancelButton = new Button
            {
                Content = "Annuler",
                Width = 80,
                Height = 35,
                Background = new SolidColorBrush(Color.FromRgb(100, 100, 100)),
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 13
            };

            string result = null;

            okButton.Click += (s, args) =>
            {
                result = textBox.Text?.Trim();
                dialog.DialogResult = true;
                dialog.Close();
            };

            cancelButton.Click += (s, args) =>
            {
                dialog.DialogResult = false;
                dialog.Close();
            };

            textBox.KeyDown += (s, args) =>
            {
                if (args.Key == System.Windows.Input.Key.Enter)
                {
                    result = textBox.Text?.Trim();
                    dialog.DialogResult = true;
                    dialog.Close();
                }
            };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            stack.Children.Add(label);
            stack.Children.Add(textBox);
            stack.Children.Add(buttonPanel);

            dialog.Content = stack;
            textBox.Focus();

            dialog.ShowDialog();

            return result;
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

        private void ChkIncludeNonInventor_Changed(object sender, RoutedEventArgs e)
        {
            // La checkbox contrôle le renommage, pas le filtrage
            UpdateRenamePreviews();
        }

        private void UpdateRenamePreviews()
        {
            // Prévisualisation pour "Inclure fichiers non-Inventor"
            if (TxtPreviewInclude != null)
            {
                if (ChkIncludeNonInventor?.IsChecked == true)
                {
                    var nonInventorCount = _files?.Count(f => !f.OriginalFileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) &&
                                                              !f.OriginalFileName.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) &&
                                                              !f.OriginalFileName.EndsWith(".idw", StringComparison.OrdinalIgnoreCase) &&
                                                              !f.OriginalFileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase)) ?? 0;
                    TxtPreviewInclude.Text = nonInventorCount > 0 ? $"({nonInventorCount} fichiers)" : "";
                }
                else
                {
                    TxtPreviewInclude.Text = "";
                }
            }

            // Prévisualisation pour "Préfixe Numéro Projet"
            if (TxtPreviewPrefix != null)
            {
                if (ChkApplyProjectPrefix?.IsChecked == true && !string.IsNullOrEmpty(_request?.FullProjectNumber))
                {
                    var exampleName = _files?.FirstOrDefault()?.OriginalFileName ?? "fichier.exemple";
                    var preview = $"{_request.FullProjectNumber}_{exampleName}";
                    if (preview.Length > 30) preview = preview.Substring(0, 27) + "...";
                    TxtPreviewPrefix.Text = $"→ {preview}";
                }
                else
                {
                    TxtPreviewPrefix.Text = "";
                }
            }

            // Prévisualisation pour "Suffixe incrémentatif"
            if (TxtPreviewIncremental != null)
            {
                if (ChkApplyIncrementalSuffix?.IsChecked == true)
                {
                    var selectedCount = _files?.Count(f => f.IsSelected) ?? 0;
                    if (selectedCount > 0)
                    {
                        var exampleName = _files?.FirstOrDefault(f => f.IsSelected)?.OriginalFileName ?? "fichier.exemple";
                        var nameWithoutExt = Path.GetFileNameWithoutExtension(exampleName);
                        var ext = Path.GetExtension(exampleName);
                        TxtPreviewIncremental.Text = $"→ {nameWithoutExt}_01{ext} ... ({selectedCount} fichiers)";
                    }
                    else
                    {
                        TxtPreviewIncremental.Text = "(aucun fichier sélectionné)";
                    }
                }
                else
                {
                    TxtPreviewIncremental.Text = "";
                }
            }
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

            // Filtre par statut (dynamique)
            if (CmbFilterSelection?.SelectedItem is ComboBoxItem selItem)
            {
                var selection = selItem.Content?.ToString();
                if (!string.IsNullOrEmpty(selection) && selection != "Tous")
                {
                    filteredFiles = filteredFiles.Where(f => f.Status == selection);
                }
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

            // Mettre à jour les statistiques de l'en-tête (sélection)
            if (TxtStatsSelected != null) TxtStatsSelected.Text = selected.ToString();
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

            // Si l'utilisateur confirme, proceder au placement
            if (previewWindow.ShowDialog() == true && previewWindow.IsConfirmed)
            {
                AddLog("Placement confirme via previsualisation", "START");
                ExecuteEquipmentPlacement();
            }
            else
            {
                AddLog("Placement annule par l'utilisateur", "INFO");
            }
        }

        private void BtnPlaceEquipment_Click(object sender, RoutedEventArgs e)
        {
            // Verifier qu'un equipement est selectionne
            if (_selectedEquipment == null)
            {
                AddLog("[-] Aucun equipement selectionne", "ERROR");
                return;
            }

            // Verifier que le projet/reference/module sont definis
            var project = TxtProject.Text?.Trim();
            var reference = CmbReference?.SelectedItem?.ToString() ?? "";
            var module = CmbModule?.SelectedItem?.ToString() ?? "";
            
            if (string.IsNullOrEmpty(project) || string.IsNullOrEmpty(reference) || string.IsNullOrEmpty(module))
            {
                AddLog("[-] Project/Reference/Module requis - Cliquez sur 'Detecter le Module'", "ERROR");
                return;
            }

            AddLog($"[>] Demarrage placement equipement: {_selectedEquipment.DisplayName}", "START");
            AddLog($"[i] Instance: {_selectedInstanceSuffix}", "INFO");
            AddLog($"[i] Destination: {project}/REF{reference}/M{module}/1-Equipment/{_selectedEquipment.Name}{_selectedInstanceSuffix}", "INFO");
            
            ExecuteEquipmentPlacement();
        }

        /// <summary>
        /// Execute le placement d'equipement complet:
        /// FLUX CORRECT:
        /// 1. Switch vers IPJ de l'equipement source
        /// 2. Copy Design de l'equipement vers 1-Equipment
        /// 3. Localiser l'assemblage copie
        /// 4. Switch vers IPJ du module destination
        /// 5. Ouvrir l'assemblage copie avec l'IPJ destination
        /// 6. Appliquer les iProperties
        /// 7. Sauvegarder et fermer l'assemblage equipement
        /// 8. Ouvrir l'assemblage principal du module et inserer l'equipement
        /// </summary>
        private async void ExecuteEquipmentPlacement()
        {
            if (_selectedEquipment == null)
            {
                AddLog("[-] Aucun equipement selectionne", "ERROR");
                return;
            }
            
            // [+] Verifier que les fichiers IPJ et IAM sont definis
            if (string.IsNullOrEmpty(_selectedEquipment.ProjectFileName))
            {
                AddLog($"[-] Fichier IPJ non defini pour '{_selectedEquipment.DisplayName}'", "ERROR");
                XnrgyMessageBox.ShowError(
                    $"Le fichier projet (.ipj) n'est pas defini pour '{_selectedEquipment.DisplayName}'.\n\nCet equipement n'est pas configure dans la liste des equipements connus.",
                    "Configuration manquante",
                    this);
                return;
            }
            
            if (string.IsNullOrEmpty(_selectedEquipment.AssemblyFileName))
            {
                AddLog($"[-] Fichier IAM non defini pour '{_selectedEquipment.DisplayName}'", "ERROR");
                XnrgyMessageBox.ShowError(
                    $"Le fichier assemblage (.iam) n'est pas defini pour '{_selectedEquipment.DisplayName}'.",
                    "Configuration manquante",
                    this);
                return;
            }

            string? moduleIpjPath = null;
            string? moduleTopAssemblyPath = null;
            
            try
            {
                BtnPlaceEquipment.IsEnabled = false;
                BtnCancel.IsEnabled = false;
                BtnPreview.IsEnabled = false;
                
                // [+] RESET complet de la progression (temps a 00:00)
                ResetProgress();
                _startTime = DateTime.Now;
                _pausedTime = TimeSpan.Zero;
                
                UpdateProgress(0, "Initialisation du placement d'equipement...");
                AddLog("[>] Debut du placement d'equipement...", "START");
                Logger.Info("═══════════════════════════════════════════════════════");
                Logger.Info("[>] PLACEMENT EQUIPEMENT - DEMARRAGE");
                Logger.Info("═══════════════════════════════════════════════════════");
                Logger.Info($"[i] Equipement: {_selectedEquipment.DisplayName}");
                Logger.Info($"[i] Instance: {_selectedInstanceSuffix}");
                Logger.Info($"[i] IPJ source: {_selectedEquipment.ProjectFileName}");
                Logger.Info($"[i] IAM source: {_selectedEquipment.AssemblyFileName}");

                // Recuperer les infos du module destination
                var project = TxtProject.Text?.Trim() ?? "";
                var reference = CmbReference?.SelectedItem?.ToString() ?? "";
                var module = CmbModule?.SelectedItem?.ToString() ?? "";
                var fullProjectNumber = $"{project}{reference}{module}";
                
                // Construire les chemins - INCLURE LE SUFFIXE D'INSTANCE
                var modulePath = Path.Combine(_projectsBasePath, project, $"REF{reference}", $"M{module}");
                var equipmentFolderName = $"{_selectedEquipment.Name}{_selectedInstanceSuffix}";
                var equipmentDestPath = Path.Combine(modulePath, "1-Equipment", equipmentFolderName);
                moduleIpjPath = Path.Combine(modulePath, $"{fullProjectNumber}.ipj");
                moduleTopAssemblyPath = Path.Combine(modulePath, $"{fullProjectNumber}.iam");
                
                AddLog($"[i] Module destination: {project}-REF{reference}-M{module}", "INFO");
                AddLog($"[i] Dossier equipement: {equipmentFolderName}", "INFO");
                AddLog($"[i] Destination equipement: {equipmentDestPath}", "INFO");
                Logger.Info($"[i] Module: {project}-REF{reference}-M{module}");
                Logger.Info($"[i] Dossier equipement: {equipmentFolderName}");
                Logger.Info($"[i] Destination: {equipmentDestPath}");
                Logger.Info($"[i] IPJ module: {moduleIpjPath}");

                // Verifier que l'IPJ du module existe
                if (!File.Exists(moduleIpjPath))
                {
                    throw new Exception($"IPJ du module introuvable: {moduleIpjPath}");
                }

                // ══════════════════════════════════════════════════════════
                // ETAPE 1: Switch vers IPJ de l'equipement source
                // IMPORTANT: Le Copy Design doit etre fait avec l'IPJ de l'equipement
                // pour que les references soient correctement resolues
                // ══════════════════════════════════════════════════════════
                UpdateProgress(5, "Switch vers IPJ de l'equipement source...");
                AddLog("[1/8] Switch vers IPJ de l'equipement source...", "INFO");
                
                // Construire le chemin complet de l'IPJ de l'equipement
                // IMPORTANT: Utiliser le ProjectFileName de AvailableEquipment, pas un IPJ random
                var equipmentIpjPath = Path.Combine(_selectedEquipment.LocalTempPath, _selectedEquipment.ProjectFileName);
                
                // DEBUG: Logger tous les fichiers IPJ dans le dossier
                Logger.Info($"[DEBUG] Recherche IPJ dans: {_selectedEquipment.LocalTempPath}");
                Logger.Info($"[DEBUG] IPJ attendu: {_selectedEquipment.ProjectFileName}");
                
                if (Directory.Exists(_selectedEquipment.LocalTempPath))
                {
                    var allIpjFiles = Directory.GetFiles(_selectedEquipment.LocalTempPath, "*.ipj", SearchOption.AllDirectories);
                    Logger.Info($"[DEBUG] Tous les fichiers IPJ trouves ({allIpjFiles.Length}):");
                    foreach (var ipj in allIpjFiles)
                    {
                        Logger.Info($"[DEBUG]   - {ipj}");
                    }
                }
                
                if (!File.Exists(equipmentIpjPath))
                {
                    AddLog($"[!] IPJ attendu non trouve: {_selectedEquipment.ProjectFileName}", "WARN");
                    Logger.Warning($"[!] IPJ attendu non trouve: {equipmentIpjPath}");
                    
                    // Chercher l'IPJ EXACT par nom dans les sous-dossiers (pas n'importe quel IPJ!)
                    var ipjFiles = Directory.GetFiles(_selectedEquipment.LocalTempPath, "*.ipj", SearchOption.AllDirectories);
                    
                    // Chercher UNIQUEMENT un IPJ qui correspond EXACTEMENT au nom attendu
                    var expectedIpjName = _selectedEquipment.ProjectFileName;
                    var matchingIpj = ipjFiles.FirstOrDefault(f => 
                        Path.GetFileName(f).Equals(expectedIpjName, StringComparison.OrdinalIgnoreCase));
                    
                    if (matchingIpj != null)
                    {
                        equipmentIpjPath = matchingIpj;
                        AddLog($"[+] IPJ equipement trouve (sous-dossier): {Path.GetFileName(equipmentIpjPath)}", "INFO");
                        Logger.Info($"[+] IPJ trouve dans sous-dossier: {equipmentIpjPath}");
                    }
                    else
                    {
                        // PAS DE FALLBACK! Si l'IPJ master n'existe pas, ERREUR!
                        // La liste master dans EquipmentPlacementService.cs doit etre mise a jour
                        AddLog($"[-] ERREUR: IPJ '{expectedIpjName}' introuvable!", "ERROR");
                        AddLog($"[-] Verifiez la liste master dans EquipmentPlacementService.cs", "ERROR");
                        Logger.Error($"[-] IPJ master '{expectedIpjName}' introuvable dans Vault!");
                        Logger.Error($"[-] IPJ disponibles: {string.Join(", ", ipjFiles.Select(Path.GetFileName))}");
                        throw new Exception($"IPJ master '{expectedIpjName}' introuvable! Mettez a jour EquipmentPlacementService.AvailableEquipment avec le bon nom IPJ.");
                    }
                }
                else
                {
                    AddLog($"[+] IPJ equipement: {Path.GetFileName(equipmentIpjPath)}", "SUCCESS");
                }
                
                Logger.Info($"[i] IPJ equipement utilise: {equipmentIpjPath}");
                
                // ══════════════════════════════════════════════════════════
                // ETAPE 2: COPY DESIGN - Copier l'equipement vers 1-Equipment
                // UTILISATION DE EquipmentCopyDesignService (PAS InventorCopyDesignService!)
                // EquipmentCopyDesignService utilise l'IPJ passe en parametre directement
                // sans chercher automatiquement un IPJ dans le dossier source
                // ══════════════════════════════════════════════════════════
                UpdateProgress(10, "Preparation du Copy Design...");
                AddLog("[2/7] Copy Design de l'equipement vers destination...", "INFO");
                
                // Creer le dossier destination si necessaire
                if (!Directory.Exists(equipmentDestPath))
                {
                    Directory.CreateDirectory(equipmentDestPath);
                    AddLog($"[+] Dossier cree: {equipmentDestPath}", "SUCCESS");
                    Logger.Info($"[+] Dossier destination cree: {equipmentDestPath}");
                }
                
                // Executer le Copy Design avec EquipmentCopyDesignService
                // Ce service utilise l'IPJ passe en parametre (equipmentIpjPath) DIRECTEMENT
                // Pas de recherche automatique = PAS de selection du mauvais IPJ "Benoit"
                using (var equipmentCopyService = new EquipmentCopyDesignService(
                    (message, level) => Dispatcher.Invoke(() => AddLog(message, level)),
                    (percent, statusText) =>
                    {
                        // Mapper la progression 0-100 du Copy Design vers 5-50
                        int mappedPercent = 5 + (int)(percent * 0.45);
                        string currentFile = "";
                        if (statusText.Contains(":"))
                        {
                            var parts = statusText.Split(':');
                            if (parts.Length > 1) currentFile = parts[parts.Length - 1].Trim();
                        }
                        UpdateProgress(mappedPercent, statusText, false, currentFile);
                    }))
                {
                    if (!equipmentCopyService.Initialize())
                    {
                        throw new Exception("Impossible de se connecter a Inventor. Verifiez qu'Inventor 2026 est demarre.");
                    }
                    
                    // Executer le Copy Design avec l'IPJ specifique de l'equipement
                    // equipmentIpjPath = IPJ correct de AvailableEquipment (ex: "Angular Filter.ipj")
                    // _selectedEquipment.LocalTempPath = dossier source de l'equipement
                    // equipmentDestPath = destination dans le module
                    // _selectedEquipment.AssemblyFileName = assemblage principal (ex: "Angular Filter.iam")
                    var result = await equipmentCopyService.ExecuteEquipmentCopyDesignAsync(
                        equipmentIpjPath,
                        _selectedEquipment.LocalTempPath,
                        equipmentDestPath,
                        _selectedEquipment.AssemblyFileName,
                        null);  // Copier tous les fichiers (pas de filtre)
                    
                    if (!result.Success)
                    {
                        throw new Exception($"Copy Design echoue: {result.ErrorMessage}");
                    }
                    
                    AddLog($"[+] Copy Design termine: {result.FilesCopied} fichiers copies", "SUCCESS");
                    Logger.Info($"[+] Copy Design termine: {result.FilesCopied} fichiers");
                }

                // ══════════════════════════════════════════════════════════
                // ETAPE 3: Determiner le chemin de l'assemblage copie
                // ══════════════════════════════════════════════════════════
                UpdateProgress(52, "Localisation de l'assemblage copie...");
                AddLog("[3/8] Recherche de l'assemblage copie...", "INFO");
                
                // L'assemblage copie devrait etre renomme avec le numero de projet
                // ou garder son nom original selon la config
                var copiedEquipmentAssembly = Path.Combine(equipmentDestPath, _selectedEquipment.AssemblyFileName);
                
                // Chercher aussi avec le nom renomme (fullProjectNumber.iam)
                if (!File.Exists(copiedEquipmentAssembly))
                {
                    // Peut-etre renomme - chercher le premier .iam dans le dossier
                    var iamFiles = Directory.GetFiles(equipmentDestPath, "*.iam", SearchOption.TopDirectoryOnly);
                    if (iamFiles.Length > 0)
                    {
                        copiedEquipmentAssembly = iamFiles[0];
                        AddLog($"[i] Assemblage trouve: {Path.GetFileName(copiedEquipmentAssembly)}", "INFO");
                    }
                    else
                    {
                        throw new Exception($"Aucun assemblage .iam trouve dans: {equipmentDestPath}");
                    }
                }
                Logger.Info($"[i] Assemblage copie: {copiedEquipmentAssembly}");

                // ══════════════════════════════════════════════════════════
                // ETAPE 4: Switch vers IPJ du module destination
                // ══════════════════════════════════════════════════════════
                UpdateProgress(55, $"Switch vers IPJ module: {Path.GetFileName(moduleIpjPath)}...");
                AddLog($"[4/8] Switch vers IPJ module: {Path.GetFileName(moduleIpjPath)}...", "INFO");
                
                using (var inventorService = new EquipmentCopyDesignService(
                    (m, l) => Dispatcher.Invoke(() => AddLog(m, l)), (p, s) => { }))
                {
                    if (!inventorService.Initialize())
                    {
                        throw new Exception("Impossible de se reconnecter a Inventor");
                    }
                    
                    inventorService.SwitchProject(moduleIpjPath);
                    AddLog($"[+] IPJ module actif: {Path.GetFileName(moduleIpjPath)}", "SUCCESS");
                    Logger.Info($"[+] Switch vers IPJ module: {moduleIpjPath}");
                    
                    await Task.Delay(500);

                    // ══════════════════════════════════════════════════════════
                    // ETAPE 5: Ouvrir l'assemblage de l'equipement copie
                    // ══════════════════════════════════════════════════════════
                    UpdateProgress(60, $"Ouverture de l'equipement copie...");
                    AddLog($"[5/8] Ouverture de l'assemblage equipement copie...", "INFO");
                    
                    inventorService.OpenDocument(copiedEquipmentAssembly);
                    AddLog($"[+] Equipement ouvert: {Path.GetFileName(copiedEquipmentAssembly)}", "SUCCESS");
                    Logger.Info($"[+] Assemblage equipement ouvert: {copiedEquipmentAssembly}");
                    
                    await Task.Delay(1000);

                    // ══════════════════════════════════════════════════════════
                    // ETAPE 6: Appliquer les iProperties a l'equipement
                    // ══════════════════════════════════════════════════════════
                    UpdateProgress(65, "Application des iProperties...");
                    AddLog("[6/8] Application des iProperties a l'equipement...", "INFO");
                    
                    // Appliquer les iProperties du formulaire a l'assemblage equipement ouvert
                    int propsApplied = ApplyIPropertiesToEquipment();
                    if (propsApplied > 0)
                    {
                        AddLog($"[+] {propsApplied} iProperties appliquees a l'equipement", "SUCCESS");
                        Logger.Info($"[+] {propsApplied} iProperties appliquees a l'equipement");
                    }
                    else
                    {
                        AddLog("[!] Aucune iProperty appliquee", "WARN");
                        Logger.Warning("[!] Aucune iProperty appliquee a l'equipement");
                    }
                    
                    // Preparer la vue pour le designer (cacher plans, vue ISO, zoom all)
                    AddLog("[>] Preparation vue equipement (plans, ISO, zoom)...", "INFO");
                    inventorService.PrepareViewForDesigner();
                    AddLog("[+] Vue equipement preparee", "SUCCESS");
                    
                    // Sauvegarder l'equipement
                    inventorService.SaveAll();
                    AddLog("[+] Equipement sauvegarde", "SUCCESS");

                    // ══════════════════════════════════════════════════════════
                    // ETAPE 7: Fermer l'equipement et ouvrir le module
                    // ══════════════════════════════════════════════════════════
                    UpdateProgress(75, "Fermeture equipement, ouverture module...");
                    AddLog("[7/8] Fermeture equipement et ouverture module...", "INFO");
                    
                    // Fermer tous les documents
                    inventorService.CloseAllDocumentsPublic();
                    AddLog("[+] Equipement ferme", "SUCCESS");
                    Logger.Info("[+] Assemblage equipement ferme");
                    
                    await Task.Delay(500);
                    
                    // Verifier que l'assemblage du module existe
                    if (!File.Exists(moduleTopAssemblyPath))
                    {
                        throw new Exception($"Assemblage module introuvable: {moduleTopAssemblyPath}");
                    }
                    
                    // Ouvrir l'assemblage principal du module
                    inventorService.OpenDocument(moduleTopAssemblyPath);
                    AddLog($"[+] Module ouvert: {Path.GetFileName(moduleTopAssemblyPath)}", "SUCCESS");
                    Logger.Info($"[+] Assemblage module ouvert: {moduleTopAssemblyPath}");
                    
                    await Task.Delay(1000);

                    // ══════════════════════════════════════════════════════════
                    // ETAPE 8: Inserer l'equipement dans le module
                    // ══════════════════════════════════════════════════════════
                    UpdateProgress(85, $"Insertion de l'equipement dans le module...");
                    AddLog($"[8/8] Insertion de {Path.GetFileName(copiedEquipmentAssembly)} dans le module...", "INFO");
                    
                    bool placed = inventorService.PlaceComponent(copiedEquipmentAssembly);
                    if (placed)
                    {
                        AddLog($"[+] Equipement insere: {Path.GetFileName(copiedEquipmentAssembly)}", "SUCCESS");
                        Logger.Info($"[+] Equipement insere dans le module");
                        
                        // Preparer la vue: cacher plans, vue ISO, zoom all
                        inventorService.PrepareViewForDesigner();
                        AddLog("[+] Vue preparee (plans caches, ISO, Zoom All)", "SUCCESS");
                        
                        // Sauvegarder le module
                        inventorService.SaveAll();
                        AddLog("[+] Module sauvegarde", "SUCCESS");
                    }
                    else
                    {
                        AddLog($"[!] Placement automatique echoue - placement manuel requis", "WARN");
                        Logger.Warning("[!] PlaceComponent a retourne false, placement manuel requis");
                    }
                }
                
                // ══════════════════════════════════════════════════════════
                // TERMINE
                // ══════════════════════════════════════════════════════════
                UpdateProgress(100, $"[+] Equipement {_selectedEquipment.DisplayName} place avec succes!");
                AddLog($"[+] PLACEMENT TERMINE: {_selectedEquipment.DisplayName}", "SUCCESS");
                Logger.Info("═══════════════════════════════════════════════════════");
                Logger.Info($"[+] PLACEMENT EQUIPEMENT TERMINE: {_selectedEquipment.DisplayName}");
                Logger.Info("═══════════════════════════════════════════════════════");
                
                XnrgyMessageBox.ShowSuccess(
                    $"L'equipement '{_selectedEquipment.DisplayName}' a ete place avec succes dans le module {project}{reference}{module}.",
                    "Placement termine",
                    this);
            }
            catch (Exception ex)
            {
                AddLog($"[-] ERREUR: {ex.Message}", "ERROR");
                Logger.Error($"[-] Erreur placement equipement: {ex.Message}");
                Logger.Debug($"    StackTrace: {ex.StackTrace}");
                UpdateProgress(0, $"[-] Erreur: {ex.Message}", isError: true);
                
                XnrgyMessageBox.ShowError(
                    $"Erreur lors du placement de l'equipement:\n\n{ex.Message}",
                    "Erreur placement",
                    this);
                
                // Tenter de restaurer l'IPJ du module en cas d'erreur
                if (!string.IsNullOrEmpty(moduleIpjPath) && File.Exists(moduleIpjPath))
                {
                    try
                    {
                        using (var restoreService = new EquipmentCopyDesignService((m, l) => { }, (p, s) => { }))
                        {
                            if (restoreService.Initialize())
                            {
                                restoreService.SwitchProject(moduleIpjPath);
                                AddLog($"[+] IPJ module restaure apres erreur", "INFO");
                            }
                        }
                    }
                    catch { }
                }
            }
            finally
            {
                BtnPlaceEquipment.IsEnabled = true;
                BtnCancel.IsEnabled = true;
                BtnPreview.IsEnabled = true;
                
                // ══════════════════════════════════════════════════════════════════
                // NETTOYAGE COMPLET A LA FIN: Vider TOUT le dossier Equipment
                // Que le placement ait reussi ou echoue, on nettoie TOUT
                // pour garantir que le dossier Equipment reste vide
                // ══════════════════════════════════════════════════════════════════
                try
                {
                    AddLog("[>] NETTOYAGE COMPLET du dossier Equipment...", "INFO");
                    Logger.Info("[>] NETTOYAGE COMPLET post-placement: C:\\Vault\\Engineering\\Library\\Equipment");
                    bool cleanSuccess = CleanEntireEquipmentFolder();
                    if (cleanSuccess)
                    {
                        AddLog("[+] Dossier Equipment nettoye completement", "SUCCESS");
                    }
                    else
                    {
                        AddLog("[!] Nettoyage partiel du dossier Equipment", "WARN");
                    }
                }
                catch (Exception cleanEx)
                {
                    Logger.Warning($"[!] Erreur nettoyage post-placement: {cleanEx.Message}");
                }
                
                AddLog("Operation terminee", "STOP");
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

        /// <summary>
        /// Ouvre la fenêtre de réglages - Accessible uniquement aux administrateurs Vault
        /// </summary>
        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AddLog("Ouverture des reglages administrateur...", "START");
                
                // Passer le VaultService pour la synchronisation Vault
                var settingsWindow = new CreateModuleSettingsWindow(_vaultService)
                {
                    Owner = this
                };
                
                var result = settingsWindow.ShowDialog();
                
                if (result == true)
                {
                    // Recharger les parametres apres sauvegarde
                    AddLog("[+] Parametres mis a jour et synchronises vers Vault", "SUCCESS");
                    ReloadSettingsFromService();
                }
                else
                {
                    AddLog("Reglages annules", "INFO");
                }
            }
            catch (Exception ex)
            {
                AddLog($"Erreur ouverture réglages: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// Recharge les paramètres depuis le service après modification
        /// </summary>
        private void ReloadSettingsFromService()
        {
            try
            {
                SettingsService.Reload();
                var settings = SettingsService.Current?.CreateModule;
                
                if (settings != null)
                {
                    // Recharger les initiales dessinateurs
                    CmbInitialeDessinateur.Items.Clear();
                    CmbInitialeCoDessinateur.Items.Clear();
                    
                    foreach (var initial in settings.DesignerInitials)
                    {
                        CmbInitialeDessinateur.Items.Add(initial);
                        CmbInitialeCoDessinateur.Items.Add(initial);
                    }
                    
                    // Ajouter "Autre..." si pas présent
                    if (!settings.DesignerInitials.Contains("Autre..."))
                    {
                        CmbInitialeDessinateur.Items.Add("Autre...");
                        CmbInitialeCoDessinateur.Items.Add("Autre...");
                    }
                    
                    // Resélectionner N/A par défaut si disponible
                    if (CmbInitialeDessinateur.Items.Contains("N/A"))
                    {
                        CmbInitialeDessinateur.SelectedItem = "N/A";
                        CmbInitialeCoDessinateur.SelectedItem = "N/A";
                    }
                    
                    AddLog($"[+] {settings.DesignerInitials.Count} initiales rechargées", "SUCCESS");
                }
            }
            catch (Exception ex)
            {
                AddLog($"[!] Erreur rechargement paramètres: {ex.Message}", "WARN");
            }
        }

        #endregion

        #region Helpers

        private void UpdateStatistics()
        {
            // Les statistiques dans le panneau gauche ont été supprimées
            // On garde juste les statistiques dans l'en-tête

            var iamCount = _files.Count(f => f.FileType == "IAM");
            var iptCount = _files.Count(f => f.FileType == "IPT");
            var idwCount = _files.Count(f => f.FileType == "IDW");
            var inventorCount = iamCount + iptCount + idwCount;
            var otherCount = _files.Count(f => f.FileType != "IAM" && f.FileType != "IPT" && f.FileType != "IDW");
            var selectedCount = _files.Count(f => f.IsSelected);

            // Statistiques dans l'en-tête (nouvelles)
            if (TxtStatsTotal != null) TxtStatsTotal.Text = _files.Count.ToString();
            if (TxtStatsInventor != null) TxtStatsInventor.Text = inventorCount.ToString();
            if (TxtStatsOther != null) TxtStatsOther.Text = otherCount.ToString();
            if (TxtStatsSelected != null) TxtStatsSelected.Text = selectedCount.ToString();
        }

        private void ValidateForm()
        {
            if (BtnPlaceEquipment == null) return;

            var validation = ValidateAllInputs();
            BtnPlaceEquipment.IsEnabled = validation.IsValid;
            
            // Afficher message READY ou ACTION selon la validation
            if (validation.IsValid)
            {
                AddLog("READY", "[+] Pret pour le placement de l'equipement");
            }
            else if (!string.IsNullOrEmpty(validation.ErrorMessage))
            {
                AddLog("ACTION", $"[!] {validation.ErrorMessage}");
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
        /// [+] Pour PlaceEquipment: destination dans modulePath\1-Equipment\EquipmentName
        /// </summary>
        private void UpdateFileDestination(FileRenameItem item)
        {
            if (item == null) return;
            
            // [+] Pour PlaceEquipment: ajouter 1-Equipment\EquipmentName au chemin de base du module
            var modulePath = _request.DestinationPath; // C:\Vault\...\12345\REF01\M02
            var equipmentName = _selectedEquipment?.Name ?? "";
            var destBase = string.IsNullOrEmpty(equipmentName) 
                ? modulePath 
                : Path.Combine(modulePath, "1-Equipment", equipmentName);
            
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
            
            AddLog($"[i] Filtre extension: {extensions.Count} types detectes", "INFO");
        }
        
        /// <summary>
        /// Met à jour le filtre des statuts avec les statuts réellement trouvés
        /// </summary>
        private void UpdateStatusFilter()
        {
            if (CmbFilterSelection == null) return;
            
            // Collecter tous les statuts uniques
            var statuses = _files
                .Select(f => f.Status)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct()
                .OrderBy(s => s)
                .ToList();
            
            // Sauvegarder la sélection actuelle
            string? currentSelection = null;
            if (CmbFilterSelection.SelectedItem is ComboBoxItem selItem)
                currentSelection = selItem.Content?.ToString();
            
            // Mettre à jour les items
            CmbFilterSelection.Items.Clear();
            CmbFilterSelection.Items.Add(new ComboBoxItem { Content = "Tous", IsSelected = true });
            foreach (var status in statuses)
            {
                CmbFilterSelection.Items.Add(new ComboBoxItem { Content = status });
            }
            
            // Restaurer la sélection ou mettre "Tous"
            CmbFilterSelection.SelectedIndex = 0;
            if (!string.IsNullOrEmpty(currentSelection))
            {
                for (int i = 0; i < CmbFilterSelection.Items.Count; i++)
                {
                    if (CmbFilterSelection.Items[i] is ComboBoxItem item && item.Content?.ToString() == currentSelection)
                    {
                        CmbFilterSelection.SelectedIndex = i;
                        break;
                    }
                }
            }
            
            AddLog($"[i] Filtre statut: {statuses.Count} statuts detectes", "INFO");
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
            UpdateRenamePreviews();
        }

        #endregion
        
        #region Equipment Folder Cleanup Utilities
        
        /// <summary>
        /// Chemin de base du dossier Equipment
        /// </summary>
        private const string EQUIPMENT_BASE_PATH = @"C:\Vault\Engineering\Library\Equipment";
        
        /// <summary>
        /// NETTOIE COMPLETEMENT le dossier C:\Vault\Engineering\Library\Equipment
        /// Supprime TOUS les sous-dossiers et fichiers AVANT chaque telechargement
        /// pour garantir que seuls les fichiers a jour de Vault sont presents
        /// </summary>
        /// <returns>True si le nettoyage complet a reussi</returns>
        private bool CleanEntireEquipmentFolder()
        {
            try
            {
                Logger.Info($"[>] NETTOYAGE COMPLET: {EQUIPMENT_BASE_PATH}");
                
                if (!Directory.Exists(EQUIPMENT_BASE_PATH))
                {
                    Logger.Info("[i] Dossier Equipment n'existe pas - rien a nettoyer");
                    return true;
                }
                
                // Utiliser DirectoryInfo pour avoir acces aux attributs
                var baseDir = new DirectoryInfo(EQUIPMENT_BASE_PATH);
                
                // Etape 1: Enlever attributs sur le dossier principal
                baseDir.Attributes = FileAttributes.Normal;
                
                // Etape 2: Parcourir TOUS les fichiers et enlever ReadOnly
                foreach (var file in baseDir.GetFiles("*.*", SearchOption.AllDirectories))
                {
                    try
                    {
                        file.Attributes = FileAttributes.Normal;
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[!] Impossible de modifier attributs: {file.FullName} - {ex.Message}");
                    }
                }
                
                // Etape 3: Parcourir TOUS les dossiers et enlever ReadOnly
                foreach (var dir in baseDir.GetDirectories("*", SearchOption.AllDirectories))
                {
                    try
                    {
                        dir.Attributes = FileAttributes.Normal;
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[!] Impossible de modifier attributs dossier: {dir.FullName} - {ex.Message}");
                    }
                }
                
                // Etape 4: Supprimer TOUS les sous-dossiers (pas le dossier Equipment lui-meme)
                int deletedDirs = 0;
                int deletedFiles = 0;
                
                foreach (var subDir in baseDir.GetDirectories())
                {
                    try
                    {
                        // Compter les fichiers avant suppression
                        deletedFiles += Directory.GetFiles(subDir.FullName, "*.*", SearchOption.AllDirectories).Length;
                        
                        // Supprimer le sous-dossier completement
                        subDir.Delete(true);
                        deletedDirs++;
                        Logger.Debug($"[+] Supprime: {subDir.Name}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"[!] Echec suppression {subDir.Name}: {ex.Message}");
                        
                        // Tentative alternative: supprimer fichier par fichier
                        try
                        {
                            foreach (var file in subDir.GetFiles("*.*", SearchOption.AllDirectories))
                            {
                                try { file.Delete(); deletedFiles++; } catch { }
                            }
                            foreach (var dir in subDir.GetDirectories("*", SearchOption.AllDirectories).OrderByDescending(d => d.FullName.Length))
                            {
                                try { dir.Delete(false); } catch { }
                            }
                            try { subDir.Delete(false); deletedDirs++; } catch { }
                        }
                        catch { }
                    }
                }
                
                // Supprimer aussi les fichiers a la racine de Equipment (s'il y en a)
                foreach (var file in baseDir.GetFiles())
                {
                    try
                    {
                        file.Delete();
                        deletedFiles++;
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[!] Echec suppression fichier: {file.Name} - {ex.Message}");
                    }
                }
                
                Logger.Info($"[+] NETTOYAGE TERMINE: {deletedDirs} dossiers, {deletedFiles} fichiers supprimes");
                
                // Verifier si le nettoyage est complet
                var remaining = baseDir.GetDirectories().Length + baseDir.GetFiles().Length;
                if (remaining > 0)
                {
                    Logger.Warning($"[!] {remaining} elements restants dans {EQUIPMENT_BASE_PATH}");
                    return false;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"[-] Erreur nettoyage complet Equipment: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Enleve recursivement les attributs ReadOnly sur un dossier et tous ses fichiers
        /// FORCE: Enleve aussi les attributs Hidden et System
        /// </summary>
        /// <param name="path">Chemin du dossier a traiter</param>
        private void RemoveReadOnlyAttributes(string path)
        {
            if (!Directory.Exists(path)) return;
            
            try
            {
                Logger.Info($"[>] Suppression attributs ReadOnly sur: {path}");
                AddLog($"[>] Suppression ReadOnly: {Path.GetFileName(path)}...", "INFO");
                
                // Enlever ReadOnly sur le dossier lui-meme
                var dirInfo = new DirectoryInfo(path);
                dirInfo.Attributes = FileAttributes.Normal;
                
                // Traiter tous les fichiers recursivement (y compris _V)
                foreach (var file in Directory.GetFiles(path, "*.*", SearchOption.AllDirectories))
                {
                    try
                    {
                        var fileInfo = new FileInfo(file);
                        // Forcer Normal pour enlever ReadOnly, Hidden, System
                        fileInfo.Attributes = FileAttributes.Normal;
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[!] Impossible d'enlever attributs sur {file}: {ex.Message}");
                    }
                }
                
                // Traiter tous les sous-dossiers (y compris _V)
                foreach (var dir in Directory.GetDirectories(path, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        var subDirInfo = new DirectoryInfo(dir);
                        subDirInfo.Attributes = FileAttributes.Normal;
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"[!] Impossible d'enlever attributs sur dossier {dir}: {ex.Message}");
                    }
                }
                
                Logger.Info($"[+] Attributs ReadOnly enleves sur: {path}");
                AddLog($"[+] ReadOnly enleve", "SUCCESS");
            }
            catch (Exception ex)
            {
                Logger.Warning($"[!] Erreur lors de la suppression ReadOnly sur {path}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// FORCE la suppression complete d'un dossier equipement
        /// - Enleve TOUS les attributs (ReadOnly, Hidden, System)
        /// - Supprime le dossier et TOUT son contenu (y compris _V)
        /// - Utilise plusieurs tentatives si necessaire
        /// </summary>
        /// <param name="equipmentPath">Chemin du dossier equipement a supprimer</param>
        /// <returns>True si le nettoyage a reussi</returns>
        private bool CleanEquipmentFolder(string equipmentPath)
        {
            if (string.IsNullOrEmpty(equipmentPath)) return true;
            if (!Directory.Exists(equipmentPath)) return true;
            
            try
            {
                Logger.Info($"[>] NETTOYAGE FORCE du dossier equipement: {equipmentPath}");
                AddLog($"[>] Nettoyage FORCE de {Path.GetFileName(equipmentPath)}...", "INFO");
                
                // Etape 1: Enlever TOUS les attributs sur TOUS les fichiers et dossiers
                RemoveReadOnlyAttributes(equipmentPath);
                
                // Etape 2: Attendre un peu pour que les handles soient liberes
                System.Threading.Thread.Sleep(100);
                
                // Etape 3: Supprimer avec plusieurs tentatives
                int maxRetries = 3;
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    try
                    {
                        // Forcer la suppression recursive
                        Directory.Delete(equipmentPath, true);
                        Logger.Info($"[+] Dossier equipement SUPPRIME: {equipmentPath}");
                        AddLog($"[+] Dossier {Path.GetFileName(equipmentPath)} SUPPRIME", "SUCCESS");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"[!] Tentative {attempt}/{maxRetries} echouee: {ex.Message}");
                        
                        if (attempt < maxRetries)
                        {
                            // Re-essayer apres avoir enleve les attributs a nouveau
                            System.Threading.Thread.Sleep(200);
                            RemoveReadOnlyAttributes(equipmentPath);
                            
                            // Tenter de supprimer fichier par fichier
                            try
                            {
                                foreach (var file in Directory.GetFiles(equipmentPath, "*.*", SearchOption.AllDirectories))
                                {
                                    try { File.Delete(file); } catch { }
                                }
                                foreach (var dir in Directory.GetDirectories(equipmentPath, "*", SearchOption.AllDirectories).OrderByDescending(d => d.Length))
                                {
                                    try { Directory.Delete(dir, false); } catch { }
                                }
                            }
                            catch { }
                        }
                    }
                }
                
                // Si on arrive ici, le nettoyage n'a pas completement reussi
                // Verifier si le dossier est vide ou presque vide
                if (Directory.Exists(equipmentPath))
                {
                    var remainingFiles = Directory.GetFiles(equipmentPath, "*.*", SearchOption.AllDirectories);
                    if (remainingFiles.Length > 0)
                    {
                        Logger.Warning($"[!] {remainingFiles.Length} fichiers restants dans {equipmentPath}");
                        AddLog($"[!] {remainingFiles.Length} fichiers non supprimes", "WARN");
                        return false;
                    }
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"[-] Erreur nettoyage dossier {equipmentPath}: {ex.Message}");
                AddLog($"[-] Erreur nettoyage: {ex.Message}", "ERROR");
                return false;
            }
        }
        
        /// <summary>
        /// Prepare le dossier Equipment de base avant toute operation
        /// - Enleve les attributs ReadOnly sur C:\Vault\Engineering\Library\Equipment
        /// - Enleve les attributs sur TOUS les sous-dossiers
        /// </summary>
        private void PrepareEquipmentBaseFolder()
        {
            try
            {
                if (Directory.Exists(EQUIPMENT_BASE_PATH))
                {
                    Logger.Info($"[>] Preparation du dossier Equipment de base: {EQUIPMENT_BASE_PATH}");
                    AddLog("[>] Preparation dossier Equipment...", "INFO");
                    
                    // Enlever attributs sur le dossier principal
                    var dirInfo = new DirectoryInfo(EQUIPMENT_BASE_PATH);
                    dirInfo.Attributes = FileAttributes.Normal;
                    
                    // Enlever attributs sur tous les sous-dossiers de premier niveau
                    foreach (var subDir in Directory.GetDirectories(EQUIPMENT_BASE_PATH))
                    {
                        try
                        {
                            var subDirInfo = new DirectoryInfo(subDir);
                            subDirInfo.Attributes = FileAttributes.Normal;
                        }
                        catch { }
                    }
                    
                    Logger.Info($"[+] Dossier Equipment prepare");
                    AddLog("[+] Dossier Equipment prepare", "SUCCESS");
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"[!] Erreur preparation dossier Equipment: {ex.Message}");
            }
        }
        
        #endregion
        
        #region Application iProperties à l'équipement
        
        /// <summary>
        /// Applique les iProperties du formulaire à l'assemblage equipement ouvert dans Inventor
        /// Mapping des propriétés:
        /// - Initiale_du_Dessinateur : CmbInitialeDessinateur
        /// - Initiale_du_Co_Dessinateur : CmbInitialeCoDessinateur
        /// - Concepteur Lead CAD : CmbInitialeLeadCAD
        /// - Job_Title : TxtJobTitle
        /// - Creation_Date : DpCreationDate
        /// - Weight : (vide, sera mis a jour par iLogic)
        /// </summary>
        /// <returns>Nombre de proprietes appliquees</returns>
        private int ApplyIPropertiesToEquipment()
        {
            Logger.Info("[>] Application des iProperties a l'equipement...");
            AddLog("[>] Application des iProperties...", "INFO");
            
            int propsApplied = 0;
            
            try
            {
                var inventorApp = _inventorService.GetInventorApplication();
                if (inventorApp == null)
                {
                    Logger.Error("[-] Impossible d'obtenir l'application Inventor");
                    AddLog("[-] Inventor non disponible", "ERROR");
                    return 0;
                }
                
                dynamic activeDoc = inventorApp.ActiveDocument;
                if (activeDoc == null)
                {
                    Logger.Error("[-] Aucun document actif dans Inventor");
                    AddLog("[-] Aucun document actif", "ERROR");
                    return 0;
                }
                
                string docName = activeDoc.DisplayName;
                Logger.Info($"[i] Document actif pour iProperties: {docName}");
                
                // Acceder aux PropertySets du document
                var propertySets = activeDoc.PropertySets;
                
                // Obtenir ou creer le PropertySet custom
                dynamic customProps = null;
                try
                {
                    customProps = propertySets["Inventor User Defined Properties"];
                }
                catch
                {
                    Logger.Warning("[!] PropertySet custom non trouve, creation...");
                }
                
                if (customProps == null)
                {
                    Logger.Error("[-] Impossible d'acceder aux proprietes custom");
                    AddLog("[-] Proprietes custom non accessibles", "ERROR");
                    return 0;
                }
                
                // ══════════════════════════════════════════════════════════════
                // 1. Initiale_du_Dessinateur
                // ══════════════════════════════════════════════════════════════
                string dessinateur = GetSelectedComboValue(CmbInitialeDessinateur);
                if (!string.IsNullOrEmpty(dessinateur) && dessinateur != "N/A")
                {
                    if (SetOrCreateProperty(customProps, "Initiale_du_Dessinateur", dessinateur))
                    {
                        Logger.Info($"[+] Initiale_du_Dessinateur = {dessinateur}");
                        AddLog($"[+] Dessinateur: {dessinateur}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                // ══════════════════════════════════════════════════════════════
                // 2. Initiale_du_Co_Dessinateur
                // ══════════════════════════════════════════════════════════════
                string coDessinateur = GetSelectedComboValue(CmbInitialeCoDessinateur);
                if (!string.IsNullOrEmpty(coDessinateur) && coDessinateur != "N/A")
                {
                    if (SetOrCreateProperty(customProps, "Initiale_du_Co_Dessinateur", coDessinateur))
                    {
                        Logger.Info($"[+] Initiale_du_Co_Dessinateur = {coDessinateur}");
                        AddLog($"[+] Co-Dessinateur: {coDessinateur}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                // ══════════════════════════════════════════════════════════════
                // 3. Concepteur Lead CAD (NOUVEAU)
                // ══════════════════════════════════════════════════════════════
                string leadCAD = GetSelectedComboValue(CmbInitialeLeadCAD);
                if (!string.IsNullOrEmpty(leadCAD) && leadCAD != "N/A")
                {
                    if (SetOrCreateProperty(customProps, "Concepteur Lead CAD", leadCAD))
                    {
                        Logger.Info($"[+] Concepteur Lead CAD = {leadCAD}");
                        AddLog($"[+] Lead CAD: {leadCAD}", "SUCCESS");
                        propsApplied++;
                    }
                }
                else
                {
                    // Creer la propriete meme vide
                    SetOrCreateProperty(customProps, "Concepteur Lead CAD", "");
                    Logger.Info("[i] Concepteur Lead CAD cree (vide)");
                }
                
                // ══════════════════════════════════════════════════════════════
                // 4. Job_Title
                // ══════════════════════════════════════════════════════════════
                string jobTitle = TxtJobTitle?.Text?.Trim() ?? "";
                if (!string.IsNullOrEmpty(jobTitle))
                {
                    if (SetOrCreateProperty(customProps, "Job_Title", jobTitle))
                    {
                        Logger.Info($"[+] Job_Title = {jobTitle}");
                        AddLog($"[+] Job Title: {jobTitle}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                // ══════════════════════════════════════════════════════════════
                // 5. Creation_Date
                // ══════════════════════════════════════════════════════════════
                if (DpCreationDate?.SelectedDate != null)
                {
                    DateTime creationDate = DpCreationDate.SelectedDate.Value;
                    if (SetOrCreateProperty(customProps, "Creation_Date", creationDate))
                    {
                        Logger.Info($"[+] Creation_Date = {creationDate:yyyy-MM-dd}");
                        AddLog($"[+] Date: {creationDate:yyyy-MM-dd}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                // ══════════════════════════════════════════════════════════════
                // 6. Weight (vide, sera mis a jour par iLogic)
                // ══════════════════════════════════════════════════════════════
                // Creer la propriete Weight meme vide, pour qu'iLogic puisse la mettre a jour
                SetOrCreateProperty(customProps, "Weight", "");
                Logger.Info("[i] Weight cree (vide - sera mis a jour par iLogic)");
                
                // ══════════════════════════════════════════════════════════════
                // 7. Project (numero de projet - ex: 10359)
                // ══════════════════════════════════════════════════════════════
                string project = TxtProject?.Text?.Trim() ?? "";
                if (!string.IsNullOrEmpty(project))
                {
                    if (SetOrCreateProperty(customProps, "Project", project))
                    {
                        Logger.Info($"[+] Project = {project}");
                        AddLog($"[+] Project: {project}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                // ══════════════════════════════════════════════════════════════
                // 8. Reference (ex: 09 de REF09)
                // ══════════════════════════════════════════════════════════════
                string reference = CmbReference?.SelectedItem?.ToString() ?? "";
                if (!string.IsNullOrEmpty(reference))
                {
                    if (SetOrCreateProperty(customProps, "Reference", reference))
                    {
                        Logger.Info($"[+] Reference = {reference}");
                        AddLog($"[+] Reference: {reference}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                // ══════════════════════════════════════════════════════════════
                // 9. Module (ex: 03 de M03)
                // ══════════════════════════════════════════════════════════════
                string module = CmbModule?.SelectedItem?.ToString() ?? "";
                if (!string.IsNullOrEmpty(module))
                {
                    if (SetOrCreateProperty(customProps, "Module", module))
                    {
                        Logger.Info($"[+] Module = {module}");
                        AddLog($"[+] Module: {module}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                // ══════════════════════════════════════════════════════════════
                // 10. Numero_de_Projet (format complet: 1035909M03)
                // ══════════════════════════════════════════════════════════════
                if (!string.IsNullOrEmpty(project) && !string.IsNullOrEmpty(reference) && !string.IsNullOrEmpty(module))
                {
                    string numeroProjet = $"{project}{reference}M{module}";
                    if (SetOrCreateProperty(customProps, "Numero_de_Projet", numeroProjet))
                    {
                        Logger.Info($"[+] Numero_de_Projet = {numeroProjet}");
                        AddLog($"[+] Numero de Projet: {numeroProjet}", "SUCCESS");
                        propsApplied++;
                    }
                }
                
                Logger.Info($"[+] {propsApplied} iProperties appliquees a l'equipement");
                AddLog($"[+] {propsApplied} iProperties appliquees", "SUCCESS");
                
                return propsApplied;
            }
            catch (Exception ex)
            {
                Logger.Error($"[-] Erreur application iProperties: {ex.Message}");
                AddLog($"[-] Erreur iProperties: {ex.Message}", "ERROR");
                return propsApplied;
            }
        }
        
        /// <summary>
        /// Obtient la valeur selectionnee d'un ComboBox
        /// </summary>
        private string GetSelectedComboValue(ComboBox? combo)
        {
            if (combo == null || combo.SelectedItem == null) return "";
            return combo.SelectedItem.ToString() ?? "";
        }
        
        /// <summary>
        /// Definit ou cree une propriete custom dans Inventor
        /// ATTENTION: L'API Inventor PropertySets.Add() a l'ordre (value, name) et non (name, value)!
        /// </summary>
        private bool SetOrCreateProperty(dynamic customProps, string propName, object value)
        {
            try
            {
                // Essayer de modifier la propriete existante
                try
                {
                    dynamic prop = customProps[propName];
                    prop.Value = value;
                    return true;
                }
                catch
                {
                    // La propriete n'existe pas, la creer
                    // IMPORTANT: Add(value, name) - ordre inverse!
                    try
                    {
                        customProps.Add(value, propName);
                        return true;
                    }
                    catch (Exception createEx)
                    {
                        Logger.Warning($"[!] Impossible de creer la propriete '{propName}': {createEx.Message}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"[!] Erreur propriete '{propName}': {ex.Message}");
                return false;
            }
        }
        
        #endregion
    }
}
