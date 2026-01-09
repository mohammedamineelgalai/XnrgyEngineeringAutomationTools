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

namespace XnrgyEngineeringAutomationTools.Modules.CreateModule.Views
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
        public CreateModuleWindow() : this(null, null)
        {
        }

        /// <summary>
        /// Constructeur avec service Vault pour vérification admin
        /// </summary>
        /// <param name="vaultService">Service Vault connecté (optionnel)</param>
        public CreateModuleWindow(VaultSdkService? vaultService) : this(vaultService, null)
        {
        }

        /// <summary>
        /// Constructeur avec services Vault et Inventor pour héritage du statut de connexion
        /// </summary>
        /// <param name="vaultService">Service Vault connecté (optionnel)</param>
        /// <param name="inventorService">Service Inventor du formulaire principal (optionnel)</param>
        public CreateModuleWindow(VaultSdkService? vaultService, InventorService? inventorService)
        {
            // IMPORTANT: Initialiser _request et _files AVANT InitializeComponent()
            // car les événements TextChanged du XAML sont déclenchés pendant l'initialisation
            _request = new CreateModuleRequest();
            _files = new ObservableCollection<FileRenameItem>();
            _vaultService = vaultService;
            // Utiliser le service Inventor du formulaire principal (héritage du statut)
            _inventorService = inventorService ?? new InventorService();
            
            // [+] Forcer la reconnexion COM à chaque ouverture de Créer Module
            // Évite les problèmes de connexion COM obsolète
            _inventorService.ForceReconnect();
            
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Erreur lors de l'initialisation de la fenêtre Créer Module:\n\n{ex.Message}\n\nDétails:\n{ex.StackTrace}",
                    "Erreur d'initialisation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                throw;
            }
            
            // S'abonner aux changements de theme
            MainWindow.ThemeChanged += OnThemeChanged;
            
            // Appliquer le theme actuel au demarrage
            ApplyTheme(MainWindow.CurrentThemeIsDark);
            
            // Attendre que la fenêtre soit chargée pour initialiser les contrôles
            this.Loaded += CreateModuleWindow_Loaded;
            this.Closed += (s, e) =>
            {
                MainWindow.ThemeChanged -= OnThemeChanged;
                _inventorStatusTimer?.Stop();
                // Nettoyer le dossier temporaire Vault si présent
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
                
                TxtStatus.Text = "Prêt - Remplissez les informations du projet";
                TxtProgressPercent.Text = "";
                TxtProgressTimeElapsed.Text = "00:00";
                TxtProgressTimeEstimated.Text = "00:00";
                TxtCurrentFile.Text = "";
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
                
                // Vérifier si l'utilisateur est administrateur Vault
                CheckVaultAdminPermissions();
                
                // Initialiser le statut Inventor
                UpdateInventorStatus();
                
                // Timer pour mettre à jour le statut Inventor périodiquement
                _inventorStatusTimer = new System.Windows.Threading.DispatcherTimer();
                _inventorStatusTimer.Interval = TimeSpan.FromSeconds(3);
                _inventorStatusTimer.Tick += (s, args) => UpdateInventorStatus();
                _inventorStatusTimer.Start();
                
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
                    TxtStatus.Text = "✅ Template Xnrgy_Module chargé automatiquement";
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
                    if (isConnected)
                    {
                        // Récupérer le nom du fichier actif
                        string? activeFileName = _inventorService.GetActiveDocumentName();
                        if (!string.IsNullOrEmpty(activeFileName))
                        {
                            // Tronquer si trop long
                            if (activeFileName.Length > 25)
                                activeFileName = activeFileName.Substring(0, 22) + "...";
                            RunInventorStatus.Text = $" Inventor : {activeFileName}";
                        }
                        else
                        {
                            RunInventorStatus.Text = " Inventor : Connecte";
                        }
                    }
                    else
                    {
                        RunInventorStatus.Text = " Inventor : Deconnecte";
                    }
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
            UpdateFullProjectNumber();
            UpdateDestinationPreview();
            ValidateForm();
        }

        private void CmbModule_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateFullProjectNumber();
            UpdateDestinationPreview();
            ValidateForm();
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
            // IMPORTANT: Mettre a jour le numero de projet AVANT les chemins de destination
            // Car UpdateFileDestinationPaths utilise _request.FullProjectNumber pour renommer les fichiers
            UpdateFullProjectNumber();
            UpdateDestinationPreview();
            UpdateRenamePreviews();
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
            // Verifier si un numero de projet valide est entre (pas juste des zeros)
            bool hasValidProjectNumber = !string.IsNullOrWhiteSpace(_request.Project) && _request.Project.Length >= 4;
            var fullProjectNumber = hasValidProjectNumber ? _request.FullProjectNumber : string.Empty;

            foreach (var file in _files)
            {
                // Calculer le chemin relatif une seule fois
                var relativeDir = Path.GetDirectoryName(file.RelativePath) ?? "";
                var isAtRoot = string.IsNullOrEmpty(relativeDir);
                
                // Renommer le Top Assembly (.iam) avec le numero de projet formate
                // Pour templates: fichier marque IsTopAssembly (Module_.iam ou 000000000.iam)
                // Pour projets existants: premier .iam a la racine
                if (!string.IsNullOrEmpty(fullProjectNumber))
                {
                    if (file.IsTopAssembly)
                    {
                        file.NewFileName = $"{fullProjectNumber}.iam";
                    }
                    else if (isFromExistingProject && isAtRoot && file.FileType == "IAM")
                    {
                        // Pour projets existants: renommer le .iam à la racine (un seul)
                        // Vérifier qu'on n'a pas déjà renommé un autre .iam
                        var alreadyRenamed = _files.Any(f => f != file && f.FileType == "IAM" && 
                            string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath)) &&
                            f.NewFileName == $"{fullProjectNumber}.iam");
                        if (!alreadyRenamed)
                        {
                            file.NewFileName = $"{fullProjectNumber}.iam";
                        }
                    }
                }
                else
                {
                    // Pas de projet valide - reinitialiser le NewFileName au nom original pour les fichiers IAM
                    if (file.IsTopAssembly || (isFromExistingProject && isAtRoot && file.FileType == "IAM"))
                    {
                        file.NewFileName = file.OriginalFileName;
                    }
                }
                
                // Renommer le fichier projet principal (.ipj)
                // Pour templates: pattern XXXXX-XX-XX_2026.ipj ou 000000000.ipj (nouveau template)
                // Pour projets existants: tout .ipj a la racine
                if (file.FileType == "IPJ" && isAtRoot)
                {
                    if (!string.IsNullOrEmpty(fullProjectNumber) && (isFromExistingProject || IsMainProjectFilePattern(file.OriginalFileName)))
                    {
                        file.NewFileName = $"{fullProjectNumber}.ipj";
                    }
                    else if (string.IsNullOrEmpty(fullProjectNumber) && IsMainProjectFilePattern(file.OriginalFileName))
                    {
                        // Reinitialiser au nom original si pas de projet valide
                        file.NewFileName = file.OriginalFileName;
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

            // Rafraîchir l'affichage - Réappliquer les filtres pour mettre à jour la vue
            ApplyFilters();
            UpdateStatistics();
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

        #region Event Handlers - Source Options

        private void SourceOption_Changed(object sender, RoutedEventArgs e)
        {
            if (RbFromTemplate == null || RbFromExisting == null || RbFromVault == null) return;
            if (PnlProjectSelector == null || PnlVaultProjectSelector == null) return;

            if (RbFromTemplate.IsChecked == true)
            {
                _request.Source = CreateModuleSource.FromTemplate;
                TxtSourcePath.Text = _defaultTemplatePath;
                TxtSourcePath.IsReadOnly = true;
                PnlProjectSelector.Visibility = Visibility.Collapsed;
                PnlVaultProjectSelector.Visibility = Visibility.Collapsed;
                TxtStatus.Text = "Mode: Création depuis Template";
                
                // Charger le template automatiquement
                if (Directory.Exists(_defaultTemplatePath))
                {
                    LoadFilesFromPath(_defaultTemplatePath);
                }
            }
            else if (RbFromExisting.IsChecked == true)
            {
                _request.Source = CreateModuleSource.FromExistingProject;
                TxtSourcePath.Text = "";
                TxtSourcePath.IsReadOnly = true;
                PnlProjectSelector.Visibility = Visibility.Visible;
                PnlVaultProjectSelector.Visibility = Visibility.Collapsed;
                TxtStatus.Text = "Mode: Création depuis Projet Existant - Sélectionnez un projet";
                
                // Charger la liste des projets
                LoadProjectsList();
            }
            else if (RbFromVault.IsChecked == true)
            {
                _request.Source = CreateModuleSource.FromVault;
                TxtSourcePath.Text = "";
                TxtSourcePath.IsReadOnly = true;
                PnlProjectSelector.Visibility = Visibility.Collapsed;
                PnlVaultProjectSelector.Visibility = Visibility.Visible;
                TxtStatus.Text = "Mode: Création depuis Vault - Sélectionnez un projet";
                
                // Charger la liste des projets Vault
                LoadVaultProjectsList();
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
                TxtStatus.Text = "⚠️ Dossier Projects non trouve";
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
                
                TxtStatus.Text = $"✅ {CmbProjects.Items.Count} modules trouves";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"⚠️ Erreur chargement projets: {ex.Message}";
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

        /// <summary>
        /// Classe pour représenter un projet Vault
        /// </summary>
        private class VaultProject
        {
            public string Project { get; set; } = string.Empty;
            public string Reference { get; set; } = string.Empty;
            public string Module { get; set; } = string.Empty;
            public string VaultPath { get; set; } = string.Empty;
            public string DisplayName => $"{Project} / {Reference} / {Module}";
            
            // Override ToString pour l'affichage dans ComboBox
            public override string ToString() => DisplayName;
        }

        private string? _tempVaultDownloadPath = null;

        private void LoadVaultProjectsList()
        {
            if (CmbVaultProjects == null) return;
            
            CmbVaultProjects.Items.Clear();
            
            if (_vaultService == null || !_vaultService.IsConnected)
            {
                TxtStatus.Text = "⚠️ Non connecté à Vault";
                return;
            }

            try
            {
                TxtStatus.Text = "Chargement des projets depuis Vault...";
                
                var vaultProjects = new List<VaultProject>();
                var projectsBasePath = "$/Engineering/Projects";
                
                var connection = _vaultService.Connection;
                if (connection == null)
                {
                    TxtStatus.Text = "⚠️ Connexion Vault non disponible";
                    return;
                }

                // Obtenir le dossier racine Projects
                var projectsFolder = connection.WebServiceManager.DocumentService.GetFolderByPath(projectsBasePath);
                if (projectsFolder == null)
                {
                    TxtStatus.Text = "⚠️ Dossier Projects non trouvé dans Vault";
                    return;
                }

                // Obtenir tous les sous-dossiers (Projets)
                var projectFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(projectsFolder.Id, false);
                if (projectFolders == null)
                {
                    TxtStatus.Text = "✅ Aucun projet trouvé";
                    return;
                }

                foreach (var projectFolder in projectFolders)
                {
                    var projectName = projectFolder.Name;
                    
                    // Obtenir les sous-dossiers REF
                    var refFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(projectFolder.Id, false);
                    if (refFolders == null) continue;
                    
                    foreach (var refFolder in refFolders)
                    {
                        var refName = refFolder.Name;
                        if (!refName.StartsWith("REF", StringComparison.OrdinalIgnoreCase)) continue;
                        
                        // Obtenir les sous-dossiers M (Modules)
                        var moduleFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(refFolder.Id, false);
                        if (moduleFolders == null) continue;
                        
                        foreach (var moduleFolder in moduleFolders)
                        {
                            var moduleName = moduleFolder.Name;
                            if (!moduleName.StartsWith("M", StringComparison.OrdinalIgnoreCase)) continue;
                            
                            var vaultPath = $"{projectsBasePath}/{projectName}/{refName}/{moduleName}";
                            vaultProjects.Add(new VaultProject
                            {
                                Project = projectName,
                                Reference = refName,
                                Module = moduleName,
                                VaultPath = vaultPath
                            });
                        }
                    }
                }
                
                // Trier par projet, référence, module
                vaultProjects = vaultProjects.OrderBy(p => p.Project)
                                             .ThenBy(p => p.Reference)
                                             .ThenBy(p => p.Module)
                                             .ToList();
                
                foreach (var project in vaultProjects)
                {
                    CmbVaultProjects.Items.Add(project);
                }
                
                TxtStatus.Text = $"✅ {CmbVaultProjects.Items.Count} projets Vault trouvés";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"⚠️ Erreur chargement projets Vault: {ex.Message}";
                AddLog($"Erreur chargement projets Vault: {ex.Message}", "ERROR");
            }
        }

        private void CmbVaultProjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbVaultProjects?.SelectedItem is VaultProject project)
            {
                TxtSourcePath.Text = $"Vault: {project.VaultPath}";
                _request.SourceExistingProjectPath = project.VaultPath;
            }
        }

        private void BtnRefreshVaultProjects_Click(object sender, RoutedEventArgs e)
        {
            LoadVaultProjectsList();
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

        private async void BtnLoadFiles_Click(object sender, RoutedEventArgs e)
        {
            if (_request.Source == CreateModuleSource.FromVault)
            {
                // Télécharger depuis Vault vers un dossier temporaire
                await LoadFilesFromVault();
            }
            else
            {
                LoadFilesFromPath(TxtSourcePath?.Text);
            }
        }

        private async Task LoadFilesFromVault()
        {
            if (CmbVaultProjects?.SelectedItem is not VaultProject project)
            {
                TxtStatus.Text = "[!] Veuillez selectionner un projet Vault";
                return;
            }

            if (_vaultService == null || !_vaultService.IsConnected)
            {
                TxtStatus.Text = "[!] Non connecte a Vault";
                return;
            }

            try
            {
                // [+] RESET du timer au debut de l'operation
                _startTime = DateTime.Now;
                _pausedTime = TimeSpan.Zero;
                
                UpdateProgress(0, "Preparation du telechargement...");
                AddLog("[>] Demarrage telechargement depuis Vault...", "START");
                BtnLoadFiles.IsEnabled = false;

                var connection = _vaultService.Connection;
                if (connection == null)
                {
                    TxtStatus.Text = "[!] Connexion Vault perdue";
                    BtnLoadFiles.IsEnabled = true;
                    return;
                }

                // Obtenir le working folder - Vault telecharge directement dedans
                var workingFolderObj = connection.WorkingFoldersManager.GetWorkingFolder("$");
                if (workingFolderObj == null || string.IsNullOrEmpty(workingFolderObj.FullPath))
                {
                    TxtStatus.Text = "[!] Working folder non configure";
                    AddLog("[-] Working folder non configure dans Vault", "ERROR");
                    BtnLoadFiles.IsEnabled = true;
                    return;
                }
                
                var workingFolder = workingFolderObj.FullPath;
                var relativePath = project.VaultPath.TrimStart('$', '/').Replace("/", "\\");
                var localProjectPath = Path.Combine(workingFolder, relativePath);
                
                AddLog($"[i] Working folder: {workingFolder}", "DEBUG");
                AddLog($"[i] Chemin local projet: {localProjectPath}", "DEBUG");
                
                // PAS de dossier temporaire - Vault telecharge directement dans le working folder
                _tempVaultDownloadPath = localProjectPath;

                // Obtenir le dossier Vault
                UpdateProgress(5, "Connexion au dossier Vault...");
                var vaultFolder = connection.WebServiceManager.DocumentService.GetFolderByPath(project.VaultPath);
                if (vaultFolder == null)
                {
                    TxtStatus.Text = "[!] Dossier Vault non trouve";
                    AddLog($"[-] Dossier Vault non trouve: {project.VaultPath}", "ERROR");
                    BtnLoadFiles.IsEnabled = true;
                    return;
                }
                AddLog($"[+] Dossier Vault trouve: {project.VaultPath}", "INFO");

                // Obtenir TOUS les fichiers RECURSIVEMENT (dossier + sous-dossiers)
                UpdateProgress(10, "Enumeration recursive des fichiers...");
                AddLog($"[>] Enumeration recursive depuis: {project.VaultPath}", "INFO");
                
                var allFiles = new List<ACW.File>();
                var allFolders = new List<ACW.Folder>();
                await Task.Run(() => GetAllFilesRecursive(connection, vaultFolder, allFiles, allFolders));
                
                AddLog($"[+] {allFolders.Count} dossiers trouves", "INFO");
                foreach (var folder in allFolders.Take(10)) // Log les 10 premiers
                {
                    AddLog($"    [i] {folder.FullName}", "DEBUG");
                }
                if (allFolders.Count > 10)
                {
                    AddLog($"    ... et {allFolders.Count - 10} autres dossiers", "DEBUG");
                }
                
                if (allFiles.Count == 0)
                {
                    TxtStatus.Text = "[!] Aucun fichier trouve dans le projet Vault";
                    AddLog("[-] Aucun fichier trouve (meme recursivement)", "ERROR");
                    BtnLoadFiles.IsEnabled = true;
                    return;
                }
                AddLog($"[+] {allFiles.Count} fichiers trouves au total (recursif)", "SUCCESS");
                
                // Convertir en array pour compatibilite
                var files = allFiles.ToArray();

                // Preparer le telechargement batch
                UpdateProgress(15, $"Preparation de {files.Length} fichiers...");
                var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection, false);
                
                int fileIndex = 0;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                
                foreach (var file in files)
                {
                    try
                    {
                        var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(connection, file);
                        downloadSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                        fileIndex++;
                    }
                    catch (Exception fileEx)
                    {
                        AddLog($"[!] Erreur preparation {file.Name}: {fileEx.Message}", "WARNING");
                    }
                }
                sw.Stop();
                AddLog($"[+] {fileIndex} fichiers prepares en {sw.ElapsedMilliseconds}ms", "INFO");

                // ══════════════════════════════════════════════════════════════════
                // TELECHARGEMENT PAR LOTS avec progression precise (20-70%)
                // Telecharge par lots de 10 fichiers pour permettre mise a jour UI
                // ══════════════════════════════════════════════════════════════════
                const int BATCH_SIZE = 10;
                int downloadedCount = 0;
                int totalToDownload = files.Length;
                var allDownloadResults = new List<VDF.Vault.Results.FileAcquisitionResult>();
                
                UpdateProgress(20, $"Telechargement de {totalToDownload} fichiers...");
                AddLog($"[>] Lancement telechargement par lots de {BATCH_SIZE}...", "INFO");
                
                sw.Restart();
                
                // Diviser en lots pour mise a jour progressive
                var fileBatches = files
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
                        AddLog($"[!] Erreur lot: {batchEx.Message}", "WARNING");
                        downloadedCount += batch.Count; // Continuer meme en cas d'erreur
                    }
                }
                
                sw.Stop();

                if (allDownloadResults.Count == 0)
                {
                    TxtStatus.Text = "[!] Aucun fichier telecharge";
                    AddLog("[-] AcquireFiles n'a retourne aucun resultat", "ERROR");
                    BtnLoadFiles.IsEnabled = true;
                    return;
                }

                var fileResultsList = allDownloadResults;
                int successCount = fileResultsList.Count(r => r.LocalPath?.FullPath != null && File.Exists(r.LocalPath.FullPath));
                AddLog($"[+] {successCount}/{fileResultsList.Count} fichiers telecharges en {sw.ElapsedMilliseconds}ms ({sw.ElapsedMilliseconds / Math.Max(1, successCount)}ms/fichier)", "SUCCESS");

                // [+] PAS DE COPIE NECESSAIRE - Vault telecharge directement dans le working folder
                // qui est deja localProjectPath (C:\Vault\Engineering\Projects\XXX)
                // La copie precedente etait inutile et causait des erreurs
                
                UpdateProgress(75, "Verification des fichiers telecharges...");
                
                // Enlever les attributs ReadOnly sur les fichiers telecharges
                RemoveReadOnlyAttributesRecursive(localProjectPath);
                
                // Verifier si des fichiers ont ete telecharges
                var downloadedFiles = Directory.GetFiles(localProjectPath, "*.*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\_V\\"))  // Exclure les fichiers de version _V
                    .ToArray();
                    
                if (downloadedFiles.Length == 0)
                {
                    TxtStatus.Text = "[!] Aucun fichier telecharge";
                    AddLog("[-] Aucun fichier trouve dans le dossier projet", "ERROR");
                    BtnLoadFiles.IsEnabled = true;
                    return;
                }
                
                AddLog($"[+] {downloadedFiles.Length} fichiers telecharges dans le dossier projet", "SUCCESS");
                UpdateProgress(90, "Chargement de la liste des fichiers...");

                // Charger les fichiers directement depuis le dossier projet (PAS de temp)
                LoadFilesFromPath(localProjectPath);
                
                UpdateProgress(100, $"[+] {_files.Count} fichiers charges depuis Vault");
                AddLog($"[+] Telechargement Vault termine: {_files.Count} fichiers prets", "SUCCESS");
                TxtStatus.Text = $"[+] {_files.Count} fichiers telecharges depuis Vault";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"[-] Erreur: {ex.Message}";
                AddLog($"[-] Erreur telechargement Vault: {ex.Message}", "ERROR");
                AddLog($"    Stack: {ex.StackTrace?.Split('\n').FirstOrDefault()}", "DEBUG");
            }
            finally
            {
                BtnLoadFiles.IsEnabled = true;
            }
        }
        
        /// <summary>
        /// Enleve recursivement les attributs ReadOnly sur un dossier et tous ses fichiers
        /// </summary>
        private void RemoveReadOnlyAttributesRecursive(string path)
        {
            if (!Directory.Exists(path)) return;
            
            try
            {
                foreach (var file in Directory.GetFiles(path, "*.*", SearchOption.AllDirectories))
                {
                    try
                    {
                        var fileInfo = new FileInfo(file);
                        if (fileInfo.IsReadOnly)
                        {
                            fileInfo.IsReadOnly = false;
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

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

                // Determiner si on est en mode "Projet Existant" ou "Depuis Vault"
                // Les deux modes utilisent la meme logique de renommage (fichiers avec noms reels, pas templates)
                bool isFromExistingProject = _request.Source == CreateModuleSource.FromExistingProject;
                bool isFromVault = _request.Source == CreateModuleSource.FromVault;
                bool useExistingProjectLogic = isFromExistingProject || isFromVault;

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

                    // Pour templates: renommer Module_.iam/000000000.iam et fichier IPJ principal
                    // Pour projets existants ET Vault: renommer le premier .iam a la racine et le premier .ipj a la racine
                    if (!useExistingProjectLogic)
                    {
                        // Renommage automatique du Top Assembly (.iam) - template Module_.iam ou 000000000.iam
                        if (isTopAssembly && !string.IsNullOrEmpty(TxtProject?.Text))
                        {
                            item.NewFileName = $"{_request.FullProjectNumber}.iam";
                        }
                        
                        // Renommage automatique UNIQUEMENT du fichier projet principal (.ipj)
                        // Pattern: XXXXX-XX-XX_2026.ipj ou 000000000.ipj (nouveau template)
                        if (isMainProjectFile && !string.IsNullOrEmpty(TxtProject?.Text))
                        {
                            item.NewFileName = $"{_request.FullProjectNumber}.ipj";
                        }
                    }
                    else
                    {
                        // PROJET EXISTANT ou VAULT: Renommer le fichier Top Assembly (.iam à la racine) avec le numéro de projet
                        // Note: Pour les projets existants/Vault, isTopAssembly est false car le fichier ne s'appelle pas "Module_.iam"
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
                
                // Identifier les fichiers à EXCLURE des options de renommage:
                // - TopAssy (sera renommé avec le numéro de projet)
                // - IPJ principal (sera renommé avec le numéro de projet)
                // - IDW (drawings - héritent du nom de leur assembly parent, ne pas modifier)
                bool isFromExistingProject = _request.Source == CreateModuleSource.FromExistingProject;
                
                // [CRITIQUE] Detecter le Top Assembly - fichier master du module
                // Pour templates: 000000000.iam est le seul fichier master
                // Pour projets existants: fichier avec IsTopAssembly = true
                var topAssemblyFile = _files.FirstOrDefault(f => f.IsTopAssembly);
                if (topAssemblyFile == null)
                {
                    // Fallback pour templates: chercher 000000000.iam specifiquement
                    topAssemblyFile = _files.FirstOrDefault(f => f.FileType == "IAM" && 
                        f.OriginalFileName.Equals("000000000.iam", StringComparison.OrdinalIgnoreCase));
                }
                
                // IPJ principal: 000000000.ipj pour templates, ou IPJ a la racine pour projets existants
                var mainIpjFile = _files.FirstOrDefault(f => f.FileType == "IPJ" && 
                    f.OriginalFileName.Equals("000000000.ipj", StringComparison.OrdinalIgnoreCase));
                if (mainIpjFile == null && isFromExistingProject)
                {
                    // Pour projets existants: premier IPJ a la racine
                    mainIpjFile = _files.FirstOrDefault(f => f.FileType == "IPJ" && 
                        string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath)));
                }

                foreach (var file in selectedFiles)
                {
                    // Skip fichiers non-Inventor si checkbox non cochée
                    if (!includeNonInventor && !file.IsInventorFile)
                    {
                        continue;
                    }
                    
                    // [!] EXCLURE du renommage avec options: TopAssy et IPJ principal SEULEMENT
                    // Ces fichiers masters sont renommés avec le numéro de projet (sans préfixe)
                    if (file == topAssemblyFile || file == mainIpjFile)
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

                // Toujours renommer le Top Assembly et IPJ principal avec le numéro de projet
                // [CRITIQUE] Seuls les fichiers master specifiques (000000000.iam/ipj ou IsTopAssembly)
                
                // Top Assembly: 000000000.iam pour templates, ou IsTopAssembly pour projets existants
                var topAssembly = _files.FirstOrDefault(f => f.IsTopAssembly);
                if (topAssembly == null)
                {
                    // Fallback pour templates: 000000000.iam specifiquement
                    topAssembly = _files.FirstOrDefault(f => f.FileType == "IAM" && 
                        f.OriginalFileName.Equals("000000000.iam", StringComparison.OrdinalIgnoreCase));
                }
                if (topAssembly != null && !string.IsNullOrEmpty(TxtProject?.Text))
                {
                    topAssembly.NewFileName = $"{_request.FullProjectNumber}.iam";
                }
                
                // IPJ principal: 000000000.ipj pour templates
                var mainIpj = _files.FirstOrDefault(f => f.FileType == "IPJ" && 
                    f.OriginalFileName.Equals("000000000.ipj", StringComparison.OrdinalIgnoreCase));
                if (mainIpj == null && _request.Source == CreateModuleSource.FromExistingProject)
                {
                    // Pour projets existants: premier IPJ a la racine
                    mainIpj = _files.FirstOrDefault(f => f.FileType == "IPJ" && 
                        string.IsNullOrEmpty(Path.GetDirectoryName(f.RelativePath)));
                }
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

            var displayed = (DgFiles.ItemsSource as IEnumerable<FileRenameItem>)?.Count() ?? _files.Count;
            var selected = _files.Count(f => f.IsSelected);
            TxtFileCount.Text = $"{selected}/{_files.Count} selectionnes";

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
                _startTime = DateTime.Now;
                _pausedTime = TimeSpan.Zero;
                UpdateProgress(0, "Initialisation du Copy Design...");
                AddLog("Début de la création du module...", "START");

                // Préparer les données depuis les ComboBox
                var initialeDessinateur = CmbInitialeDessinateur?.SelectedItem?.ToString() ?? "";
                var initialeCoDessinateur = CmbInitialeCoDessinateur?.SelectedItem?.ToString() ?? "";
                var initialeLeadCAD = CmbInitialeLeadCAD?.SelectedItem?.ToString() ?? "";
                
                // Si N/A est sélectionné, on le traite comme vide
                _request.InitialeDessinateur = initialeDessinateur == "N/A" ? "" : initialeDessinateur;
                _request.InitialeCoDessinateur = initialeCoDessinateur == "N/A" ? "" : initialeCoDessinateur;
                _request.InitialeLeadCAD = initialeLeadCAD == "N/A" ? "" : initialeLeadCAD;
                _request.JobTitle = TxtJobTitle?.Text ?? "";
                _request.CreationDate = DpCreationDate.SelectedDate ?? DateTime.Now;
                _request.FilesToCopy = _files;

                AddLog($"Dessinateur: {_request.InitialeDessinateur}", "INFO");
                if (!string.IsNullOrEmpty(_request.InitialeCoDessinateur))
                    AddLog($"Co-Dessinateur: {_request.InitialeCoDessinateur}", "INFO");
                if (!string.IsNullOrEmpty(_request.InitialeLeadCAD))
                    AddLog($"Lead CAD: {_request.InitialeLeadCAD}", "INFO");
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
                        // Extraire le nom du fichier du statusText si présent
                        string currentFile = "";
                        if (statusText.Contains(":"))
                        {
                            var parts = statusText.Split(':');
                            if (parts.Length > 1)
                            {
                                currentFile = parts[parts.Length - 1].Trim();
                            }
                        }
                        UpdateProgress(percent, statusText, false, currentFile);
                    }))
                {
                    // Initialiser la connexion à Inventor
                    UpdateProgress(2, "Connexion à Inventor...");
                    AddLog("Connexion à Inventor...", "INFO");
                    if (!copyDesignService.Initialize())
                    {
                        UpdateProgress(0, "[-] Erreur: Inventor non disponible", isError: true);
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
                        
                        // ═══════════════════════════════════════════════════════════════════════
                        // [+] REMPLISSAGE PDFs DE COUVERTURE BATCHPRINT
                        // Remplit automatiquement les 11 PDFs de couverture avec:
                        // - Numero de projet (NUMBER)
                        // - Reference REF (Dropdown7)
                        // - Module MOD (Dropdown10)  
                        // - Job Title (Nom of the job)
                        // ═══════════════════════════════════════════════════════════════════════
                        Logger.Info("═══════════════════════════════════════════════════════════════");
                        Logger.Info("[>] APPEL PdfCoverService.FillAllCoverPdfs()");
                        Logger.Info($"    DestinationPath: {_request.DestinationPath}");
                        Logger.Info($"    Project: {_request.Project}");
                        Logger.Info($"    Reference: {_request.Reference}");
                        Logger.Info($"    Module: {_request.Module}");
                        Logger.Info($"    JobTitle: {_request.JobTitle}");
                        Logger.Info("═══════════════════════════════════════════════════════════════");
                        
                        // Verification du chemin PDF
                        var pdfFolder = System.IO.Path.Combine(_request.DestinationPath, "6-Shop Drawing PDF", "Production");
                        Logger.Info($"[DEBUG] Chemin PDF complet: {pdfFolder}");
                        Logger.Info($"[DEBUG] Dossier existe: {System.IO.Directory.Exists(pdfFolder)}");
                        
                        if (System.IO.Directory.Exists(pdfFolder))
                        {
                            var pdfFiles = System.IO.Directory.GetFiles(pdfFolder, "*.pdf");
                            Logger.Info($"[DEBUG] Nombre de PDFs trouves: {pdfFiles.Length}");
                            foreach (var pdf in pdfFiles.Take(5))
                            {
                                Logger.Info($"[DEBUG]   - {System.IO.Path.GetFileName(pdf)}");
                            }
                        }
                        
                        try
                        {
                            AddLog("[>] Demarrage remplissage PDFs de couverture...", "START");
                            Logger.Info("[>] Creation de PdfCoverService...");
                            
                            var pdfService = new Services.PdfCoverService((msg, lvl) => {
                                AddLog(msg, lvl);
                                Logger.Info($"[PdfCoverService] {msg}");
                            });
                            
                            Logger.Info("[>] Appel FillAllCoverPdfs...");
                            int pdfCount = pdfService.FillAllCoverPdfs(
                                _request.DestinationPath,
                                _request.Project,
                                _request.Reference,
                                _request.Module,
                                _request.JobTitle);
                            
                            Logger.Info($"[+] PdfCoverService retourne: {pdfCount} PDFs remplis");
                            AddLog($"[+] Resultat: {pdfCount} PDFs de couverture traites", pdfCount > 0 ? "SUCCESS" : "WARN");
                        }
                        catch (Exception pdfEx)
                        {
                            AddLog($"[-] Exception PDFs couverture: {pdfEx.Message}", "ERROR");
                            Logger.Error($"[-] Exception PdfCoverService: {pdfEx}");
                            Logger.Error($"    StackTrace: {pdfEx.StackTrace}");
                        }

                        // Ouvrir automatiquement le dossier de destination
                        AddLog($"Ouverture du dossier: {_request.DestinationPath}", "INFO");
                        System.Diagnostics.Process.Start("explorer.exe", _request.DestinationPath);
                        
                        AddLog("✓ Note: Utilisez l'outil principal pour uploader vers Vault", "INFO");
                        
                        // Attendre un peu pour que l'utilisateur voie le message de succès
                        await Task.Delay(1500);
                        
                        // Supprimer le dossier temporaire Vault si présent
                        CleanupTempVaultFolder();
                        
                        DialogResult = true;
                    }
                    else
                    {
                        UpdateProgress(0, $"[-] Erreur: {result.ErrorMessage}", isError: true);
                        AddLog($"ERREUR: {result.ErrorMessage}", "ERROR");
                        AddLog("ACTION", "[!] Verifiez les parametres et reessayez");
                        
                        // Supprimer le dossier temporaire Vault même en cas d'erreur
                        CleanupTempVaultFolder();
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog($"ERREUR CRITIQUE: {ex.Message}", "ERROR");
                UpdateProgress(0, $"[-] Erreur critique: {ex.Message}", isError: true);
                AddLog("ACTION", "[!] Une erreur inattendue s'est produite. Verifiez Inventor et reessayez.");
                
                // Supprimer le dossier temporaire Vault en cas d'erreur
                CleanupTempVaultFolder();
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
            if (BtnCreateModule == null) return;

            var validation = ValidateAllInputs();
            BtnCreateModule.IsEnabled = validation.IsValid;
            
            // Afficher message READY ou ACTION selon la validation
            if (validation.IsValid)
            {
                AddLog("READY", "[+] Pret pour la creation du module");
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
        /// Verifie si un fichier .ipj correspond au pattern du fichier projet principal
        /// Pattern: XXXXX-XX-XX_2026.ipj, 000000000.ipj (nouveau template), ou similaire
        /// </summary>
        private bool IsMainProjectFilePattern(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return false;
            
            // Le fichier projet principal contient generalement "_2026" ou "_202" dans le nom
            // Exemples: XXXXX-XX-XX_2026.ipj, Module_2026.ipj, etc.
            var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            
            // Pattern 1: Nouveau template vierge 000000000.ipj
            if (nameWithoutExt.Equals("000000000", StringComparison.OrdinalIgnoreCase))
                return true;
            
            // Pattern 2: Contient _202X (annee)
            if (nameWithoutExt.Contains("_202"))
                return true;
            
            // Pattern 3: Format XXXXX-XX-XX (numero de projet avec tirets)
            if (System.Text.RegularExpressions.Regex.IsMatch(nameWithoutExt, @"^\d{5}-\d{2}-\d{2}"))
                return true;
            
            // Pattern 4: Le nom contient "Module" (fichier projet du module)
            if (nameWithoutExt.IndexOf("Module", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            
            // Pattern 5: Format XXXXXXXXX (9 chiffres - numero de projet complet)
            if (System.Text.RegularExpressions.Regex.IsMatch(nameWithoutExt, @"^\d{9}$"))
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
    }}

