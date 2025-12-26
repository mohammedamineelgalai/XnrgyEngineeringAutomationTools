using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Views;

namespace XnrgyEngineeringAutomationTools
{
    public partial class MainWindow : Window
    {
        private readonly VaultSdkService _vaultService;
        private readonly InventorService _inventorService;
        private bool _isVaultConnected = false;
        private bool _isInventorConnected = false;
        private bool _isDarkTheme = true;

        // === CHEMINS DES APPLICATIONS ===
        private readonly string _basePath = @"c:\Users\mohammedamine.elgala\source\repos";
        private string VaultUploadExePath => Path.Combine(_basePath, @"VaultAutomationTool\bin\Release\VaultAutomationTool.exe");
        private string DXFVerifierExePath => Path.Combine(_basePath, @"DXFVerifier\bin\Release\DXFVerifier.exe");
        private string ChecklistHVACPath => Path.Combine(_basePath, @"ChecklistHVAC\Checklist HVACAHU - By Mohammed Amine Elgalai.html");

        private readonly string[] _workspaceFolders = new[]
        {
            "$/Content Center Files",
            "$/Engineering/Inventor_Standards",
            "$/Engineering/Library/Cabinet",
            "$/Engineering/Library/Xnrgy_M99",
            "$/Engineering/Library/Xnrgy_Module"
        };

        private readonly SolidColorBrush _greenBrush = new SolidColorBrush(Color.FromRgb(16, 124, 16));
        private readonly SolidColorBrush _redBrush = new SolidColorBrush(Color.FromRgb(232, 17, 35));

        public MainWindow()
        {
            InitializeComponent();
            Logger.Initialize();
            _vaultService = new VaultSdkService();
            _inventorService = new InventorService();
        }
        
        /// <summary>
        /// Ajoute une entree au journal avec couleur et clignotement pour les erreurs
        /// </summary>
        private void AddLog(string message, string level = "INFO")
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string icon = level switch
            {
                "ERROR" => "X",
                "WARN" => "!",
                "SUCCESS" => "OK",
                "START" => ">>",
                "STOP" => "||",
                "CRITICAL" => "XX",
                _ => "i"
            };
            
            string text = "[" + timestamp + "] " + icon + " " + message;
            
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
                        StartBlinkAnimation(textBlock);
                        break;
                    case "WARN":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(255, 180, 0));
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    case "SUCCESS":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(80, 200, 80));
                        break;
                    case "START":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(100, 180, 255));
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    case "STOP":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(180, 100, 255));
                        break;
                    default:
                        // INFO - couleur selon le thème (blanc foncé ou noir foncé)
                        textBlock.Foreground = _isDarkTheme 
                            ? new SolidColorBrush(Color.FromRgb(220, 220, 220))  // Blanc léger pour dark
                            : new SolidColorBrush(Color.FromRgb(40, 40, 40));    // Noir foncé pour light
                        break;
                }
                
                LogListBox.Items.Add(textBlock);
                LogListBox.ScrollIntoView(textBlock);
                
                while (LogListBox.Items.Count > 150)
                {
                    LogListBox.Items.RemoveAt(0);
                }
            });
            
            Logger.Log(message, level == "ERROR" || level == "CRITICAL" ? Logger.LogLevel.ERROR : Logger.LogLevel.INFO);
        }
        
        private void UpdateLogColors()
        {
            // Met à jour les couleurs des logs INFO existants selon le thème
            var infoBrush = _isDarkTheme 
                ? new SolidColorBrush(Color.FromRgb(220, 220, 220))  // Blanc léger pour dark
                : new SolidColorBrush(Color.FromRgb(40, 40, 40));    // Noir foncé pour light
            
            foreach (var item in LogListBox.Items)
            {
                if (item is TextBlock tb)
                {
                    // Vérifier si c'est un log INFO (couleur grise/blanche)
                    var brush = tb.Foreground as SolidColorBrush;
                    if (brush != null)
                    {
                        var color = brush.Color;
                        // Si c'est une couleur grise/blanc/noir (INFO), la mettre à jour
                        bool isInfoColor = (color.R == color.G && color.G == color.B) || 
                                          (color.R >= 180 && color.G >= 180 && color.B >= 180) ||
                                          (color.R <= 60 && color.G <= 60 && color.B <= 60);
                        if (isInfoColor && tb.FontWeight != FontWeights.Bold && tb.FontWeight != FontWeights.SemiBold)
                        {
                            tb.Foreground = infoBrush;
                        }
                    }
                }
            }
        }
        
        private void StartBlinkAnimation(TextBlock textBlock)
        {
            var animation = new ColorAnimation
            {
                From = Color.FromRgb(255, 80, 80),
                To = Color.FromRgb(255, 200, 200),
                Duration = TimeSpan.FromMilliseconds(300),
                AutoReverse = true,
                RepeatBehavior = new RepeatBehavior(5)
            };
            
            var brush = new SolidColorBrush(Color.FromRgb(255, 80, 80));
            textBlock.Foreground = brush;
            brush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
        }
        
        private void ClearLogButton_Click(object sender, RoutedEventArgs e)
        {
            LogListBox.Items.Clear();
            AddLog("Journal efface", "INFO");
        }
        
        private void ThemeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            _isDarkTheme = !_isDarkTheme;
            
            var whiteBrush = Brushes.White;
            var darkBgBrush = new SolidColorBrush(Color.FromRgb(37, 37, 54));
            var lightBgBrush = new SolidColorBrush(Color.FromRgb(250, 250, 250));
            var darkTextBrush = new SolidColorBrush(Color.FromRgb(30, 30, 30));
            var lightDescText = new SolidColorBrush(Color.FromRgb(220, 220, 220)); // Blanc léger pour descriptions (dark)
            var darkDescText = new SolidColorBrush(Color.FromRgb(50, 50, 50)); // Noir foncé pour descriptions (light)
            
            // Couleurs boutons
            var btnDarkBg = new SolidColorBrush(Color.FromRgb(37, 37, 54));
            var btnLightBg = new SolidColorBrush(Color.FromRgb(245, 245, 248));
            var btnLightBorder = new SolidColorBrush(Color.FromRgb(200, 200, 210));
            var btnDarkBorder = new SolidColorBrush(Color.FromRgb(64, 64, 96));
            
            // Status indicators
            var statusDarkBg = new SolidColorBrush(Color.FromRgb(37, 37, 54));
            var statusLightBg = new SolidColorBrush(Color.FromRgb(245, 245, 248));
            
            if (_isDarkTheme)
            {
                // Theme SOMBRE
                MainWindowRoot.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46));
                ConnectionBorder.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212)); // Bleu
                LogBorder.Background = new SolidColorBrush(Color.FromRgb(18, 18, 28));
                LogHeaderBorder.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46));
                StatusBorder.Background = new SolidColorBrush(Color.FromRgb(37, 37, 54));
                
                // GroupBox
                GroupBox1.Background = darkBgBrush;
                GroupBox2.Background = darkBgBrush;
                GroupBox1.BorderBrush = btnDarkBorder;
                GroupBox2.BorderBrush = btnDarkBorder;
                
                // Textes header et status
                TitleText.Foreground = whiteBrush;
                StatusText.Foreground = whiteBrush;
                CopyrightText.Foreground = new SolidColorBrush(Color.FromRgb(200, 200, 200));
                
                // Status indicators
                VaultStatusBorder.Background = statusDarkBg;
                VaultLabel.Foreground = whiteBrush;
                VaultStatusText.Foreground = whiteBrush;
                InventorStatusBorder.Background = statusDarkBg;
                InventorLabel.Foreground = whiteBrush;
                InventorStatusText.Foreground = whiteBrush;
                
                // Journal header
                LogHeaderText.Foreground = whiteBrush;
                
                // Boutons - fond sombre
                BtnVaultUpload.Background = btnDarkBg;
                BtnPackAndGo.Background = btnDarkBg;
                BtnSmartTools.Background = btnDarkBg;
                BtnBuildModule.Background = btnDarkBg;
                BtnDXFVerifier.Background = btnDarkBg;
                BtnChecklistHVAC.Background = btnDarkBg;
                
                // Boutons - titres et descriptions (BLANC - pas de gris)
                BtnVaultUploadTitle.Foreground = whiteBrush;
                BtnVaultUploadDesc.Foreground = lightDescText;
                BtnPackAndGoTitle.Foreground = whiteBrush;
                BtnPackAndGoDesc.Foreground = lightDescText;
                BtnSmartToolsTitle.Foreground = whiteBrush;
                BtnSmartToolsDesc.Foreground = lightDescText;
                BtnBuildModuleTitle.Foreground = whiteBrush;
                BtnBuildModuleDesc.Foreground = lightDescText;
                BtnDXFVerifierTitle.Foreground = whiteBrush;
                BtnDXFVerifierDesc.Foreground = lightDescText;
                BtnChecklistHVACTitle.Foreground = whiteBrush;
                BtnChecklistHVACDesc.Foreground = lightDescText;
                
                // Bouton Theme - fond sombre pour thème sombre
                ThemeToggleButton.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46));
                ThemeToggleButton.Foreground = whiteBrush;
                ThemeToggleButton.Content = "☀️ Theme Clair";
                
                // Bouton Update - vert
                UpdateButton.Background = new SolidColorBrush(Color.FromRgb(16, 124, 16));
                UpdateButton.Foreground = whiteBrush;
                
                AddLog("Theme sombre active", "INFO");
            }
            else
            {
                // Theme CLAIR - Couleurs élégantes
                MainWindowRoot.Background = new SolidColorBrush(Color.FromRgb(245, 247, 250)); // Bleu-gris très clair
                ConnectionBorder.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212)); // Bleu reste
                LogBorder.Background = new SolidColorBrush(Color.FromRgb(252, 253, 255)); // Blanc bleuté
                LogHeaderBorder.Background = new SolidColorBrush(Color.FromRgb(235, 240, 248)); // Bleu très pâle
                StatusBorder.Background = new SolidColorBrush(Color.FromRgb(235, 240, 248)); // Bleu très pâle
                
                // GroupBox - fond blanc avec teinte bleutée
                GroupBox1.Background = new SolidColorBrush(Color.FromRgb(252, 253, 255));
                GroupBox2.Background = new SolidColorBrush(Color.FromRgb(252, 253, 255));
                GroupBox1.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 215, 235)); // Bordure bleue douce
                GroupBox2.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 215, 235));
                
                // Textes - noir foncé
                TitleText.Foreground = whiteBrush; // Reste blanc sur fond bleu
                StatusText.Foreground = darkTextBrush;
                CopyrightText.Foreground = new SolidColorBrush(Color.FromRgb(40, 40, 40));
                
                // Status indicators - fond bleu très pâle
                VaultStatusBorder.Background = new SolidColorBrush(Color.FromRgb(235, 242, 252));
                VaultLabel.Foreground = darkTextBrush;
                VaultStatusText.Foreground = darkTextBrush;
                InventorStatusBorder.Background = new SolidColorBrush(Color.FromRgb(235, 242, 252));
                InventorLabel.Foreground = darkTextBrush;
                InventorStatusText.Foreground = darkTextBrush;
                
                // Journal header
                LogHeaderText.Foreground = darkTextBrush;
                
                // Boutons - fond bleu très pâle élégant
                var btnLightBgElegant = new SolidColorBrush(Color.FromRgb(240, 245, 252));
                BtnVaultUpload.Background = btnLightBgElegant;
                BtnPackAndGo.Background = btnLightBgElegant;
                BtnSmartTools.Background = btnLightBgElegant;
                BtnBuildModule.Background = btnLightBgElegant;
                BtnDXFVerifier.Background = btnLightBgElegant;
                BtnChecklistHVAC.Background = btnLightBgElegant;
                
                // Boutons - titres et descriptions (NOIR FONCE - pas de gris)
                BtnVaultUploadTitle.Foreground = darkTextBrush;
                BtnVaultUploadDesc.Foreground = darkDescText;
                BtnPackAndGoTitle.Foreground = darkTextBrush;
                BtnPackAndGoDesc.Foreground = darkDescText;
                BtnSmartToolsTitle.Foreground = darkTextBrush;
                BtnSmartToolsDesc.Foreground = darkDescText;
                BtnBuildModuleTitle.Foreground = darkTextBrush;
                BtnBuildModuleDesc.Foreground = darkDescText;
                BtnDXFVerifierTitle.Foreground = darkTextBrush;
                BtnDXFVerifierDesc.Foreground = darkDescText;
                BtnChecklistHVACTitle.Foreground = darkTextBrush;
                BtnChecklistHVACDesc.Foreground = darkDescText;
                
                // Bouton Theme - fond bleu pâle élégant pour thème clair
                ThemeToggleButton.Background = new SolidColorBrush(Color.FromRgb(235, 240, 250));
                ThemeToggleButton.Foreground = darkTextBrush;
                ThemeToggleButton.Content = "🌙 Theme Sombre";
                
                // Bouton Update - vert brillant
                UpdateButton.Background = new SolidColorBrush(Color.FromRgb(16, 185, 16));
                UpdateButton.Foreground = Brushes.White;
                
                AddLog("Theme clair active", "INFO");
            }
            
            // Mettre à jour les couleurs des logs existants
            UpdateLogColors();
            
            // Mettre à jour la couleur du bouton Connecter selon son état
            UpdateConnectionStatus();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            AddLog("===============================================================", "INFO");
            AddLog("XNRGY ENGINEERING AUTOMATION TOOLS v1.0", "START");
            AddLog("Developpe par Mohammed Amine Elgalai - XNRGY Climate Systems", "INFO");
            AddLog("===============================================================", "INFO");
            AddLog("Initialisation des services...", "INFO");
            UpdateConnectionStatus();
            TryConnectInventorAuto();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            AddLog("Fermeture de l'application...", "STOP");
            if (_isVaultConnected) _vaultService?.Disconnect();
            _inventorService?.Disconnect();
        }

        private void TryConnectInventorAuto()
        {
            AddLog("Recherche d'Inventor en cours...", "INFO");
            try
            {
                if (_inventorService.TryConnect())
                {
                    _isInventorConnected = true;
                    UpdateConnectionStatus();
                    AddLog("Connexion automatique a Inventor reussie!", "SUCCESS");
                }
                else
                {
                    AddLog("Inventor non detecte - Demarrez Inventor pour activer les fonctions", "WARN");
                }
            }
            catch (Exception ex)
            {
                AddLog("Inventor non disponible: " + ex.Message, "WARN");
            }
        }

        private void UpdateConnectionStatus()
        {
            // Couleurs pour bouton Connecter
            var redBtnBrush = new SolidColorBrush(Color.FromRgb(232, 17, 35));      // Rouge clair #E81123
            var greenBtnBrush = new SolidColorBrush(Color.FromRgb(16, 185, 16));    // Vert brillant #10B910
            
            if (_isVaultConnected && _vaultService.IsConnected)
            {
                VaultIndicator.Fill = _greenBrush;
                VaultStatusText.Text = _vaultService.VaultName + " (" + _vaultService.UserName + ")";
                ConnectButton.Content = "Deconnecter";
                ConnectButton.Background = greenBtnBrush;  // Vert brillant quand connecté
            }
            else
            {
                VaultIndicator.Fill = _redBrush;
                VaultStatusText.Text = "Deconnecte";
                ConnectButton.Content = "Connecter";
                ConnectButton.Background = redBtnBrush;    // Rouge clair quand déconnecté
            }

            if (_isInventorConnected && _inventorService.IsConnected)
            {
                InventorIndicator.Fill = _greenBrush;
                string docName = _inventorService.GetActiveDocumentName();
                InventorStatusText.Text = !string.IsNullOrEmpty(docName) ? docName : "Connecte";
            }
            else
            {
                InventorIndicator.Fill = _redBrush;
                InventorStatusText.Text = "Deconnecte";
            }
        }

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isVaultConnected)
            {
                AddLog("Deconnexion de Vault...", "INFO");
                _vaultService.Disconnect();
                _isVaultConnected = false;
                StatusText.Text = "Deconnecte de Vault";
                AddLog("Deconnecte de Vault", "SUCCESS");
            }
            else
            {
                AddLog("Ouverture de la fenetre de connexion Vault...", "INFO");
                var loginWindow = new LoginWindow(_vaultService);
                loginWindow.Owner = this;
                if (loginWindow.ShowDialog() == true)
                {
                    _isVaultConnected = true;
                    StatusText.Text = "Connecte a " + _vaultService.VaultName;
                    AddLog("Connecte a Vault: " + _vaultService.VaultName + " (" + _vaultService.UserName + ")", "SUCCESS");
                    AddLog("Cliquez sur Update Workspace pour synchroniser les bibliotheques", "INFO");
                }
                else
                {
                    AddLog("Connexion Vault annulee", "WARN");
                }
            }
            UpdateConnectionStatus();
        }

        private void OpenVaultUpload_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Lancement de Vault Upload Tool...", "START");
            StatusText.Text = "Lancement de Vault Upload...";
            try
            {
                if (File.Exists(VaultUploadExePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = VaultUploadExePath,
                        UseShellExecute = true
                    });
                    AddLog("VaultAutomationTool lance avec succes", "SUCCESS");
                    AddLog("Chemin: " + VaultUploadExePath, "INFO");
                    StatusText.Text = "Vault Upload lance";
                }
                else
                {
                    AddLog("ERREUR: Application non trouvee!", "CRITICAL");
                    AddLog("Chemin attendu: " + VaultUploadExePath, "ERROR");
                    AddLog("Compilez le projet VaultAutomationTool d'abord", "WARN");
                    StatusText.Text = "Erreur: Application non trouvee";
                }
            }
            catch (Exception ex)
            {
                AddLog("Erreur lancement VaultUpload: " + ex.Message, "CRITICAL");
                StatusText.Text = "Erreur";
            }
        }

        private void OpenPackAndGo_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Ouverture de l'outil Créer Module...", "START");
            StatusText.Text = "Création de module...";
            
            try
            {
                var createModuleWindow = new CreateModuleWindow();
                createModuleWindow.Owner = this;
                
                AddLog("Fenêtre Créer Module ouverte", "INFO");
                AddLog("Options disponibles:", "INFO");
                AddLog("  - Créer depuis Template ($/Engineering/Library)", "INFO");
                AddLog("  - Créer depuis Projet Existant", "INFO");
                
                var result = createModuleWindow.ShowDialog();
                
                if (result == true)
                {
                    AddLog("Module créé avec succès!", "SUCCESS");
                    StatusText.Text = "Module créé";
                }
                else
                {
                    AddLog("Création de module annulée", "INFO");
                    StatusText.Text = "Création annulée";
                }
            }
            catch (Exception ex)
            {
                AddLog("Erreur ouverture Créer Module: " + ex.Message, "ERROR");
                StatusText.Text = "Erreur";
            }
        }

        private void OpenSmartTools_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Smart Tools - Module en developpement", "WARN");
            AddLog("Fonctionnalites prevues (depuis SmartToolsAmineAddin):", "INFO");
            AddLog("  - Creation IPT automatique", "INFO");
            AddLog("  - Export STEP batch", "INFO");
            AddLog("  - Generation PDF", "INFO");
            AddLog("  - iLogic Forms integres", "INFO");
            StatusText.Text = "Smart Tools - En developpement";
        }
        
        private void OpenBuildModule_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Build Module - Fonctionnalite future", "WARN");
            AddLog("Fonctionnalites prevues:", "INFO");
            AddLog("  - Assemblage automatique de modules HVAC", "INFO");
            AddLog("  - Configuration parametrique", "INFO");
            AddLog("  - Generation BOM automatique", "INFO");
            AddLog("  - Integration avec standards XNRGY", "INFO");
            StatusText.Text = "Build Module - En developpement";
        }

        private void OpenDXFVerifier_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Lancement de DXF Verifier...", "START");
            StatusText.Text = "Lancement de DXF Verifier...";
            try
            {
                if (File.Exists(DXFVerifierExePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = DXFVerifierExePath,
                        UseShellExecute = true
                    });
                    AddLog("DXFVerifier lance avec succes", "SUCCESS");
                    AddLog("Chemin: " + DXFVerifierExePath, "INFO");
                    StatusText.Text = "DXF Verifier lance";
                }
                else
                {
                    AddLog("ERREUR: Application non trouvee!", "CRITICAL");
                    AddLog("Chemin attendu: " + DXFVerifierExePath, "ERROR");
                    AddLog("Compilez le projet DXFVerifier d'abord", "WARN");
                    StatusText.Text = "Erreur: Application non trouvee";
                }
            }
            catch (Exception ex)
            {
                AddLog("Erreur lancement DXFVerifier: " + ex.Message, "CRITICAL");
                StatusText.Text = "Erreur";
            }
        }

        private void OpenChecklistHVAC_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Ouverture de Checklist HVAC...", "START");
            StatusText.Text = "Ouverture de Checklist HVAC...";
            try
            {
                if (File.Exists(ChecklistHVACPath))
                {
                    // Ouvrir dans une fenêtre intégrée
                    var checklistWindow = new Views.ChecklistHVACWindow(ChecklistHVACPath);
                    checklistWindow.Show();
                    AddLog("Checklist HVAC ouvert dans l'application", "SUCCESS");
                    AddLog("Fichier: " + ChecklistHVACPath, "INFO");
                    StatusText.Text = "Checklist HVAC ouvert";
                }
                else
                {
                    AddLog("ERREUR: Fichier non trouve!", "CRITICAL");
                    AddLog("Chemin attendu: " + ChecklistHVACPath, "ERROR");
                    StatusText.Text = "Erreur: Fichier non trouve";
                }
            }
            catch (Exception ex)
            {
                AddLog("Erreur ouverture ChecklistHVAC: " + ex.Message, "CRITICAL");
                StatusText.Text = "Erreur";
            }
        }

        private async void UpdateWorkspace_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckVaultConnection()) return;
            
            AddLog("===============================================================", "INFO");
            AddLog("DEMARRAGE MISE A JOUR DU WORKSPACE", "START");
            AddLog("===============================================================", "INFO");
            StatusText.Text = "Mise a jour du Workspace...";
            ConnectButton.IsEnabled = false;
            
            try
            {
                int success = 0;
                int total = _workspaceFolders.Length;
                
                foreach (var folder in _workspaceFolders)
                {
                    AddLog("GET: " + folder + "...", "INFO");
                    StatusText.Text = "GET: " + folder;
                    
                    if (await _vaultService.GetFolderAsync(folder))
                    {
                        success++;
                        AddLog("GET reussi: " + folder, "SUCCESS");
                    }
                    else
                    {
                        AddLog("GET echoue: " + folder, "ERROR");
                    }
                }
                
                AddLog("===============================================================", "INFO");
                if (success == total)
                {
                    AddLog("MISE A JOUR TERMINEE: " + success + "/" + total + " dossiers", "SUCCESS");
                }
                else
                {
                    AddLog("MISE A JOUR PARTIELLE: " + success + "/" + total + " dossiers", "WARN");
                }
                AddLog("===============================================================", "INFO");
                StatusText.Text = "Workspace mis a jour (" + success + "/" + total + ")";
            }
            catch (Exception ex)
            {
                AddLog("ERREUR CRITIQUE: " + ex.Message, "CRITICAL");
            }
            finally
            {
                ConnectButton.IsEnabled = true;
            }
        }

        private bool CheckVaultConnection()
        {
            if (!_isVaultConnected || !_vaultService.IsConnected)
            {
                AddLog("CONNEXION VAULT REQUISE", "CRITICAL");
                AddLog("Cliquez sur Connecter pour vous connecter a Vault", "WARN");
                StatusText.Text = "Connexion Vault requise";
                return false;
            }
            return true;
        }

        private bool CheckInventorConnection()
        {
            if (!_isInventorConnected || !_inventorService.IsConnected)
            {
                AddLog("Inventor non connecte - Tentative de connexion...", "WARN");
                TryConnectInventorAuto();
                return _isInventorConnected;
            }
            return true;
        }
    }
}
