using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Shared.Views;
using XnrgyEngineeringAutomationTools.Modules.CreateModule.Views;
using XnrgyEngineeringAutomationTools.Modules.UploadTemplate.Views;
using XnrgyEngineeringAutomationTools.Modules.ChecklistHVAC.Views;
using XnrgyEngineeringAutomationTools.Modules.SmartTools.Views;

namespace XnrgyEngineeringAutomationTools
{
    public partial class MainWindow : Window
    {
        private readonly VaultSdkService _vaultService;
        private readonly InventorService _inventorService;
        private bool _isVaultConnected = false;
        private bool _isInventorConnected = false;
        private bool _isDarkTheme = true;
        private System.Windows.Threading.DispatcherTimer? _inventorReconnectTimer;

        // === VERSIONS REQUISES ===
        private const string REQUIRED_INVENTOR_VERSION = "2026";  // Inventor Professional 2026.2
        private const string REQUIRED_VAULT_VERSION = "2026";     // Vault Professional 2026.2

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

        // === THEME GLOBAL (accessible par tous les sous-formulaires) ===
        public static bool CurrentThemeIsDark { get; private set; } = true;
        public static event Action<bool>? ThemeChanged;

        public MainWindow()
        {
            InitializeComponent();
            Logger.Initialize();
            _vaultService = new VaultSdkService();
            _inventorService = new InventorService();
            
            // Charger le theme sauvegarde
            _isDarkTheme = UserPreferencesManager.LoadTheme();
            CurrentThemeIsDark = _isDarkTheme;
            
            // Timer pour réessayer la connexion Inventor si elle échoue au démarrage
            // Utile quand l'app est lancée par script avant que COM soit prêt
            _inventorReconnectTimer = new System.Windows.Threading.DispatcherTimer();
            _inventorReconnectTimer.Interval = TimeSpan.FromSeconds(3);
            _inventorReconnectTimer.Tick += InventorReconnectTimer_Tick;
        }
        
        /// <summary>
        /// Ajoute une entree au journal avec couleur et clignotement pour les erreurs
        /// Utilise JournalColorService pour uniformite des couleurs
        /// </summary>
        private void AddLog(string message, string level = "INFO")
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string icon = level switch
            {
                "ERROR" => "[-]",
                "WARN" => "[!]",
                "SUCCESS" => "[+]",
                "START" => "[>]",
                "STOP" => "[~]",
                "CRITICAL" => "[X]",
                _ => "[i]"
            };
            
            string text = "[" + timestamp + "] " + icon + " " + message;
            
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
                
                // Utilise JournalColorService pour les couleurs uniformisees
                switch (level)
                {
                    case "ERROR":
                    case "CRITICAL":
                        textBlock.Foreground = Services.JournalColorService.ErrorBrush;
                        textBlock.FontWeight = FontWeights.Bold;
                        StartBlinkAnimation(textBlock);
                        break;
                    case "WARN":
                        textBlock.Foreground = Services.JournalColorService.WarningBrush;
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    case "SUCCESS":
                        textBlock.Foreground = Services.JournalColorService.SuccessBrush;
                        break;
                    case "START":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(100, 180, 255)); // Bleu clair pour START
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    case "STOP":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(180, 100, 255)); // Violet pour STOP
                        break;
                    default:
                        // INFO - Blanc pur depuis JournalColorService
                        textBlock.Foreground = Services.JournalColorService.InfoBrush;
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
            // Le journal a un fond noir FIXE, donc les logs INFO restent toujours blancs
            // Les autres couleurs (rouge erreur, vert succes, etc.) ne changent pas
            // Cette methode ne fait plus rien car le fond est fixe
        }
        
        private void StartBlinkAnimation(TextBlock textBlock)
        {
            // Utilise JournalColorService pour la couleur d'erreur
            var animation = new ColorAnimation
            {
                From = Services.JournalColorService.ErrorColor,
                To = Color.FromRgb(255, 200, 200),
                Duration = TimeSpan.FromMilliseconds(300),
                AutoReverse = true,
                RepeatBehavior = new RepeatBehavior(5)
            };
            
            var brush = new SolidColorBrush(Services.JournalColorService.ErrorColor);
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
            
            // Sauvegarder le theme et notifier les sous-formulaires
            CurrentThemeIsDark = _isDarkTheme;
            UserPreferencesManager.SaveTheme(_isDarkTheme);
            ThemeChanged?.Invoke(_isDarkTheme);
            
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
                ConnectionBorder.Background = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F - FIXE
                LogBorder.Background = new SolidColorBrush(Color.FromRgb(26, 26, 40)); // #1A1A28 - FIXE NOIR
                LogBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                LogHeaderBorder.Background = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                StatusBorder.Background = new SolidColorBrush(Color.FromRgb(37, 37, 54));
                
                // GroupBox
                GroupBox1.Background = darkBgBrush;
                GroupBox2.Background = darkBgBrush;
                GroupBox1.BorderBrush = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                GroupBox2.BorderBrush = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                
                // Titres GroupBox - BLANC sur fond sombre
                GroupBox1Header.Foreground = whiteBrush;
                GroupBox2Header.Foreground = whiteBrush;
                
                // Textes header et status
                TitleText.Foreground = whiteBrush;
                StatusText.Foreground = whiteBrush;
                CopyrightText.Foreground = new SolidColorBrush(Color.FromRgb(200, 200, 200));
                
                // Status indicators Vault/Inventor - FIXE NOIR
                VaultStatusBorder.Background = new SolidColorBrush(Color.FromRgb(26, 26, 40)); // #1A1A28 - FIXE NOIR
                VaultStatusText.Foreground = whiteBrush;
                InventorStatusBorder.Background = new SolidColorBrush(Color.FromRgb(26, 26, 40)); // #1A1A28 - FIXE NOIR
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
                BtnUploadTemplate.Background = btnDarkBg;
                
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
                BtnUploadTemplateTitle.Foreground = whiteBrush;
                BtnUploadTemplateDesc.Foreground = lightDescText;
                
                // Bouton Theme - fond sombre pour thème sombre
                ThemeToggleButton.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46));
                ThemeToggleButton.Foreground = whiteBrush;
                ThemeToggleButton.Content = "☀️ Theme Clair";
                
                // Bouton Update - violet
                UpdateButton.Background = new SolidColorBrush(Color.FromRgb(124, 58, 237)); // #7C3AED
                UpdateButton.Foreground = whiteBrush;
                
                AddLog("Theme sombre active", "INFO");
            }
            else
            {
                // Theme CLAIR - Couleurs élégantes
                MainWindowRoot.Background = new SolidColorBrush(Color.FromRgb(245, 247, 250)); // Bleu-gris très clair
                ConnectionBorder.Background = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F - FIXE
                LogBorder.Background = new SolidColorBrush(Color.FromRgb(26, 26, 40)); // #1A1A28 - FIXE NOIR
                LogBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                LogHeaderBorder.Background = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                StatusBorder.Background = new SolidColorBrush(Color.FromRgb(235, 240, 248)); // Bleu très pâle

                // GroupBox - fond blanc avec teinte bleutée
                GroupBox1.Background = new SolidColorBrush(Color.FromRgb(252, 253, 255));
                GroupBox2.Background = new SolidColorBrush(Color.FromRgb(252, 253, 255));
                GroupBox1.BorderBrush = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                GroupBox2.BorderBrush = new SolidColorBrush(Color.FromRgb(42, 74, 111)); // Bleu marine #2A4A6F
                
                // Titres GroupBox - BLANC sur fond bleu marine (header reste bleu marine meme en theme clair)
                GroupBox1Header.Foreground = whiteBrush;
                GroupBox2Header.Foreground = whiteBrush;
                
                // Textes - noir foncé
                TitleText.Foreground = whiteBrush; // Reste blanc sur fond bleu
                StatusText.Foreground = darkTextBrush;
                CopyrightText.Foreground = new SolidColorBrush(Color.FromRgb(40, 40, 40));
                
                // Status indicators Vault/Inventor - FIXE NOIR meme en theme clair (texte blanc)
                VaultStatusBorder.Background = new SolidColorBrush(Color.FromRgb(26, 26, 40)); // #1A1A28 - FIXE NOIR
                VaultStatusText.Foreground = whiteBrush;
                InventorStatusBorder.Background = new SolidColorBrush(Color.FromRgb(26, 26, 40)); // #1A1A28 - FIXE NOIR
                InventorLabel.Foreground = whiteBrush;
                InventorStatusText.Foreground = whiteBrush;
                
                // Journal header - reste BLANC sur barre bleu marine (meme en theme clair)
                LogHeaderText.Foreground = whiteBrush;
                
                // Boutons - fond bleu très pâle élégant
                var btnLightBgElegant = new SolidColorBrush(Color.FromRgb(240, 245, 252));
                BtnVaultUpload.Background = btnLightBgElegant;
                BtnPackAndGo.Background = btnLightBgElegant;
                BtnSmartTools.Background = btnLightBgElegant;
                BtnBuildModule.Background = btnLightBgElegant;
                BtnDXFVerifier.Background = btnLightBgElegant;
                BtnChecklistHVAC.Background = btnLightBgElegant;
                BtnUploadTemplate.Background = btnLightBgElegant;
                
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
                BtnUploadTemplateTitle.Foreground = darkTextBrush;
                BtnUploadTemplateDesc.Foreground = darkDescText;
                
                // Bouton Theme - fond bleu pâle élégant pour thème clair
                ThemeToggleButton.Background = new SolidColorBrush(Color.FromRgb(235, 240, 250));
                ThemeToggleButton.Foreground = darkTextBrush;
                ThemeToggleButton.Content = "🌙 Theme Sombre";
                
                // Bouton Update - violet
                UpdateButton.Background = new SolidColorBrush(Color.FromRgb(124, 58, 237)); // #7C3AED
                UpdateButton.Foreground = Brushes.White;
                
                AddLog("Theme clair active", "INFO");
            }
            
            // Mettre à jour les couleurs des logs existants
            UpdateLogColors();
            
            // Mettre à jour la couleur du bouton Connecter selon son état
            UpdateConnectionStatus();
        }

        /// <summary>
        /// Applique le theme actuel (appele au demarrage et lors du changement)
        /// </summary>
        private void ApplyCurrentTheme()
        {
            // Simuler un clic sur le bouton theme pour appliquer les couleurs
            // mais sans changer le theme (on le remet apres)
            bool currentTheme = _isDarkTheme;
            _isDarkTheme = !_isDarkTheme; // Inverser temporairement
            ThemeToggleButton_Click(this, new RoutedEventArgs()); // Cela remet le bon theme
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            AddLog("===============================================================", "INFO");
            AddLog("XNRGY ENGINEERING AUTOMATION TOOLS v1.0", "START");
            AddLog("Developpe par Mohammed Amine Elgalai - XNRGY Climate Systems", "INFO");
            AddLog("===============================================================", "INFO");
            
            // Appliquer le theme sauvegarde au demarrage
            if (!_isDarkTheme)
            {
                // Si le theme sauvegarde est clair, on doit l'appliquer
                // (par defaut l'UI est en mode sombre)
                _isDarkTheme = true; // Forcer sombre d'abord
                ThemeToggleButton_Click(this, new RoutedEventArgs()); // Passer en clair
            }
            
            // Lancer la checklist de démarrage
            Dispatcher.BeginInvoke(new Action(async () => await RunStartupChecklist()), 
                System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// Exécute la checklist de démarrage: vérifie/lance Inventor et ouvre connexion Vault
        /// </summary>
        private async Task RunStartupChecklist()
        {
            AddLog("", "INFO");
            AddLog("[i] CHECKLIST DE DEMARRAGE", "START");
            AddLog("───────────────────────────────────────────────", "INFO");

            // === ETAPE 1: Verification/Lancement d'Inventor ===
            AddLog("", "INFO");
            AddLog("[>] [1/3] Verification d'Inventor Professional 2026...", "INFO");
            
            bool inventorOk = await CheckAndLaunchInventorAsync();
            
            // === ETAPE 2: Verification/Lancement du Vault Client ===
            AddLog("", "INFO");
            AddLog("[>] [2/3] Verification du Vault Client 2026...", "INFO");
            
            bool vaultClientOk = await CheckAndLaunchVaultClientAsync();
            
            // === RESUME ===
            AddLog("", "INFO");
            AddLog("───────────────────────────────────────────────", "INFO");
            AddLog("[i] RESUME DE LA CHECKLIST:", "INFO");
            AddLog("   • Inventor 2026: " + (inventorOk ? "OK" : "EN COURS DE DEMARRAGE"), inventorOk ? "SUCCESS" : "WARN");
            AddLog("   • Vault Client:  " + (vaultClientOk ? "OK" : "EN COURS DE DEMARRAGE"), vaultClientOk ? "SUCCESS" : "WARN");
            AddLog("───────────────────────────────────────────────", "INFO");
            
            if (inventorOk && vaultClientOk)
            {
                AddLog("[OK] Environnement pret - Toutes les verifications reussies!", "SUCCESS");
            }
            else
            {
                AddLog("[~] Applications en cours de demarrage...", "INFO");
            }
            
            AddLog("", "INFO");
            UpdateConnectionStatus();
            
            // Si Inventor pas encore connecté, démarrer le timer de reconnexion automatique
            if (!_isInventorConnected || !_inventorService.IsConnected)
            {
                AddLog("[~] Timer de reconnexion Inventor active (toutes les 3 sec)...", "INFO");
                _inventorReconnectTimer?.Start();
            }
            
            // === ETAPE 3: Ouvrir la fenetre de connexion (mode auto-connect) ===
            AddLog("[>] [3/3] Connexion a Vault...", "INFO");
            await Task.Delay(500); // Petit délai pour laisser le temps aux apps de démarrer
            
            // Toujours ouvrir la fenetre de connexion avec auto-connect
            // Si credentials sauvegardes → connexion auto → fenetre se ferme
            // Sinon → fenêtre reste pour intervention utilisateur
            if (!_isVaultConnected)
            {
                AddLog("   [>] Ouverture de la fenetre de connexion...", "INFO");
                OpenLoginWindowWithAutoConnect();
            }
            else
            {
                AddLog("   [+] Deja connecte a Vault", "SUCCESS");
            }
        }
        
        /// <summary>
        /// Ouvre la fenetre de connexion avec tentative de connexion automatique
        /// Comportement pro: affiche le spinner, se ferme auto si succes, reste ouverte sinon
        /// </summary>
        private void OpenLoginWindowWithAutoConnect()
        {
            var loginWindow = new LoginWindow(_vaultService, autoConnect: true)
            {
                Owner = this
            };
            
            if (loginWindow.ShowDialog() == true)
            {
                _isVaultConnected = true;
                UpdateConnectionStatus();
                AddLog("[+] Connexion Vault etablie avec succes!", "SUCCESS");
                
                // Synchronisation automatique des parametres depuis Vault au demarrage
                SyncSettingsFromVaultAsync();
            }
            else
            {
                AddLog("[!] Connexion Vault annulee ou echouee", "WARN");
            }
        }
        
        /// <summary>
        /// Ouvre la fenetre de connexion Vault (mode manuel)
        /// </summary>
        private void OpenLoginWindow()
        {
            var loginWindow = new LoginWindow(_vaultService)
            {
                Owner = this
            };
            
            if (loginWindow.ShowDialog() == true)
            {
                _isVaultConnected = true;
                UpdateConnectionStatus();
                AddLog("[+] Connexion Vault etablie avec succes!", "SUCCESS");
                
                // Synchronisation automatique des parametres depuis Vault au demarrage
                SyncSettingsFromVaultAsync();
            }
        }

        /// <summary>
        /// Synchronise les parametres de l'application depuis Vault en arriere-plan
        /// Appele automatiquement apres connexion Vault reussie
        /// </summary>
        private async void SyncSettingsFromVaultAsync()
        {
            try
            {
                await Task.Run(() =>
                {
                    var settingsService = new VaultSettingsService(_vaultService);
                    bool synced = settingsService.SyncFromVault();
                    
                    Dispatcher.Invoke(() =>
                    {
                        if (synced)
                        {
                            AddLog("[+] Parametres synchronises depuis Vault", "SUCCESS");
                        }
                        else
                        {
                            AddLog("[i] Parametres locaux utilises (fichier Vault non trouve)", "INFO");
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                    AddLog($"[!] Erreur sync parametres: {ex.Message}", "WARN"));
            }
        }

        /// <summary>
        /// Verifie si Inventor est lance, sinon le lance
        /// </summary>
        private async Task<bool> CheckAndLaunchInventorAsync()
        {
            return await Task.Run(async () =>
            {
                try
                {
                    // Verifier si le processus Inventor est en cours
                    var inventorProcesses = Process.GetProcessesByName("Inventor");
                    
                    if (inventorProcesses.Length == 0)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            AddLog("   [!] Inventor n'est pas en cours d'execution", "WARN");
                            AddLog("   [>] Lancement d'Inventor Professional 2026...", "INFO");
                        });
                        
                        // Chercher et lancer Inventor
                        string inventorPath = FindInventorExecutable();
                        if (!string.IsNullOrEmpty(inventorPath))
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = inventorPath,
                                UseShellExecute = true
                            });
                            
                            Dispatcher.Invoke(() =>
                                AddLog("   [~] Inventor demarre, veuillez patienter...", "INFO"));
                            
                            // Attendre qu'Inventor demarre (processus visible)
                            bool processStarted = false;
                            for (int i = 0; i < 30; i++)
                            {
                                await Task.Delay(1000);
                                inventorProcesses = Process.GetProcessesByName("Inventor");
                                if (inventorProcesses.Length > 0)
                                {
                                    processStarted = true;
                                    Dispatcher.Invoke(() =>
                                        AddLog("   [+] Processus Inventor detecte!", "SUCCESS"));
                                    break;
                                }
                            }
                            
                            if (!processStarted)
                            {
                                Dispatcher.Invoke(() =>
                                    AddLog("   [-] Inventor n'a pas demarre", "ERROR"));
                                return false;
                            }
                            
                            // Attendre que COM soit pret (Inventor a besoin de temps pour s'initialiser)
                            Dispatcher.Invoke(() =>
                                AddLog("   [~] Attente initialisation COM (peut prendre 15-30 sec)...", "INFO"));
                            
                            // Tentatives de connexion COM avec attente progressive
                            for (int attempt = 1; attempt <= 6; attempt++)
                            {
                                await Task.Delay(5000); // Attendre 5 secondes entre chaque tentative
                                
                                Dispatcher.Invoke(() =>
                                    AddLog($"   [>] Tentative de connexion COM {attempt}/6...", "INFO"));
                                
                                if (_inventorService.TryConnect())
                                {
                                    _isInventorConnected = true;
                                    string version = _inventorService.GetInventorVersion();
                                    
                                    Dispatcher.Invoke(() =>
                                    {
                                        if (!string.IsNullOrEmpty(version))
                                            AddLog("   [i] Version: " + version, "INFO");
                                        AddLog("   [+] Inventor Professional - Connexion COM etablie!", "SUCCESS");
                                        UpdateConnectionStatus();
                                    });
                                    return true;
                                }
                            }
                            
                            // Si toujours pas connecte apres 6 tentatives (30 sec)
                            Dispatcher.Invoke(() =>
                            {
                                AddLog("   [!] Inventor demarre mais COM pas encore pret", "WARN");
                                AddLog("   -> L'application reessayera automatiquement", "INFO");
                            });
                            return false;
                        }
                        else
                        {
                            Dispatcher.Invoke(() =>
                                AddLog("   [-] Inventor 2026 non trouve sur ce poste", "ERROR"));
                            return false;
                        }
                    }
                    
                    Dispatcher.Invoke(() => 
                        AddLog("   [>] Processus Inventor deja en cours", "INFO"));
                    
                    // Inventor deja lance - connexion directe
                    if (_inventorService.TryConnect())
                    {
                        _isInventorConnected = true;
                        
                        // Recuperer la version
                        string version = _inventorService.GetInventorVersion();
                        
                        Dispatcher.Invoke(() =>
                        {
                            if (!string.IsNullOrEmpty(version))
                            {
                                AddLog("   [i] Version: " + version, "INFO");
                                AddLog("   [+] Inventor Professional - Connexion etablie!", "SUCCESS");
                            }
                            else
                            {
                                AddLog("   [+] Connexion a Inventor etablie", "SUCCESS");
                            }
                            // Mettre a jour l'indicateur de connexion
                            UpdateConnectionStatus();
                        });
                        
                        return true;
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            AddLog("   [!] Inventor en cours mais connexion COM echouee", "WARN");
                            AddLog("   -> Verifiez qu'Inventor est completement charge", "INFO");
                        });
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AddLog("   [-] Erreur verification Inventor: " + ex.Message, "ERROR");
                    });
                    return false;
                }
            });
        }
        
        /// <summary>
        /// Trouve l'exécutable Inventor sur le système
        /// </summary>
        private string FindInventorExecutable()
        {
            // Chemins standards pour Inventor 2026
            string[] possiblePaths = new[]
            {
                @"C:\Program Files\Autodesk\Inventor 2026\Bin\Inventor.exe",
                @"C:\Program Files\Autodesk\Inventor 2025\Bin\Inventor.exe",
                @"C:\Program Files\Autodesk\Inventor 2024\Bin\Inventor.exe",
                Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Autodesk\Inventor 2026\Bin\Inventor.exe"),
            };
            
            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                    return path;
            }
            
            // Chercher dans Program Files
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string autodeskPath = Path.Combine(programFiles, "Autodesk");
            
            if (Directory.Exists(autodeskPath))
            {
                foreach (var dir in Directory.GetDirectories(autodeskPath, "Inventor*"))
                {
                    string inventorExe = Path.Combine(dir, "Bin", "Inventor.exe");
                    if (File.Exists(inventorExe))
                        return inventorExe;
                }
            }
            
            return null;
        }
        
        /// <summary>
        /// Trouve l'exécutable Vault Client sur le système
        /// </summary>
        private string FindVaultClientExecutable()
        {
            // Chemins standards pour Vault Client 2026
            string[] possiblePaths = new[]
            {
                @"C:\Program Files\Autodesk\Vault Professional 2026\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Client 2026\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Professional 2025\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Autodesk Vault Professional 2026\Explorer\Connectivity.VaultPro.exe",
            };
            
            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                    return path;
            }
            
            // Chercher dans Program Files
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string autodeskPath = Path.Combine(programFiles, "Autodesk");
            
            if (Directory.Exists(autodeskPath))
            {
                foreach (var dir in Directory.GetDirectories(autodeskPath, "Vault*"))
                {
                    string explorerDir = Path.Combine(dir, "Explorer");
                    if (Directory.Exists(explorerDir))
                    {
                        string vaultExe = Path.Combine(explorerDir, "Connectivity.VaultPro.exe");
                        if (File.Exists(vaultExe))
                            return vaultExe;
                    }
                }
            }
            
            return null;
        }

        /// <summary>
        /// Vérifie si le Vault Client est lancé, sinon le lance
        /// </summary>
        private async Task<bool> CheckAndLaunchVaultClientAsync()
        {
            return await Task.Run(async () =>
            {
                try
                {
                    // Noms de processus pour Vault Client
                    string[] vaultProcessNames = { "Connectivity.VaultPro", "Connectivity.Vault", "Explorer" };
                    bool vaultClientFound = false;
                    string foundProcessName = null;
                    
                    foreach (var processName in vaultProcessNames)
                    {
                        var processes = Process.GetProcessesByName(processName);
                        if (processes.Length > 0)
                        {
                            foreach (var proc in processes)
                            {
                                try
                                {
                                    string path = proc.MainModule?.FileName;
                                    if (path != null && path.ToLower().Contains("autodesk"))
                                    {
                                        vaultClientFound = true;
                                        foundProcessName = processName;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                        if (vaultClientFound) break;
                    }
                    
                    // Alternative: chercher fenêtre avec "Vault" dans le titre
                    if (!vaultClientFound)
                    {
                        var allProcesses = Process.GetProcesses();
                        foreach (var proc in allProcesses)
                        {
                            try
                            {
                                if (proc.MainWindowTitle.Contains("Vault") && 
                                    !proc.MainWindowTitle.Contains("VaultAutomation"))
                                {
                                    string path = proc.MainModule?.FileName;
                                    if (path != null && path.ToLower().Contains("autodesk"))
                                    {
                                        vaultClientFound = true;
                                        foundProcessName = proc.ProcessName;
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    
                    if (vaultClientFound)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            AddLog("   [>] Vault Client detecte (processus: " + foundProcessName + ")", "INFO");
                            AddLog("   [+] Vault Professional Client - Pret pour connexion", "SUCCESS");
                        });
                        return true;
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            AddLog("   [!] Vault Client non detecte", "WARN");
                            AddLog("   [>] Lancement de Vault Client 2026...", "INFO");
                        });
                        
                        // Chercher et lancer Vault Client
                        string vaultPath = FindVaultClientExecutable();
                        if (!string.IsNullOrEmpty(vaultPath))
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = vaultPath,
                                UseShellExecute = true
                            });
                            
                            Dispatcher.Invoke(() =>
                                AddLog("   [~] Vault Client demarre, veuillez patienter...", "INFO"));
                            
                            // Attendre que Vault demarre (max 15 sec)
                            for (int i = 0; i < 15; i++)
                            {
                                await Task.Delay(1000);
                                foreach (var pName in vaultProcessNames)
                                {
                                    var procs = Process.GetProcessesByName(pName);
                                    if (procs.Length > 0)
                                    {
                                        Dispatcher.Invoke(() =>
                                            AddLog("   [+] Vault Client demarre avec succes!", "SUCCESS"));
                                        return true;
                                    }
                                }
                            }
                            
                            Dispatcher.Invoke(() =>
                                AddLog("   [!] Vault Client en cours de demarrage...", "WARN"));
                            return false;
                        }
                        else
                        {
                            Dispatcher.Invoke(() =>
                            {
                                AddLog("   [!] Vault Client 2026 non trouve sur ce poste", "WARN");
                                AddLog("   -> La connexion SDK sera utilisee", "INFO");
                            });
                            return false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AddLog("   [-] Erreur verification Vault Client: " + ex.Message, "ERROR");
                    });
                    return false;
                }
            });
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
                RunVaultName.Text = $" Vault : {_vaultService.VaultName}  /  ";
                RunUserName.Text = $" Utilisateur : {_vaultService.UserName}  /  ";
                RunStatus.Text = " Statut : Connecte";
                ConnectButton.Content = "Deconnecter";
                ConnectButton.Background = greenBtnBrush;  // Vert brillant quand connecté
            }
            else
            {
                VaultIndicator.Fill = _redBrush;
                RunVaultName.Text = " Vault : --  /  ";
                RunUserName.Text = " Utilisateur : --  /  ";
                RunStatus.Text = " Statut : Deconnecte";
                ConnectButton.Content = "Connecter";
                ConnectButton.Background = redBtnBrush;    // Rouge clair quand déconnecté
            }

            if (_isInventorConnected && _inventorService.IsConnected)
            {
                InventorIndicator.Fill = _greenBrush;
                string docName = _inventorService.GetActiveDocumentName();
                InventorStatusText.Text = !string.IsNullOrEmpty(docName) ? docName : "Connecte";
                
                // Arrêter le timer de reconnexion si on est connecté
                _inventorReconnectTimer?.Stop();
            }
            else
            {
                InventorIndicator.Fill = _redBrush;
                InventorStatusText.Text = "Deconnecte";
            }
        }

        /// <summary>
        /// Timer de reconnexion automatique à Inventor
        /// Se déclenche toutes les 3 secondes si Inventor n'est pas connecté
        /// Optimise pour eviter le spam de logs
        /// </summary>
        private void InventorReconnectTimer_Tick(object? sender, EventArgs e)
        {
            // Vérifier si Inventor est en cours d'exécution
            var inventorProcesses = Process.GetProcessesByName("Inventor");
            if (inventorProcesses.Length == 0)
            {
                // Inventor pas lancé, ne pas spammer les logs
                return;
            }

            // Tentative de reconnexion (le throttling est gere dans InventorService)
            if (_inventorService.TryConnect())
            {
                _isInventorConnected = true;
                string version = _inventorService.GetInventorVersion() ?? "";
                
                AddLog("[+] Connexion a Inventor etablie!", "SUCCESS");
                if (!string.IsNullOrEmpty(version))
                {
                    AddLog("   [i] Version: " + version, "INFO");
                }
                
                _inventorReconnectTimer?.Stop();
                UpdateConnectionStatus();
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
            AddLog("Ouverture de l'outil Upload Module vers Vault...", "START");
            StatusText.Text = "Upload Module...";
            
            try
            {
                // Ouvrir le module integre (plus besoin de l'exe externe)
                var uploadWindow = new Modules.UploadModule.Views.UploadModuleWindow(_isVaultConnected ? _vaultService : null);
                uploadWindow.Owner = this;
                
                AddLog("Fenetre Upload Module ouverte", "INFO");
                AddLog("Options disponibles:", "INFO");
                AddLog("  - Selection de module depuis C:\\Vault\\Engineering\\Projects", "INFO");
                AddLog("  - Parcourir dossier local", "INFO");
                AddLog("  - Depuis Inventor (document actif)", "INFO");
                
                var result = uploadWindow.ShowDialog();
                
                if (result == true)
                {
                    AddLog("Upload Module termine avec succes!", "SUCCESS");
                    StatusText.Text = "Upload termine";
                }
                else
                {
                    AddLog("Upload Module ferme", "INFO");
                    StatusText.Text = "Upload annule";
                }
            }
            catch (Exception ex)
            {
                AddLog("Erreur ouverture Upload Module: " + ex.Message, "ERROR");
                StatusText.Text = "Erreur";
            }
        }

        private void OpenUploadTemplate_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Ouverture de l'outil Upload Template...", "START");
            StatusText.Text = "Upload Template...";
            
            try
            {
                // Liste des admins autorisés (case-insensitive)
                var admins = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Administrator",
                    "mohammedamine.elgalai"
                };

                // Vérifier si l'utilisateur actuel est admin
                string currentUser = _vaultService?.GetCurrentUsername() ?? "";
                bool isAdmin = admins.Contains(currentUser);

                if (!isAdmin)
                {
                    XnrgyMessageBox.ShowInfo(
                        "Cette fonctionnalite est reservee aux Administrateurs.\n\n" +
                        "Si vous avez une demande d'amelioration ou besoin d'acces,\n" +
                        "veuillez contacter les Admins ou le developpeur de l'application.\n\n" +
                        "Contact: mohammedamine.elgalai@xnrgy.com",
                        "Acces Reserve - Admin Only",
                        this);
                    AddLog("Acces refuse - fonctionnalite reservee aux admins", "WARN");
                    StatusText.Text = "Acces refuse";
                    return;
                }

                // Verifier connexion Vault
                if (!_isVaultConnected || _vaultService == null)
                {
                    XnrgyMessageBox.ShowWarning(
                        "Veuillez d'abord vous connecter au Vault.",
                        "Connexion Vault requise",
                        this);
                    AddLog("Upload Template - Connexion Vault requise", "WARN");
                    return;
                }

                // Ouvrir la fenetre Upload Template avec le service partage
                var uploadTemplateWindow = new UploadTemplateWindow(_vaultService);
                uploadTemplateWindow.Owner = this;
                
                AddLog("Fenetre Upload Template ouverte", "INFO");
                AddLog("Fonctionnalite: Upload de templates vers Vault PROD", "INFO");
                
                var result = uploadTemplateWindow.ShowDialog();
                
                if (result == true)
                {
                    AddLog("Upload Template termine avec succes!", "SUCCESS");
                    StatusText.Text = "Upload termine";
                }
                else
                {
                    AddLog("Upload Template ferme", "INFO");
                    StatusText.Text = "Upload annule";
                }
            }
            catch (Exception ex)
            {
                AddLog("Erreur ouverture Upload Template: " + ex.Message, "ERROR");
                StatusText.Text = "Erreur";
            }
        }

        private void OpenPackAndGo_Click(object sender, RoutedEventArgs e)
        {
            AddLog("Ouverture de l'outil Créer Module...", "START");
            StatusText.Text = "Création de module...";
            
            try
            {
                // [+] Rafraîchir la connexion COM Inventor avant d'ouvrir le module
                if (_isInventorConnected)
                {
                    AddLog("[>] Rafraichissement connexion Inventor...", "DEBUG");
                    _inventorService.ForceReconnect();
                    AddLog("[+] Connexion Inventor rafraichie", "DEBUG");
                }
                
                // Passer les services Vault et Inventor pour héritage du statut de connexion
                var createModuleWindow = new CreateModuleWindow(
                    _isVaultConnected ? _vaultService : null,
                    _isInventorConnected ? _inventorService : null);
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
            AddLog("Ouverture de Smart Tools...", "START");
            StatusText.Text = "Smart Tools...";
            
            try
            {
                // [+] Rafraîchir la connexion COM Inventor avant d'ouvrir le module
                if (_isInventorConnected)
                {
                    AddLog("[>] Rafraichissement connexion Inventor...", "DEBUG");
                    _inventorService.ForceReconnect();
                    AddLog("[+] Connexion Inventor rafraichie", "DEBUG");
                }
                
                // Passer le service Vault et le callback de log pour affichage du statut de connexion
                var smartToolsWindow = new SmartToolsWindow(_isVaultConnected ? _vaultService : null, AddLog);
                smartToolsWindow.Owner = this;
                smartToolsWindow.Show();
                
                AddLog("Smart Tools ouvert avec succes", "SUCCESS");
            }
            catch (Exception ex)
            {
                AddLog($"Erreur lors de l'ouverture de Smart Tools: {ex.Message}", "ERROR");
                MessageBox.Show($"Erreur lors de l'ouverture de Smart Tools:\n{ex.Message}", 
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
                    // Ouvrir dans une fenêtre intégrée avec le service Vault
                    var checklistWindow = new ChecklistHVACWindow(ChecklistHVACPath, _vaultService);
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
