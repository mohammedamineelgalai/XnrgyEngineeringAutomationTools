using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using VDF = Autodesk.DataManagement.Client.Framework;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Shared.Views;

namespace XnrgyEngineeringAutomationTools.Modules.UpdateWorkspace.Views
{
    /// <summary>
    /// Fenetre de mise a jour du workspace local depuis Vault.
    /// Affiche la progression des telechargements, copies et installations.
    /// Supporte le mode automatique (au demarrage) et manuel (bouton Update Workspace).
    /// </summary>
    public partial class UpdateWorkspaceWindow : Window
    {
        #region Fields

        private readonly VDF.Vault.Currency.Connections.Connection? _connection;
        private readonly UpdateWorkspaceService _updateService;
        private CancellationTokenSource? _cancellationTokenSource;
        private readonly DispatcherTimer _elapsedTimer;
        private readonly Stopwatch _stopwatch;
        private bool _isRunning = false;
        private bool _wasSkipped = false;
        private bool _closeAllowed = false;
        private readonly bool _autoStart;
        private readonly bool _autoCloseOnSuccess;

        // Elements UI pour les etapes (stockes par numero)
        private readonly TextBlock[] _stepIcons;
        private readonly TextBlock[] _stepTexts;
        private readonly TextBlock[] _stepDetails;

        // Collection pour le journal avec couleurs
        public ObservableCollection<LogEntry> LogEntries { get; } = new ObservableCollection<LogEntry>();

        #endregion

        #region Log Entry Class

        /// <summary>
        /// Classe pour les entrees de log avec couleur
        /// </summary>
        public class LogEntry
        {
            public string Message { get; set; } = "";
            public SolidColorBrush Color { get; set; } = new SolidColorBrush(Colors.White);
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructeur de la fenetre UpdateWorkspace
        /// </summary>
        /// <param name="connection">Connexion Vault active (peut etre null pour mode skip)</param>
        /// <param name="autoStart">Si true, demarre automatiquement l'update sans attendre un clic</param>
        /// <param name="autoCloseOnSuccess">Si true, ferme automatiquement la fenetre apres succes</param>
        public UpdateWorkspaceWindow(
            VDF.Vault.Currency.Connections.Connection? connection = null,
            bool autoStart = false,
            bool autoCloseOnSuccess = false)
        {
            InitializeComponent();
            DataContext = this;

            _connection = connection;
            _autoStart = autoStart;
            _autoCloseOnSuccess = autoCloseOnSuccess;
            _updateService = new UpdateWorkspaceService();
            _stopwatch = new Stopwatch();

            // Timer pour le temps ecoule
            _elapsedTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _elapsedTimer.Tick += ElapsedTimer_Tick;

            // Initialiser les tableaux d'elements UI
            _stepIcons = new TextBlock[9];
            _stepTexts = new TextBlock[9];
            _stepDetails = new TextBlock[9];

            Loaded += UpdateWorkspaceWindow_Loaded;
        }

        #endregion

        #region Initialization

        private async void UpdateWorkspaceWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Remplir les tableaux avec les elements XAML
            _stepIcons[0] = IconStep1;
            _stepIcons[1] = IconStep2;
            _stepIcons[2] = IconStep3;
            _stepIcons[3] = IconStep4;
            _stepIcons[4] = IconStep5;
            _stepIcons[5] = IconStep6;
            _stepIcons[6] = IconStep7;
            _stepIcons[7] = IconStep8;
            _stepIcons[8] = IconStep9;

            _stepTexts[0] = TxtStep1;
            _stepTexts[1] = TxtStep2;
            _stepTexts[2] = TxtStep3;
            _stepTexts[3] = TxtStep4;
            _stepTexts[4] = TxtStep5;
            _stepTexts[5] = TxtStep6;
            _stepTexts[6] = TxtStep7;
            _stepTexts[7] = TxtStep8;
            _stepTexts[8] = TxtStep9;

            _stepDetails[0] = DetailStep1;
            _stepDetails[1] = DetailStep2;
            _stepDetails[2] = DetailStep3;
            _stepDetails[3] = DetailStep4;
            _stepDetails[4] = DetailStep5;
            _stepDetails[5] = DetailStep6;
            _stepDetails[6] = DetailStep7;
            _stepDetails[7] = DetailStep8;
            _stepDetails[8] = DetailStep9;

            // Initialiser toutes les etapes a l'etat "en attente"
            for (int i = 0; i < 9; i++)
            {
                UpdateStepUI(i + 1, UpdateWorkspaceService.StepStatus.Pending, null);
            }

            // S'abonner aux evenements du service
            _updateService.ProgressChanged += OnProgressChanged;
            _updateService.LogMessage += OnLogMessage;
            _updateService.StepChanged += OnStepChanged;

            // Mettre a jour le statut Vault dans le header
            UpdateVaultStatusUI();

            // Message initial
            if (_autoStart)
            {
                AddLog("[i] Mode automatique - demarrage immediat de la mise a jour");
            }
            else
            {
                AddLog("[i] Pret pour la mise a jour du workspace");
            }
            AddLog($"[i] Connexion Vault: {(_connection != null ? "Active" : "Non connecte")}");

            // Activer/desactiver les boutons selon l'etat de connexion
            if (_autoStart)
            {
                // Mode auto: demarrer immediatement
                BtnContinue.IsEnabled = false;
                BtnContinue.Content = "En cours...";
                
                // Delai court pour laisser la fenetre s'afficher
                await Task.Delay(500);
                StartUpdate();
            }
            else
            {
                // Mode manuel: attendre clic utilisateur
                BtnContinue.IsEnabled = _connection != null;
                BtnContinue.Content = _connection != null ? "Demarrer" : "Connexion requise";
            }
        }

        /// <summary>
        /// Empeche la fermeture de la fenetre pendant le processus de mise a jour
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Permettre la fermeture si autorisee (succes) ou si pas en cours
            if (_isRunning && !_closeAllowed)
            {
                // Empecher la fermeture pendant le processus
                e.Cancel = true;
                XnrgyMessageBox.ShowWarning(
                    "La mise a jour est en cours.\nVeuillez attendre la fin du processus.",
                    "Fermeture impossible",
                    this);
            }
        }

        /// <summary>
        /// Met a jour l'affichage du statut Vault dans le header
        /// </summary>
        private void UpdateVaultStatusUI()
        {
            if (_connection != null)
            {
                // Connexion active - afficher les infos reelles
                VaultStatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#10B981"));
                RunVaultName.Text = $" Vault: {_connection.Vault}";
                RunUserName.Text = $" {_connection.UserName}";
                RunStatus.Text = " Connecte";
            }
            else
            {
                // Pas de connexion
                VaultStatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E81123"));
                RunVaultName.Text = " Vault: --";
                RunUserName.Text = " --";
                RunStatus.Text = " Deconnecte";
            }
        }

        #endregion

        #region Event Handlers - UI

        private void BtnContinue_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRunning)
            {
                StartUpdate();
            }
            else
            {
                // Deja termine - fermer la fenetre
                _closeAllowed = true;
                DialogResult = true;
                Close();
            }
        }

        private void BtnSkip_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning)
            {
                // Annuler l'operation en cours
                _cancellationTokenSource?.Cancel();
            }
            
            _wasSkipped = true;
            _closeAllowed = true;
            AddLog("[!] Mise a jour ignoree par l'utilisateur");
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning)
            {
                _cancellationTokenSource?.Cancel();
                AddLog("[!] Annulation en cours...");
                BtnCancel.IsEnabled = false;
            }
            else
            {
                _closeAllowed = true;
                DialogResult = false;
                Close();
            }
        }

        private void ElapsedTimer_Tick(object? sender, EventArgs e)
        {
            if (_stopwatch.IsRunning)
            {
                var elapsed = _stopwatch.Elapsed;
                TxtElapsedTime.Text = $"{elapsed.Minutes:D2}:{elapsed.Seconds:D2}";
            }
        }

        #endregion

        #region Update Execution

        private async void StartUpdate()
        {
            if (_connection == null)
            {
                AddLog("[-] Connexion Vault requise pour la mise a jour");
                return;
            }

            _isRunning = true;
            _cancellationTokenSource = new CancellationTokenSource();

            // Mettre a jour l'UI
            BtnContinue.IsEnabled = false;
            BtnCancel.IsEnabled = false;  // Desactiver Annuler des le debut du processus
            BtnSkip.Content = "Interrompre";

            // Demarrer le chrono
            _stopwatch.Start();
            _elapsedTimer.Start();

            AddLog("[>] Demarrage de la mise a jour...");
            TxtStatus.Text = "Mise a jour en cours...";

            try
            {
                var result = await _updateService.ExecuteFullUpdateAsync(
                    _connection, 
                    _cancellationTokenSource.Token);

                _stopwatch.Stop();
                _elapsedTimer.Stop();

                if (result.Success)
                {
                    AddLog($"[+] Mise a jour terminee avec succes!");
                    AddLog($"    - {result.DownloadedFiles} fichiers telecharges");
                    AddLog($"    - {result.CopiedPluginFiles} fichiers plugins copies");
                    AddLog($"    - {result.InstalledTools} applications installees");
                    AddLog($"    - Duree: {result.Duration.TotalSeconds:F1}s");

                    TxtStatus.Text = "Mise a jour terminee avec succes!";
                    UpdateProgress(100, "Termine!");

                    // Lancer Vault Client et Inventor apres succes
                    AddLog("[>] Lancement des applications...");
                    await LaunchApplicationsAfterUpdateAsync();

                    // Mode automatique avec fermeture auto: fermer apres 2 secondes
                    if (_autoCloseOnSuccess)
                    {
                        AddLog("[i] Fermeture automatique dans 2 secondes...");
                        await Task.Delay(2000);
                        _closeAllowed = true;  // Autoriser la fermeture apres succes
                        DialogResult = true;
                        Close();
                        return;
                    }

                    BtnContinue.Content = "Continuer";
                    BtnContinue.IsEnabled = true;
                    BtnSkip.Visibility = Visibility.Collapsed;
                    BtnCancel.Visibility = Visibility.Collapsed;
                }
                else
                {
                    AddLog($"[-] Mise a jour terminee avec erreurs: {result.ErrorMessage}");
                    TxtStatus.Text = $"Erreur: {result.ErrorMessage}";

                    BtnContinue.Content = "Continuer malgre erreurs";
                    BtnContinue.IsEnabled = true;
                }
            }
            catch (OperationCanceledException)
            {
                _stopwatch.Stop();
                _elapsedTimer.Stop();
                AddLog("[!] Mise a jour interrompue");
                TxtStatus.Text = "Mise a jour interrompue";

                BtnContinue.Content = "Continuer sans mise a jour";
                BtnContinue.IsEnabled = true;
                // BtnCancel reste desactive - l'utilisateur a deja interrompu le processus
            }
            catch (Exception ex)
            {
                _stopwatch.Stop();
                _elapsedTimer.Stop();
                AddLog($"[-] Erreur critique: {ex.Message}");
                TxtStatus.Text = $"Erreur: {ex.Message}";

                BtnContinue.Content = "Continuer malgre erreurs";
                BtnContinue.IsEnabled = true;
            }
            finally
            {
                _isRunning = false;
            }
        }

        #endregion

        #region Service Event Handlers

        private void OnProgressChanged(object? sender, UpdateWorkspaceService.UpdateProgressEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                UpdateProgress(e.Percent, e.Status);
                if (!string.IsNullOrEmpty(e.CurrentFile))
                {
                    TxtCurrentFile.Text = e.CurrentFile;
                }
            });
        }

        private void OnLogMessage(object? sender, UpdateWorkspaceService.UpdateLogEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                AddLog(e.Message, e.Level);
            });
        }

        private void OnStepChanged(object? sender, UpdateWorkspaceService.UpdateStepEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                UpdateStepUI(e.StepNumber, e.Status, e.Message);
            });
        }

        #endregion

        #region UI Helpers

        private void UpdateProgress(int percent, string status)
        {
            // Mettre a jour la barre de progression
            double barWidth = (percent / 100.0) * (ActualWidth - 70); // Ajuster pour les marges
            if (barWidth < 0) barWidth = 0;
            ProgressBarFill.Width = barWidth;

            // Mettre a jour le pourcentage
            TxtProgressPercent.Text = $"{percent}%";

            // Mettre a jour le temps estime restant
            if (_stopwatch.IsRunning && percent > 0 && percent < 100)
            {
                var elapsed = _stopwatch.Elapsed;
                double estimatedTotalSeconds = elapsed.TotalSeconds * 100 / percent;
                double remainingSeconds = estimatedTotalSeconds - elapsed.TotalSeconds;
                if (remainingSeconds > 0)
                {
                    var remaining = TimeSpan.FromSeconds(remainingSeconds);
                    TxtEstimatedTime.Text = $"{(int)remaining.TotalMinutes:D2}:{remaining.Seconds:D2}";
                }
                else
                {
                    TxtEstimatedTime.Text = "00:00";
                }
            }
            else if (percent >= 100)
            {
                TxtEstimatedTime.Text = "00:00";
            }

            // Mettre a jour le statut
            if (!string.IsNullOrEmpty(status))
            {
                TxtStatus.Text = status;
            }
        }

        private void UpdateStepUI(int stepNumber, UpdateWorkspaceService.StepStatus status, string? message)
        {
            if (stepNumber < 1 || stepNumber > 9) return;

            int index = stepNumber - 1;
            var iconBlock = _stepIcons[index];
            var textBlock = _stepTexts[index];
            var detailBlock = _stepDetails[index];

            if (iconBlock == null || textBlock == null) return;

            // Mettre a jour l'icone selon le statut - COULEURS XNRGY STANDARD
            switch (status)
            {
                case UpdateWorkspaceService.StepStatus.Pending:
                    iconBlock.Text = "⏳";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF)); // Blanc
                    break;

                case UpdateWorkspaceService.StepStatus.InProgress:
                    iconBlock.Text = "🔄";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)); // XnrgyBlue #0078D4
                    break;

                case UpdateWorkspaceService.StepStatus.Completed:
                    iconBlock.Text = "✅";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0xD2, 0x6A)); // XnrgyGreen #00D26A
                    break;

                case UpdateWorkspaceService.StepStatus.Failed:
                    iconBlock.Text = "❌";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xE8, 0x11, 0x23)); // XnrgyRed #E81123
                    break;

                case UpdateWorkspaceService.StepStatus.Warning:
                    iconBlock.Text = "⚠️";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0x8C, 0x00)); // XnrgyOrange #FF8C00
                    break;

                case UpdateWorkspaceService.StepStatus.Skipped:
                    iconBlock.Text = "⏸️";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88)); // TextMuted #888888
                    break;
            }

            // Mettre a jour le detail si fourni - avec couleur correspondante
            if (detailBlock != null && !string.IsNullOrEmpty(message))
            {
                detailBlock.Text = $"({message})";
                // Couleur du detail selon le statut - COULEURS XNRGY STANDARD
                detailBlock.Foreground = status switch
                {
                    UpdateWorkspaceService.StepStatus.Completed => new SolidColorBrush(Color.FromRgb(0x00, 0xD2, 0x6A)),   // XnrgyGreen
                    UpdateWorkspaceService.StepStatus.InProgress => new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)),  // XnrgyBlue
                    UpdateWorkspaceService.StepStatus.Failed => new SolidColorBrush(Color.FromRgb(0xE8, 0x11, 0x23)),      // XnrgyRed
                    UpdateWorkspaceService.StepStatus.Warning => new SolidColorBrush(Color.FromRgb(0xFF, 0x8C, 0x00)),     // XnrgyOrange
                    _ => new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88))                                               // TextMuted
                };
            }
        }

        private void AddLog(string message, UpdateWorkspaceService.LogLevel level = UpdateWorkspaceService.LogLevel.INFO)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string formattedMessage = $"[{timestamp}]    {message}";
            
            // Detection automatique du niveau basee sur le prefixe du message
            // Utilise JournalColorService pour uniformite avec les autres formulaires
            var trimmedMsg = message.TrimStart();
            
            SolidColorBrush color;
            if (trimmedMsg.StartsWith("[+]"))
                color = Services.JournalColorService.SuccessBrush;  // Vert #00FF7F
            else if (trimmedMsg.StartsWith("[-]"))
                color = Services.JournalColorService.ErrorBrush;    // Rouge #FF4444
            else if (trimmedMsg.StartsWith("[!]"))
                color = Services.JournalColorService.WarningBrush;  // Jaune #FFD700
            else
                color = Services.JournalColorService.InfoBrush;     // Blanc #FFFFFF (defaut)
            
            LogEntries.Add(new LogEntry { Message = formattedMessage, Color = color });

            // Faire defiler vers le bas
            if (LogListBox.Items.Count > 0)
            {
                LogListBox.ScrollIntoView(LogListBox.Items[LogListBox.Items.Count - 1]);
            }
        }

        /// <summary>
        /// Efface le journal des operations
        /// </summary>
        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            LogEntries.Clear();
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Indique si la mise a jour a ete ignoree
        /// </summary>
        public bool WasSkipped => _wasSkipped;

        #endregion

        #region Cleanup

        protected override void OnClosed(EventArgs e)
        {
            _elapsedTimer.Stop();
            _cancellationTokenSource?.Cancel();
            _cancellationTokenSource?.Dispose();

            // Se desabonner des evenements
            _updateService.ProgressChanged -= OnProgressChanged;
            _updateService.LogMessage -= OnLogMessage;
            _updateService.StepChanged -= OnStepChanged;

            base.OnClosed(e);
        }

        #endregion

        #region Application Launch

        /// <summary>
        /// Lance Vault Client et Inventor apres la mise a jour reussie
        /// </summary>
        private async Task LaunchApplicationsAfterUpdateAsync()
        {
            await Task.Run(async () =>
            {
                try
                {
                    // === LANCER VAULT CLIENT ===
                    string[] vaultProcessNames = { "Connectivity.VaultPro", "Connectivity.Vault" };
                    bool vaultRunning = false;
                    
                    foreach (var pName in vaultProcessNames)
                    {
                        if (Process.GetProcessesByName(pName).Length > 0)
                        {
                            vaultRunning = true;
                            break;
                        }
                    }
                    
                    if (!vaultRunning)
                    {
                        string vaultPath = FindVaultClientExecutable();
                        if (!string.IsNullOrEmpty(vaultPath))
                        {
                            Dispatcher.Invoke(() => AddLog("   [>] Lancement de Vault Client..."));
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = vaultPath,
                                UseShellExecute = true
                            });
                            Dispatcher.Invoke(() => AddLog("   [+] Vault Client demarre"));
                        }
                    }
                    else
                    {
                        Dispatcher.Invoke(() => AddLog("   [i] Vault Client deja en cours"));
                    }

                    // Petit delai entre les lancements
                    await Task.Delay(1000);

                    // === LANCER INVENTOR ===
                    var inventorProcesses = Process.GetProcessesByName("Inventor");
                    if (inventorProcesses.Length == 0)
                    {
                        string inventorPath = FindInventorExecutable();
                        if (!string.IsNullOrEmpty(inventorPath))
                        {
                            Dispatcher.Invoke(() => AddLog("   [>] Lancement d'Inventor Professional 2026..."));
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = inventorPath,
                                UseShellExecute = true
                            });
                            Dispatcher.Invoke(() => AddLog("   [+] Inventor demarre"));
                        }
                        else
                        {
                            Dispatcher.Invoke(() => AddLog("   [!] Inventor non trouve", UpdateWorkspaceService.LogLevel.WARNING));
                        }
                    }
                    else
                    {
                        Dispatcher.Invoke(() => AddLog("   [i] Inventor deja en cours"));
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => AddLog($"   [!] Erreur lancement applications: {ex.Message}", UpdateWorkspaceService.LogLevel.WARNING));
                }
            });
        }

        /// <summary>
        /// Trouve l'executable Inventor 2026
        /// </summary>
        private static string? FindInventorExecutable()
        {
            string[] possiblePaths = new[]
            {
                @"C:\Program Files\Autodesk\Inventor 2026\Bin\Inventor.exe",
                @"C:\Program Files\Autodesk\Inventor 2025\Bin\Inventor.exe",
                @"C:\Program Files\Autodesk\Inventor 2024\Bin\Inventor.exe"
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                    return path;
            }
            return null;
        }

        /// <summary>
        /// Trouve l'executable Vault Client
        /// </summary>
        private static string? FindVaultClientExecutable()
        {
            string[] possiblePaths = new[]
            {
                @"C:\Program Files\Autodesk\Vault Client 2026\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Client 2025\Explorer\Connectivity.VaultPro.exe",
                @"C:\Program Files\Autodesk\Vault Client 2024\Explorer\Connectivity.VaultPro.exe"
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                    return path;
            }
            return null;
        }

        #endregion
    }
}
