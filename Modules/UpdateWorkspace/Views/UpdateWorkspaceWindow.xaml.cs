using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using VDF = Autodesk.DataManagement.Client.Framework;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Modules.UpdateWorkspace.Views
{
    /// <summary>
    /// Fenetre de mise a jour du workspace local depuis Vault.
    /// Affiche la progression des telechargements, copies et installations.
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
        public UpdateWorkspaceWindow(VDF.Vault.Currency.Connections.Connection? connection = null)
        {
            InitializeComponent();
            DataContext = this;

            _connection = connection;
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

        private void UpdateWorkspaceWindow_Loaded(object sender, RoutedEventArgs e)
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

            // Message initial
            AddLog("[i] Pret pour la mise a jour du workspace");
            AddLog($"[i] Connexion Vault: {(_connection != null ? "Active" : "Non connecte")}");

            // Activer/desactiver les boutons selon l'etat de connexion
            BtnContinue.IsEnabled = _connection != null;
            BtnContinue.Content = _connection != null ? "Demarrer" : "Connexion requise";
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
                    AddLog($"    - {result.InstalledTools} outils installes");
                    AddLog($"    - Duree: {result.Duration.TotalSeconds:F1}s");

                    TxtStatus.Text = "Mise a jour terminee avec succes!";
                    UpdateProgress(100, "Termine!");

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
                AddLog("[!] Mise a jour annulee");
                TxtStatus.Text = "Mise a jour annulee";

                BtnContinue.Content = "Continuer sans mise a jour";
                BtnContinue.IsEnabled = true;
                BtnCancel.IsEnabled = true;
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

            // Mettre a jour l'icone selon le statut - COULEURS VIVES
            switch (status)
            {
                case UpdateWorkspaceService.StepStatus.Pending:
                    iconBlock.Text = "⏳";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF)); // Blanc
                    break;

                case UpdateWorkspaceService.StepStatus.InProgress:
                    iconBlock.Text = "🔄";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0x40, 0xC4, 0xFF)); // Bleu clair vif
                    break;

                case UpdateWorkspaceService.StepStatus.Completed:
                    iconBlock.Text = "✅";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0xE6, 0x76)); // Vert vif
                    break;

                case UpdateWorkspaceService.StepStatus.Failed:
                    iconBlock.Text = "❌";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0x52, 0x52)); // Rouge vif
                    break;

                case UpdateWorkspaceService.StepStatus.Warning:
                    iconBlock.Text = "⚠️";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0xAB, 0x40)); // Orange vif
                    break;

                case UpdateWorkspaceService.StepStatus.Skipped:
                    iconBlock.Text = "⏸️";
                    textBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xB0, 0xB0, 0xB0)); // Gris clair
                    break;
            }

            // Mettre a jour le detail si fourni - avec couleur correspondante
            if (detailBlock != null && !string.IsNullOrEmpty(message))
            {
                detailBlock.Text = $"({message})";
                // Couleur du detail selon le statut
                detailBlock.Foreground = status switch
                {
                    UpdateWorkspaceService.StepStatus.Completed => new SolidColorBrush(Color.FromRgb(0x00, 0xE6, 0x76)),
                    UpdateWorkspaceService.StepStatus.InProgress => new SolidColorBrush(Color.FromRgb(0x40, 0xC4, 0xFF)),
                    UpdateWorkspaceService.StepStatus.Failed => new SolidColorBrush(Color.FromRgb(0xFF, 0x52, 0x52)),
                    UpdateWorkspaceService.StepStatus.Warning => new SolidColorBrush(Color.FromRgb(0xFF, 0xAB, 0x40)),
                    _ => new SolidColorBrush(Color.FromRgb(0xB0, 0xB0, 0xB0))
                };
            }
        }

        private void AddLog(string message, UpdateWorkspaceService.LogLevel level = UpdateWorkspaceService.LogLevel.INFO)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string formattedMessage = $"[{timestamp}]    {message}";
            
            // Couleur selon le niveau de log
            var color = level switch
            {
                UpdateWorkspaceService.LogLevel.SUCCESS => new SolidColorBrush(Color.FromRgb(0x00, 0xE6, 0x76)),  // Vert vif
                UpdateWorkspaceService.LogLevel.ERROR => new SolidColorBrush(Color.FromRgb(0xFF, 0x52, 0x52)),    // Rouge vif
                UpdateWorkspaceService.LogLevel.WARNING => new SolidColorBrush(Color.FromRgb(0xFF, 0xAB, 0x40)),  // Orange vif
                UpdateWorkspaceService.LogLevel.INFO => new SolidColorBrush(Color.FromRgb(0x40, 0xC4, 0xFF)),     // Bleu clair vif
                _ => new SolidColorBrush(Colors.White)
            };
            
            LogEntries.Add(new LogEntry { Message = formattedMessage, Color = color });

            // Faire defiler vers le bas
            if (LogListBox.Items.Count > 0)
            {
                LogListBox.ScrollIntoView(LogListBox.Items[LogListBox.Items.Count - 1]);
            }
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
    }
}
