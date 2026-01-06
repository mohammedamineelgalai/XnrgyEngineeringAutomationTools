using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Animation;
using XnrgyEngineeringAutomationTools.Shared.Services;

namespace XnrgyEngineeringAutomationTools.Shared.Views
{
    /// <summary>
    /// Fenetre de mise a jour du workspace - Telecharge et synchronise les standards XNRGY depuis Vault
    /// </summary>
    public partial class UpdateWorkspaceWindow : Window
    {
        #region Fields

        private readonly UpdateWorkspaceService _updateService;
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly Stopwatch _stopwatch = new Stopwatch();
        private System.Windows.Threading.DispatcherTimer? _timer;
        private bool _updateCompleted = false;
        private bool _updateSkipped = false;
        private int _totalFiles = 0;
        private long _totalBytes = 0;

        // Mapping des items de checklist
        private readonly Dictionary<string, (TextBlock Emoji, TextBlock Status, Border Item)> _checklistItems = new();

        #endregion

        #region Constructor

        public UpdateWorkspaceWindow(VaultSDKService vaultService)
        {
            InitializeComponent();
            _updateService = new UpdateWorkspaceService(vaultService);
            
            // S'abonner aux evenements du service
            _updateService.ProgressChanged += OnProgressChanged;
            _updateService.LogMessage += OnLogMessage;
            _updateService.ChecklistItemChanged += OnChecklistItemChanged;
            _updateService.StatsUpdated += OnStatsUpdated;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Indique si la mise a jour a ete completee avec succes
        /// </summary>
        public bool UpdateSuccessful => _updateCompleted && !_updateSkipped;

        /// <summary>
        /// Indique si l'utilisateur a passe la mise a jour
        /// </summary>
        public bool UpdateSkipped => _updateSkipped;

        #endregion

        #region Window Events

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Initialiser le mapping des items de checklist
                InitializeChecklistMapping();
                
                // Initialiser le timer pour le temps ecoule
                InitializeTimer();
                
                // Log de demarrage
                AddLog("[i] Demarrage de la mise a jour workspace...", LogLevel.INFO);
                AddLog($"[i] Connexion Vault etablie", LogLevel.INFO);
                
                // Demarrer la mise a jour automatiquement
                await StartUpdateAsync();
            }
            catch (Exception ex)
            {
                AddLog($"[-] Erreur initialisation: {ex.Message}", LogLevel.ERROR);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Annuler les operations en cours si la fenetre est fermee
            if (!_updateCompleted && !_updateSkipped)
            {
                var result = MessageBox.Show(
                    "La mise a jour est en cours. Voulez-vous vraiment annuler?\n\nNote: Passer la mise a jour peut causer des problemes de compatibilite.",
                    "Confirmer l'annulation",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                _cts.Cancel();
                _updateSkipped = true;
            }

            _timer?.Stop();
        }

        #endregion

        #region Initialization

        private void InitializeChecklistMapping()
        {
            _checklistItems["InventorStandards"] = (EmojiInventorStandards, StatusInventorStandards, ItemInventorStandards);
            _checklistItems["Cabinet"] = (EmojiCabinet, StatusCabinet, ItemCabinet);
            _checklistItems["XnrgyM99"] = (EmojiXnrgyM99, StatusXnrgyM99, ItemXnrgyM99);
            _checklistItems["XnrgyModule"] = (EmojiXnrgyModule, StatusXnrgyModule, ItemXnrgyModule);
            _checklistItems["SiblAddins"] = (EmojiSiblAddins, StatusSiblAddins, ItemSiblAddins);
            _checklistItems["XnrgyAddins"] = (EmojiXnrgyAddins, StatusXnrgyAddins, ItemXnrgyAddins);
            _checklistItems["DxfVerifier"] = (EmojiDxfVerifier, StatusDxfVerifier, ItemDxfVerifier);
            _checklistItems["SmartTools"] = (EmojiSmartTools, StatusSmartTools, ItemSmartTools);
            _checklistItems["BatchPrint"] = (EmojiBatchPrint, StatusBatchPrint, ItemBatchPrint);
        }

        private void InitializeTimer()
        {
            _timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _timer.Tick += (s, e) =>
            {
                if (_stopwatch.IsRunning)
                {
                    TxtTimeElapsed.Text = FormatTimeSpan(_stopwatch.Elapsed);
                }
            };
        }

        #endregion

        #region Update Process

        private async Task StartUpdateAsync()
        {
            try
            {
                _stopwatch.Start();
                _timer?.Start();
                
                BtnSkip.IsEnabled = true;
                BtnContinue.IsEnabled = false;
                
                UpdateStatus("Demarrage de la synchronisation...", 0);
                
                // Executer la mise a jour complete
                var result = await _updateService.ExecuteFullUpdateAsync(_cts.Token);
                
                _stopwatch.Stop();
                _timer?.Stop();
                
                if (result.Success)
                {
                    _updateCompleted = true;
                    UpdateStatus($"Mise a jour terminee! {result.FilesDownloaded} fichiers synchronises", 100);
                    AddLog($"[+] MISE A JOUR TERMINEE AVEC SUCCES!", LogLevel.SUCCESS);
                    AddLog($"    Fichiers telecharges: {result.FilesDownloaded}", LogLevel.INFO);
                    AddLog($"    Fichiers copies: {result.FilesCopied}", LogLevel.INFO);
                    AddLog($"    Installations: {result.InstallationsCompleted}", LogLevel.INFO);
                    AddLog($"    Temps total: {FormatTimeSpan(_stopwatch.Elapsed)}", LogLevel.INFO);
                    
                    // Changer le bouton Continuer
                    BtnContinue.IsEnabled = true;
                    BtnContinue.Content = "✅ CONTINUER";
                    BtnSkip.IsEnabled = false;
                    
                    // Effet visuel de succes sur la barre
                    SetProgressBarSuccess();
                }
                else
                {
                    UpdateStatus($"Mise a jour terminee avec {result.Errors.Count} erreur(s)", 100);
                    AddLog($"[!] MISE A JOUR TERMINEE AVEC ERREURS", LogLevel.WARNING);
                    foreach (var error in result.Errors)
                    {
                        AddLog($"    [-] {error}", LogLevel.ERROR);
                    }
                    
                    BtnContinue.IsEnabled = true;
                    BtnContinue.Content = "⚠️ CONTINUER QUAND MEME";
                    
                    // Effet visuel d'avertissement
                    SetProgressBarWarning();
                }
            }
            catch (OperationCanceledException)
            {
                _stopwatch.Stop();
                _timer?.Stop();
                
                AddLog("[!] Mise a jour annulee par l'utilisateur", LogLevel.WARNING);
                UpdateStatus("Mise a jour annulee", 0);
                
                _updateSkipped = true;
                BtnContinue.IsEnabled = true;
                BtnContinue.Content = "⏭️ CONTINUER SANS MAJ";
            }
            catch (Exception ex)
            {
                _stopwatch.Stop();
                _timer?.Stop();
                
                AddLog($"[-] Erreur critique: {ex.Message}", LogLevel.ERROR);
                UpdateStatus($"Erreur: {ex.Message}", 0);
                
                BtnContinue.IsEnabled = true;
                BtnContinue.Content = "⚠️ CONTINUER QUAND MEME";
                
                SetProgressBarError();
            }
        }

        #endregion

        #region Event Handlers from Service

        private void OnProgressChanged(object? sender, ProgressInfo info)
        {
            Dispatcher.Invoke(() =>
            {
                UpdateProgressBar(info.Percent, info.Status, info.CurrentFile);
            });
        }

        private void OnLogMessage(object? sender, LogMessageEventArgs args)
        {
            Dispatcher.Invoke(() =>
            {
                AddLog(args.Message, args.Level);
            });
        }

        private void OnChecklistItemChanged(object? sender, ChecklistItemEventArgs args)
        {
            Dispatcher.Invoke(() =>
            {
                UpdateChecklistItem(args.ItemKey, args.Status, args.StatusText);
            });
        }

        private void OnStatsUpdated(object? sender, StatsEventArgs args)
        {
            Dispatcher.Invoke(() =>
            {
                _totalFiles = args.TotalFiles;
                _totalBytes = args.TotalBytes;
                TxtFilesCount.Text = $"{_totalFiles} fichiers";
                TxtTotalSize.Text = FormatBytes(_totalBytes);
            });
        }

        #endregion

        #region UI Updates

        private void UpdateStatus(string status, int percent)
        {
            TxtStatus.Text = status;
            TxtProgressPercent.Text = $"{percent}%";
            UpdateProgressBarWidth(percent);
        }

        private void UpdateProgressBar(int percent, string status, string? currentFile = null)
        {
            TxtStatus.Text = status;
            TxtProgressPercent.Text = $"{percent}%";
            
            if (!string.IsNullOrEmpty(currentFile))
            {
                TxtCurrentFile.Text = currentFile;
            }
            
            UpdateProgressBarWidth(percent);
        }

        private void UpdateProgressBarWidth(int percent)
        {
            var container = ProgressBarFill.Parent as FrameworkElement;
            if (container != null)
            {
                double maxWidth = container.ActualWidth > 0 ? container.ActualWidth : 600;
                double targetWidth = (percent / 100.0) * maxWidth;

                var widthAnimation = new DoubleAnimation
                {
                    To = targetWidth,
                    Duration = TimeSpan.FromMilliseconds(300),
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                };

                ProgressBarFill.BeginAnimation(WidthProperty, widthAnimation);
            }
        }

        private void UpdateChecklistItem(string itemKey, ChecklistStatus status, string? statusText = null)
        {
            if (!_checklistItems.TryGetValue(itemKey, out var item))
                return;

            // Mettre a jour l'emoji
            item.Emoji.Text = status switch
            {
                ChecklistStatus.Pending => "⏳",
                ChecklistStatus.InProgress => "🔄",
                ChecklistStatus.Success => "✅",
                ChecklistStatus.Error => "❌",
                ChecklistStatus.Skipped => "⏭️",
                _ => "⏳"
            };

            // Mettre a jour le texte de statut
            if (!string.IsNullOrEmpty(statusText))
            {
                item.Status.Text = statusText;
                item.Status.Foreground = status switch
                {
                    ChecklistStatus.Success => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50")),
                    ChecklistStatus.Error => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336")),
                    ChecklistStatus.InProgress => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2196F3")),
                    _ => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#808080"))
                };
            }

            // Mettre a jour le style du border
            item.Item.BorderBrush = status switch
            {
                ChecklistStatus.Success => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2E7D32")),
                ChecklistStatus.Error => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C62828")),
                ChecklistStatus.InProgress => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1565C0")),
                _ => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3A3A5A"))
            };
        }

        private void SetProgressBarSuccess()
        {
            var gradient = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(1, 0)
            };
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00C853"), 0));
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#69F0AE"), 0.5));
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#00C853"), 1));
            ProgressBarFill.Background = gradient;
        }

        private void SetProgressBarWarning()
        {
            var gradient = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(1, 0)
            };
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#FF9800"), 0));
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#FFB74D"), 0.5));
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#FF9800"), 1));
            ProgressBarFill.Background = gradient;
        }

        private void SetProgressBarError()
        {
            var gradient = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(1, 0)
            };
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#F44336"), 0));
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#E57373"), 0.5));
            gradient.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#F44336"), 1));
            ProgressBarFill.Background = gradient;
        }

        #endregion

        #region Journal

        private void AddLog(string message, LogLevel level)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            var logEntry = new TextBlock
            {
                Text = $"[{timestamp}] {message}",
                FontFamily = new FontFamily("Consolas"),
                FontSize = 12,
                Margin = new Thickness(10, 2, 10, 2),
                TextWrapping = TextWrapping.Wrap,
                Foreground = level switch
                {
                    LogLevel.SUCCESS => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50")),
                    LogLevel.ERROR => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336")),
                    LogLevel.WARNING => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF9800")),
                    LogLevel.INFO => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4FC3F7")),
                    LogLevel.DEBUG => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9E9E9E")),
                    _ => new SolidColorBrush(Colors.White)
                }
            };

            LogListBox.Items.Add(logEntry);
            LogListBox.ScrollIntoView(logEntry);
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            LogListBox.Items.Clear();
        }

        #endregion

        #region Button Handlers

        private void BtnSkip_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Etes-vous sur de vouloir passer la mise a jour?\n\n" +
                "ATTENTION: Cela peut causer des problemes de compatibilite avec les standards XNRGY.\n" +
                "Il est fortement recommande de faire la mise a jour.",
                "Passer la mise a jour",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                _cts.Cancel();
                _updateSkipped = true;
                _stopwatch.Stop();
                _timer?.Stop();
                
                AddLog("[!] Mise a jour passee par l'utilisateur", LogLevel.WARNING);
                
                // Marquer tous les items comme passes
                foreach (var key in _checklistItems.Keys)
                {
                    UpdateChecklistItem(key, ChecklistStatus.Skipped, "Passe");
                }
                
                DialogResult = true;
                Close();
            }
        }

        private void BtnContinue_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        #endregion

        #region Helpers

        private string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalHours >= 1)
                return $"{(int)ts.TotalHours}h {ts.Minutes:D2}m {ts.Seconds:D2}s";
            else if (ts.TotalMinutes >= 1)
                return $"{ts.Minutes}m {ts.Seconds:D2}s";
            else
                return $"{ts.Seconds}s";
        }

        private string FormatBytes(long bytes)
        {
            if (bytes >= 1073741824)
                return $"{bytes / 1073741824.0:F2} GB";
            else if (bytes >= 1048576)
                return $"{bytes / 1048576.0:F2} MB";
            else if (bytes >= 1024)
                return $"{bytes / 1024.0:F2} KB";
            else
                return $"{bytes} B";
        }

        #endregion
    }

    #region Enums and Event Args

    public enum LogLevel
    {
        INFO,
        SUCCESS,
        WARNING,
        ERROR,
        DEBUG
    }

    public enum ChecklistStatus
    {
        Pending,
        InProgress,
        Success,
        Error,
        Skipped
    }

    public class ProgressInfo
    {
        public int Percent { get; set; }
        public string Status { get; set; } = "";
        public string? CurrentFile { get; set; }
    }

    public class LogMessageEventArgs : EventArgs
    {
        public string Message { get; set; } = "";
        public LogLevel Level { get; set; }
    }

    public class ChecklistItemEventArgs : EventArgs
    {
        public string ItemKey { get; set; } = "";
        public ChecklistStatus Status { get; set; }
        public string? StatusText { get; set; }
    }

    public class StatsEventArgs : EventArgs
    {
        public int TotalFiles { get; set; }
        public long TotalBytes { get; set; }
    }

    #endregion
}
