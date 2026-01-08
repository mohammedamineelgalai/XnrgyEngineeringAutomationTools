using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Documents;
using XnrgyEngineeringAutomationTools.Modules.SmartTools.Services;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Shared.Views;
using IProgressWindow = XnrgyEngineeringAutomationTools.Modules.SmartTools.Views.IProgressWindow;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenêtre principale pour Smart Tools - Outils d'automatisation Inventor
    /// Migré depuis SmartToolsAmineAddin vers application externe
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public partial class SmartToolsWindow : Window
    {
        private readonly SmartToolsService _smartToolsService;
        private readonly InventorService _inventorService;
        private readonly VaultSdkService? _vaultService;
        private System.Windows.Threading.DispatcherTimer? _inventorStatusTimer;
        
        /// <summary>
        /// Callback pour propager les logs vers le journal principal de MainWindow
        /// </summary>
        public Action<string, string>? MainWindowLogCallback { get; set; }

        /// <summary>
        /// Constructeur par défaut (sans service Vault)
        /// </summary>
        public SmartToolsWindow() : this(null, null)
        {
        }

        /// <summary>
        /// Constructeur avec service Vault pour affichage du statut
        /// </summary>
        /// <param name="vaultService">Service Vault connecté (optionnel)</param>
        /// <param name="mainLogCallback">Callback pour propager les logs vers MainWindow (optionnel)</param>
        public SmartToolsWindow(VaultSdkService? vaultService, Action<string, string>? mainLogCallback = null)
        {
            InitializeComponent();
            _inventorService = new InventorService();
            
            // [+] Forcer la reconnexion COM à chaque ouverture de Smart Tools
            // Évite les problèmes de connexion COM obsolète
            _inventorService.ForceReconnect();
            
            _smartToolsService = new SmartToolsService(_inventorService);
            _vaultService = vaultService;
            MainWindowLogCallback = mainLogCallback;
            
            // Passer le callback pour les popups HTML au service
            _smartToolsService.SetHtmlPopupCallback(ShowHtmlPopup);
            _smartToolsService.SetExportOptionsCallback(ShowExportOptions);
            _smartToolsService.SetProgressWindowCallback(ShowProgressWindow);
            _smartToolsService.SetIPropertiesPopupCallback(ShowIPropertiesPopup);
            _smartToolsService.SetIPropertiesWindowCallback(ShowIPropertiesWindow);
            
            Loaded += SmartToolsWindow_Loaded;
            Closed += SmartToolsWindow_Closed;
            UpdateConnectionStatuses();
        }

        private void SmartToolsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateConnectionStatuses();
            
            // Timer pour mettre à jour le statut Inventor périodiquement
            _inventorStatusTimer = new System.Windows.Threading.DispatcherTimer();
            _inventorStatusTimer.Interval = TimeSpan.FromSeconds(3);
            _inventorStatusTimer.Tick += (s, args) => UpdateInventorStatus();
            _inventorStatusTimer.Start();
        }

        private void SmartToolsWindow_Closed(object sender, EventArgs e)
        {
            _inventorStatusTimer?.Stop();
        }

        /// <summary>
        /// Met à jour les statuts de connexion Vault et Inventor
        /// </summary>
        private void UpdateConnectionStatuses()
        {
            UpdateVaultConnectionStatus();
            UpdateInventorStatus();
        }

        /// <summary>
        /// Met à jour l'indicateur de connexion Vault dans l'en-tête
        /// </summary>
        private void UpdateVaultConnectionStatus()
        {
            Dispatcher.Invoke(() =>
            {
                bool isConnected = _vaultService != null && _vaultService.IsConnected;
                
                if (VaultStatusIndicator != null)
                {
                    VaultStatusIndicator.Fill = new SolidColorBrush(
                        isConnected ? (Color)ColorConverter.ConvertFromString("#107C10") : (Color)ColorConverter.ConvertFromString("#E81123"));
                }
                
                if (RunVaultName != null && RunUserName != null && RunStatus != null)
                {
                    if (isConnected)
                    {
                        RunVaultName.Text = $" Vault : {_vaultService!.VaultName}  /  ";
                        RunUserName.Text = $" Utilisateur : {_vaultService.UserName}  /  ";
                        RunStatus.Text = " Statut : Connecte";
                    }
                    else
                    {
                        RunVaultName.Text = " Vault : --  /  ";
                        RunUserName.Text = " Utilisateur : --  /  ";
                        RunStatus.Text = " Statut : Deconnecte";
                    }
                }
            });
        }

        /// <summary>
        /// Met à jour l'indicateur de connexion Inventor dans l'en-tête
        /// Affiche le nom du document actif comme dans le formulaire principal
        /// </summary>
        private void UpdateInventorStatus()
        {
            Dispatcher.Invoke(() =>
            {
                bool isConnected = _inventorService.IsConnected;
                
                if (InventorIndicator != null)
                {
                    InventorIndicator.Fill = new SolidColorBrush(
                        isConnected ? (Color)ColorConverter.ConvertFromString("#107C10") : (Color)ColorConverter.ConvertFromString("#E81123"));
                }
                
                if (RunInventorStatus != null)
                {
                    if (isConnected)
                    {
                        // Afficher le nom du document actif comme dans MainWindow
                        string docName = _inventorService.GetActiveDocumentName();
                        RunInventorStatus.Text = !string.IsNullOrEmpty(docName) ? $" Inventor : {docName}" : " Inventor : Connecte";
                    }
                    else
                    {
                        RunInventorStatus.Text = " Inventor : Deconnecte";
                    }
                }
            });
        }

        private async void BtnHideBox_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteHideBoxAsync((msg, level) => Log(msg, level));
                // Tous les messages sont déjà loggés dans le service
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'exécution de HideBox: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnToggleRefVisibility_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteToggleRefVisibilityAsync((msg, level) => Log(msg, level));
                // Tous les messages (début, progression, succès, erreurs) sont déjà loggés dans le service
            }
            catch (Exception ex)
            {
                Log($"Erreur lors du basculage: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnToggleSketchVisibility_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteToggleSketchVisibilityAsync((msg, level) => Log(msg, level));
                // Tous les messages sont déjà loggés dans le service
            }
            catch (Exception ex)
            {
                Log($"Erreur lors du basculage: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnConstraintReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteConstraintReportAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la génération du rapport: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnIPropertiesSummary_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteIPropertiesSummaryAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la génération du résumé: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnSafeClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteSafeCloseAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la fermeture: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnSmartSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteSmartSaveAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la sauvegarde: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnExportIAMToIPT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteExportIAMToIPTAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'export: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnExportIDWtoShopPDF_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteExportIDWtoShopPDFAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'export: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnFormCenteringUtility_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteFormCenteringUtilityAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors du centrage: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnILogicFormsCentred_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteILogicFormsCentredAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors du centrage: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnInfoCommand_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteInfoCommandAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'affichage des informations: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnAutoSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteAutoSaveAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'activation de l'auto-sauvegarde: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnCheckSaveStatus_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteCheckSaveStatusAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la vérification: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnCollapseExpandAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteCollapseExpandAllAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'opération: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnInsertSpecificScrewInHoles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteInsertSpecificScrewInHolesAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'insertion: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnIPropertyCustomBatch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteIPropertyCustomBatchAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la modification: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnFixPromotedVaries_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteFixPromotedVariesAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la correction: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void Btn2DIsometricView_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.Execute2DIsometricViewAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'application de la vue: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnOpenSelectedComponent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteOpenSelectedComponentAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de l'ouverture: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnReturnToFrontView_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteReturnToFrontViewAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors du retour à la vue: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnUpdateCurrentSheet_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteUpdateCurrentSheetAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors de la mise à jour: {ex.Message}", LogLevel.ERROR);
            }
        }

        private async void BtnHideBoxTemplateMultiOpening_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _smartToolsService.ExecuteHideBoxTemplateMultiOpeningAsync((msg, level) => Log(msg, level));
            }
            catch (Exception ex)
            {
                Log($"Erreur lors du masquage: {ex.Message}", LogLevel.ERROR);
            }
        }

        // ====================================================================
        // Journal et UI
        // ====================================================================
        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            TxtLog.Document.Blocks.Clear();
        }

        private enum LogLevel
        {
            INFO,
            SUCCESS,
            WARNING,
            ERROR
        }

        private void Log(string message, LogLevel level)
        {
            Log(message, level.ToString());
        }

        private void Log(string message, string level)
        {
            // Convertir string en LogLevel
            LogLevel logLevel = level switch
            {
                "SUCCESS" => LogLevel.SUCCESS,
                "ERROR" => LogLevel.ERROR,
                "WARNING" => LogLevel.WARNING,
                _ => LogLevel.INFO
            };

            // Ecrire aussi dans le fichier log principal de l'application
            var loggerLevel = logLevel switch
            {
                LogLevel.SUCCESS => Logger.LogLevel.INFO,
                LogLevel.ERROR => Logger.LogLevel.ERROR,
                LogLevel.WARNING => Logger.LogLevel.WARNING,
                _ => Logger.LogLevel.INFO
            };
            Logger.Log($"[SmartTools] {message}", loggerLevel);
            
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                
                // Utilise JournalColorService pour les couleurs uniformisees
                var serviceLevel = logLevel switch
                {
                    LogLevel.SUCCESS => JournalColorService.LogLevel.SUCCESS,
                    LogLevel.ERROR => JournalColorService.LogLevel.ERROR,
                    LogLevel.WARNING => JournalColorService.LogLevel.WARNING,
                    _ => JournalColorService.LogLevel.INFO
                };
                
                // Obtenir couleur et prefix depuis le service centralise
                var color = JournalColorService.GetColorForLevel(serviceLevel);
                string prefix = JournalColorService.GetPrefixForLevel(serviceLevel);
                
                // Creer le paragraph avec couleur
                var paragraph = new Paragraph();
                paragraph.Margin = new Thickness(0, 2, 0, 2);
                
                // Timestamp en gris (depuis le service)
                var timestampRun = new Run($"[{timestamp}] ")
                {
                    Foreground = JournalColorService.TimestampBrush
                };
                paragraph.Inlines.Add(timestampRun);
                
                // Message avec couleur
                var messageRun = new Run($"{prefix} {message}")
                {
                    Foreground = new SolidColorBrush(color)
                };
                paragraph.Inlines.Add(messageRun);
                
                TxtLog.Document.Blocks.Add(paragraph);
                TxtLog.ScrollToEnd();
                
                // Propager vers le journal principal de MainWindow si callback défini
                MainWindowLogCallback?.Invoke($"[SmartTools] {message}", level);
            });
        }

        /// <summary>
        /// Callback pour afficher une popup HTML (iProperties, rapports, etc.)
        /// </summary>
        private void ShowHtmlPopup(string title, string htmlContent)
        {
            Dispatcher.Invoke(() =>
            {
                HtmlPopupWindow.ShowHtml(title, htmlContent, this);
            });
        }
        
        /// <summary>
        /// Callback pour afficher la fenêtre iProperties WPF avec les fonctions de modification
        /// </summary>
        private void ShowIPropertiesPopup(string title, string htmlContent, 
            Func<string, string, bool> applyPropertyChange, 
            Func<string, string, bool> addNewProperty)
        {
            // Note: htmlContent n'est plus utilisé - la nouvelle fenêtre WPF gère tout
            // Ce callback sera remplacé par ShowIPropertiesWindow dans une prochaine mise à jour
            Dispatcher.Invoke(() =>
            {
                var popup = new HtmlPopupWindow(title, htmlContent);
                popup.SetPropertyCallbacks(applyPropertyChange, addNewProperty);
                popup.ShowDialog();
            });
        }
        
        /// <summary>
        /// Callback pour afficher la fenêtre iProperties WPF native
        /// </summary>
        private void ShowIPropertiesWindow(string fileName, string fullPath, 
            Dictionary<string, string> properties,
            Func<string, string, bool> applyPropertyChange, 
            Func<string, string, bool> addNewProperty)
        {
            Dispatcher.Invoke(() =>
            {
                IPropertiesWindow.ShowProperties(fileName, fullPath, properties, applyPropertyChange, addNewProperty);
            });
        }

        /// <summary>
        /// Callback pour créer une fenêtre de progression HTML
        /// </summary>
        private IProgressWindow ShowProgressWindow(string title, string htmlContent)
        {
            IProgressWindow? result = null;
            Dispatcher.Invoke(() =>
            {
                var window = ProgressWindow.ShowProgress(title, htmlContent, this);
                result = new ProgressWindowWrapper(window);
            });
            return result!;
        }

        /// <summary>
        /// Callback pour afficher la fenêtre d'options d'export IAM
        /// </summary>
        private ExportOptionsResult? ShowExportOptions(string sourceFileName, string sourcePath)
        {
            ExportOptionsResult? result = null;
            Dispatcher.Invoke(() =>
            {
                var options = ExportOptionsWindow.ShowOptions(sourceFileName, sourcePath, this);
                if (options != null)
                {
                    result = new ExportOptionsResult
                    {
                        Format = options.SelectedFormat,
                        DestinationPath = options.DestinationPath,
                        OutputFileName = options.OutputFileName,
                        FullOutputPath = options.FullOutputPath,
                        HideReferences = options.HideReferences,
                        ActivateDefaultRepresentation = options.ActivateDefaultRepresentation,
                        OpenAfterExport = options.OpenAfterExport
                    };
                }
            });
            return result;
        }
    }

    /// <summary>
    /// Résultat des options d'export
    /// </summary>
    public class ExportOptionsResult
    {
        public ExportFormat Format { get; set; }
        public string DestinationPath { get; set; } = "";
        public string OutputFileName { get; set; } = "";
        public string FullOutputPath { get; set; } = "";
        public bool HideReferences { get; set; }
        public bool ActivateDefaultRepresentation { get; set; }
        public bool OpenAfterExport { get; set; }
    }
}
