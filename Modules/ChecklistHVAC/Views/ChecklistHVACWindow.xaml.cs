using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Modules.ChecklistHVAC.Services;
using XnrgyEngineeringAutomationTools.Modules.ChecklistHVAC.Models;
using XnrgyEngineeringAutomationTools.Shared.Views;

namespace XnrgyEngineeringAutomationTools.Modules.ChecklistHVAC.Views
{
    /// <summary>
    /// Checklist HVAC Window - Affiche la checklist HTML comme une application native
    /// Utilise WebView2 (Edge/Chromium) pour supporter React et JavaScript moderne
    /// Synchronisation bidirectionnelle avec Vault toutes les 4-5 minutes
    /// </summary>
    public partial class ChecklistHVACWindow : Window
    {
        private readonly string _htmlFilePath;
        private readonly VaultSdkService _vaultService;
        private ChecklistSyncService? _syncService;
        private System.Windows.Threading.DispatcherTimer? _syncStatusTimer;

        // Module actuel (extrait depuis le chemin ou saisi par l'utilisateur)
        private string _currentProjectNumber = "";
        private string _currentReference = "";
        private string _currentModule = "";

        public ChecklistHVACWindow(string htmlFilePath, VaultSdkService vaultService = null)
        {
            InitializeComponent();
            _htmlFilePath = htmlFilePath;
            _vaultService = vaultService;
            
            // Initialiser le service de synchronisation si Vault est connecté
            if (_vaultService != null && _vaultService.IsConnected)
            {
                _syncService = new ChecklistSyncService(_vaultService);
                _syncService.SyncStatusChanged += OnSyncStatusChanged;
                
                // Démarrer la synchronisation automatique
                _syncService.StartAutoSync();
                
                // Timer pour mettre à jour le statut de sync toutes les 30 secondes
                _syncStatusTimer = new System.Windows.Threading.DispatcherTimer();
                _syncStatusTimer.Interval = TimeSpan.FromSeconds(30);
                _syncStatusTimer.Tick += (s, e) => UpdateSyncStatus();
                _syncStatusTimer.Start();
            }
            
            // Afficher statut Vault
            UpdateVaultStatus();
            
            // S'abonner aux changements de theme
            MainWindow.ThemeChanged += OnThemeChanged;
            this.Closed += (s, e) =>
            {
                MainWindow.ThemeChanged -= OnThemeChanged;
                _syncService?.StopAutoSync();
                _syncStatusTimer?.Stop();
                _syncService?.Dispose();
            };
            
            // Appliquer le theme actuel au demarrage
            ApplyTheme(MainWindow.CurrentThemeIsDark);
            
            // Initialiser WebView2 et charger le fichier HTML
            InitializeWebView();
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
        
        /// <summary>
        /// Met a jour le statut Vault avec le format uniforme
        /// </summary>
        private void UpdateVaultStatus()
        {
            if (_vaultService != null && _vaultService.IsConnected)
            {
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(16, 124, 16)); // Vert
                RunVaultName.Text = $" Vault: {_vaultService.VaultName}";
                RunUserName.Text = $" {_vaultService.UserName}";
                RunStatus.Text = " Connecte";
            }
            else
            {
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(232, 17, 35)); // Rouge
                RunVaultName.Text = " Vault: --";
                RunUserName.Text = " --";
                RunStatus.Text = " Deconnecte";
            }
        }

        /// <summary>
        /// Initialise WebView2 avec le runtime Edge/Chromium
        /// Configure le pont JavaScript ↔ C# pour la synchronisation
        /// </summary>
        private async void InitializeWebView()
        {
            try
            {
                FilePathText.Text = "Initialisation WebView2...";
                
                // S'assurer que WebView2 est prêt
                await WebViewControl.EnsureCoreWebView2Async(null);
                
                // Ajouter le pont JavaScript ↔ C#
                SetupJavaScriptBridge();
                
                // Charger le fichier HTML
                LoadHtmlFile();
            }
            catch (Exception ex)
            {
                Logger.Log("Erreur WebView2: " + ex.Message, Logger.LogLevel.ERROR);
                FilePathText.Text = "Erreur WebView2";
                
                // Message d'erreur avec instruction d'installation
                XnrgyMessageBox.ShowWarning(
                    "WebView2 Runtime n'est pas installe.\n\n" +
                    "Pour afficher la checklist, veuillez installer:\n" +
                    "Microsoft Edge WebView2 Runtime\n\n" +
                    "Telechargez depuis:\n" +
                    "https://developer.microsoft.com/microsoft-edge/webview2/\n\n" +
                    "Erreur: " + ex.Message,
                    "WebView2 requis", this);
            }
        }

        /// <summary>
        /// Configure le pont JavaScript ↔ C# pour communication bidirectionnelle
        /// Permet au HTML de sauvegarder/charger les données via C#
        /// </summary>
        private void SetupJavaScriptBridge()
        {
            try
            {
                if (WebViewControl.CoreWebView2 == null)
                {
                    Logger.Log("[ChecklistSync] WebView2 CoreWebView2 non initialisé", Logger.LogLevel.WARNING);
                    return;
                }

                // Créer un objet hôte C# pour JavaScript (WebView2 utilise AddHostObjectToScript)
                WebViewControl.CoreWebView2.AddHostObjectToScript("checklistHost", new ChecklistHostObject(this));
                
                // Injecter le script JavaScript pour exposer les fonctions
                // Note: WebView2 utilise window.chrome.webview.hostObjects pour accéder aux objets C#
                string bridgeScript = @"
                    (function() {
                        if (typeof window.checklistSync === 'undefined') {
                            window.checklistSync = {
                                saveData: function(moduleId, projectNumber, reference, module, data) {
                                    try {
                                        const host = window.chrome.webview.hostObjects.checklistHost;
                                        if (!host) {
                                            console.error('[ChecklistSync] Bridge C# non disponible');
                                            return 'ERROR: Bridge non disponible';
                                        }
                                        const json = typeof data === 'string' ? data : JSON.stringify(data);
                                        return host.SaveChecklistData(moduleId, projectNumber, reference, module, json);
                                    } catch (e) {
                                        console.error('[ChecklistSync] Erreur saveData:', e);
                                        return 'ERROR: ' + e.message;
                                    }
                                },
                                loadData: function(moduleId) {
                                    try {
                                        const host = window.chrome.webview.hostObjects.checklistHost;
                                        if (!host) {
                                            console.warn('[ChecklistSync] Bridge C# non disponible - utilisation localStorage');
                                            const localData = localStorage.getItem('checklist_' + moduleId);
                                            return localData ? JSON.parse(localData) : null;
                                        }
                                        const json = host.LoadChecklistData(moduleId);
                                        return json ? JSON.parse(json) : null;
                                    } catch (e) {
                                        console.error('[ChecklistSync] Erreur loadData:', e);
                                        return null;
                                    }
                                },
                                syncNow: function(moduleId, projectNumber, reference, module) {
                                    try {
                                        const host = window.chrome.webview.hostObjects.checklistHost;
                                        if (host) {
                                            host.SyncNow(moduleId, projectNumber, reference, module);
                                        }
                                    } catch (e) {
                                        console.error('[ChecklistSync] Erreur syncNow:', e);
                                    }
                                },
                                getSyncStatus: function() {
                                    try {
                                        const host = window.chrome.webview.hostObjects.checklistHost;
                                        if (host) {
                                            return host.GetSyncStatus();
                                        }
                                        return '{""connected"": false, ""autoSync"": false}';
                                    } catch (e) {
                                        return '{""connected"": false, ""error"": ""' + e.message + '""}';
                                    }
                                }
                            };
                            console.log('[ChecklistSync] Bridge JavaScript initialisé');
                        }
                    })();
                ";
                
                // Injecter le script après que le DOM soit chargé
                WebViewControl.CoreWebView2.DOMContentLoaded += async (sender, e) =>
                {
                    await WebViewControl.CoreWebView2.ExecuteScriptAsync(bridgeScript);
                    Logger.Log("[ChecklistSync] Script bridge JavaScript injecté", Logger.LogLevel.INFO);
                };
            }
            catch (Exception ex)
            {
                Logger.LogException("SetupJavaScriptBridge", ex, Logger.LogLevel.ERROR);
            }
        }

        private void LoadHtmlFile()
        {
            try
            {
                if (File.Exists(_htmlFilePath))
                {
                    // Naviguer vers le fichier HTML local
                    WebViewControl.CoreWebView2.Navigate(new Uri(_htmlFilePath).AbsoluteUri);
                    FilePathText.Text = "Chargement: " + _htmlFilePath;
                    Logger.Log("ChecklistHVAC charge avec WebView2: " + _htmlFilePath, Logger.LogLevel.INFO);
                }
                else
                {
                    FilePathText.Text = "Erreur: Fichier non trouve - " + _htmlFilePath;
                    XnrgyMessageBox.ShowWarning("Fichier Checklist non trouve:\n" + _htmlFilePath, 
                        "Erreur", this);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Erreur chargement ChecklistHVAC: " + ex.Message, Logger.LogLevel.ERROR);
                FilePathText.Text = "Erreur de chargement";
                XnrgyMessageBox.ShowError("Erreur: " + ex.Message, "Erreur", this);
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FilePathText.Text = "Rafraichissement...";
                WebViewControl.Reload();
                FilePathText.Text = "[+] Actualise - " + _htmlFilePath;
            }
            catch (Exception ex)
            {
                Logger.Log("Erreur refresh: " + ex.Message, Logger.LogLevel.ERROR);
            }
        }

        private void OpenExternalButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists(_htmlFilePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = _htmlFilePath,
                        UseShellExecute = true
                    });
                    FilePathText.Text = "[+] Ouvert dans le navigateur";
                }
            }
            catch (Exception ex)
            {
                XnrgyMessageBox.ShowError("Erreur: " + ex.Message, "Erreur", this);
            }
        }

        private void WebViewControl_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess)
            {
                FilePathText.Text = "[+] Pret - " + Path.GetFileName(_htmlFilePath);
                
                // Réinjecter le pont JavaScript après navigation
                SetupJavaScriptBridge();
            }
            else
            {
                FilePathText.Text = "[-] Erreur de navigation";
                Logger.Log("WebView2 navigation error: " + e.WebErrorStatus, Logger.LogLevel.ERROR);
            }
        }

        /// <summary>
        /// Gestionnaire d'événements de changement de statut de synchronisation
        /// </summary>
        private void OnSyncStatusChanged(string level, string message)
        {
            Dispatcher.Invoke(() =>
            {
                string icon = level switch
                {
                    "SUCCESS" => "[+]",
                    "ERROR" => "[-]",
                    "WARN" => "[!]",
                    _ => "[i]"
                };
                
                FilePathText.Text = $"{icon} Sync: {message}";
                
                Logger.Log($"[ChecklistSync] {message}", 
                    level == "ERROR" ? Logger.LogLevel.ERROR : 
                    level == "WARN" ? Logger.LogLevel.WARNING : Logger.LogLevel.INFO);
            });
        }

        /// <summary>
        /// Met à jour le statut de synchronisation
        /// </summary>
        private void UpdateSyncStatus()
        {
            if (_syncService != null && _vaultService != null && _vaultService.IsConnected)
            {
                Dispatcher.Invoke(() =>
                {
                    RunStatus.Text = " Statut : Synchronisation auto active";
                    VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(16, 124, 16)); // Vert
                });
            }
        }

        /// <summary>
        /// Synchronise manuellement le module actuel
        /// </summary>
        private async void SyncNowButton_Click(object sender, RoutedEventArgs e)
        {
            if (_syncService == null || string.IsNullOrEmpty(_currentProjectNumber))
            {
                XnrgyMessageBox.ShowInfo(
                    "Veuillez d'abord selectionner un module dans la checklist.",
                    "Module requis", this);
                return;
            }

            await _syncService.SyncModuleAsync(_currentProjectNumber, _currentReference, _currentModule);
        }

        /// <summary>
        /// Sauvegarde les données depuis JavaScript (appelé par le HTML)
        /// </summary>
        public string SaveChecklistData(string moduleId, string projectNumber, string reference, string module, string jsonData)
        {
            try
            {
                _currentProjectNumber = projectNumber;
                _currentReference = reference;
                _currentModule = module;

                // Désérialiser les données
                var data = System.Text.Json.JsonSerializer.Deserialize<ChecklistDataModel>(jsonData);
                if (data == null) return "ERROR: Données invalides";

                // Sauvegarder en cache local immédiatement
                string localPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XnrgyEngineeringAutomationTools",
                    "ChecklistHVAC",
                    $"Checklist_{moduleId}.json"
                );

                Directory.CreateDirectory(Path.GetDirectoryName(localPath));
                File.WriteAllText(localPath, jsonData, System.Text.Encoding.UTF8);

                // Synchroniser avec Vault en arrière-plan (non bloquant)
                if (_syncService != null)
                {
                    Task.Run(async () => await _syncService.SyncModuleAsync(projectNumber, reference, module, data));
                }

                return "OK";
            }
            catch (Exception ex)
            {
                Logger.LogException("SaveChecklistData", ex, Logger.LogLevel.ERROR);
                return $"ERROR: {ex.Message}";
            }
        }

        /// <summary>
        /// Charge les données depuis Vault (appelé par le HTML)
        /// </summary>
        public string LoadChecklistData(string moduleId)
        {
            try
            {
                if (_syncService == null) return null;

                // Charger depuis le cache local
                string localPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XnrgyEngineeringAutomationTools",
                    "ChecklistHVAC",
                    $"Checklist_{moduleId}.json"
                );

                if (File.Exists(localPath))
                {
                    string json = File.ReadAllText(localPath, System.Text.Encoding.UTF8);
                    return json;
                }

                return null;
            }
            catch (Exception ex)
            {
                Logger.LogException("LoadChecklistData", ex, Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// Synchronise maintenant (appelé par le HTML)
        /// </summary>
        public void SyncNow(string moduleId, string projectNumber, string reference, string module)
        {
            if (_syncService != null)
            {
                Task.Run(async () => await _syncService.SyncModuleAsync(projectNumber, reference, module));
            }
        }

        /// <summary>
        /// Obtient le statut de synchronisation (appelé par le HTML)
        /// </summary>
        public string GetSyncStatus()
        {
            if (_syncService == null || _vaultService == null || !_vaultService.IsConnected)
            {
                return "{\"connected\": false, \"autoSync\": false}";
            }

            return "{\"connected\": true, \"autoSync\": true, \"interval\": 4}";
        }
    }

    /// <summary>
    /// Objet hôte pour communication JavaScript ↔ C#
    /// Doit être public et utiliser ComVisible
    /// </summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public class ChecklistHostObject
    {
        private readonly ChecklistHVACWindow _window;

        public ChecklistHostObject(ChecklistHVACWindow window)
        {
            _window = window;
        }

        public string SaveChecklistData(string moduleId, string projectNumber, string reference, string module, string jsonData)
        {
            return _window.SaveChecklistData(moduleId, projectNumber, reference, module, jsonData);
        }

        public string LoadChecklistData(string moduleId)
        {
            return _window.LoadChecklistData(moduleId) ?? "";
        }

        public void SyncNow(string moduleId, string projectNumber, string reference, string module)
        {
            _window.SyncNow(moduleId, projectNumber, reference, module);
        }

        public string GetSyncStatus()
        {
            return _window.GetSyncStatus();
        }
    }
}
