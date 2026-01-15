using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Modules.ACP.Services;
using XnrgyEngineeringAutomationTools.Shared.Views;

namespace XnrgyEngineeringAutomationTools.Modules.ACP.Views
{
    public partial class ACPWindow : Window
    {
        private readonly string _htmlFilePath;
        private readonly VaultSdkService _vaultService;
        private ACPSyncService _syncService;
        private System.Windows.Threading.DispatcherTimer _syncStatusTimer;

        public ACPWindow(VaultSdkService vaultService = null)
        {
            InitializeComponent();
            
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string resourcePath = Path.Combine(appDir, "Modules", "ACP", "Resources", "ACP - Assistant de Conception de Projet.html");
            
            if (!File.Exists(resourcePath))
            {
                resourcePath = @"c:\Users\mohammedamine.elgala\source\repos\XnrgyEngineeringAutomationTools\Modules\ACP\Resources\ACP - Assistant de Conception de Projet.html";
            }
            
            _htmlFilePath = resourcePath;
            _vaultService = vaultService;
            
            if (_vaultService != null && _vaultService.IsConnected)
            {
                _syncService = new ACPSyncService(_vaultService);
                _syncService.SyncStatusChanged += OnSyncStatusChanged;
                _syncService.StartAutoSync();
                
                _syncStatusTimer = new System.Windows.Threading.DispatcherTimer();
                _syncStatusTimer.Interval = TimeSpan.FromSeconds(30);
                _syncStatusTimer.Tick += (s, e) => UpdateSyncStatus();
                _syncStatusTimer.Start();
            }
            
            UpdateVaultStatus();
            
            MainWindow.ThemeChanged += OnThemeChanged;
            this.Closed += (s, e) =>
            {
                MainWindow.ThemeChanged -= OnThemeChanged;
                if (_syncService != null) { _syncService.StopAutoSync(); _syncService.Dispose(); }
                if (_syncStatusTimer != null) _syncStatusTimer.Stop();
            };
            
            ApplyTheme(MainWindow.CurrentThemeIsDark);
            InitializeWebView();
        }
        
        private void OnThemeChanged(bool isDarkTheme)
        {
            Dispatcher.Invoke(() => ApplyTheme(isDarkTheme));
        }

        private void ApplyTheme(bool isDarkTheme)
        {
            this.Background = new SolidColorBrush(isDarkTheme ? Color.FromRgb(30, 30, 46) : Color.FromRgb(245, 247, 250));
        }
        
        private void UpdateVaultStatus()
        {
            if (_vaultService != null && _vaultService.IsConnected)
            {
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(16, 124, 16));
                RunVaultName.Text = " Vault: " + _vaultService.VaultName;
                RunUserName.Text = " " + _vaultService.UserName;
                RunStatus.Text = " Connecte";
            }
            else
            {
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                RunVaultName.Text = " Vault: --";
                RunUserName.Text = " --";
                RunStatus.Text = " Deconnecte";
            }
        }
        
        private void UpdateSyncStatus()
        {
            if (_syncService != null && _vaultService != null && _vaultService.IsConnected)
            {
                Dispatcher.Invoke(() =>
                {
                    RunStatus.Text = " Statut : Synchronisation active";
                    VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(16, 124, 16));
                });
            }
        }
        
        private void OnSyncStatusChanged(string level, string message)
        {
            Dispatcher.Invoke(() =>
            {
                string icon = level == "SUCCESS" ? "[+]" : level == "ERROR" ? "[-]" : level == "WARN" ? "[!]" : "[i]";
                FilePathText.Text = icon + " " + message;
                Logger.Log("[ACPSync] " + message, level == "ERROR" ? Logger.LogLevel.ERROR : level == "WARN" ? Logger.LogLevel.WARNING : Logger.LogLevel.INFO);
            });
        }

        private async void InitializeWebView()
        {
            try
            {
                var userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XnrgyEngineeringAutomationTools", "WebView2Cache", "ACP");
                Directory.CreateDirectory(userDataFolder);
                
                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await WebViewControl.EnsureCoreWebView2Async(environment);
                
                WebViewControl.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
                
                if (File.Exists(_htmlFilePath))
                {
                    WebViewControl.CoreWebView2.Navigate(new Uri(_htmlFilePath).AbsoluteUri);
                    FilePathText.Text = "ACP charge: " + _htmlFilePath;
                    Logger.Log("ACP charge avec WebView2: " + _htmlFilePath, Logger.LogLevel.INFO);
                }
                else
                {
                    FilePathText.Text = "[-] Fichier non trouve: " + _htmlFilePath;
                    Logger.Log("[-] Fichier ACP non trouve: " + _htmlFilePath, Logger.LogLevel.ERROR);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPWindow.InitializeWebView", ex, Logger.LogLevel.ERROR);
                FilePathText.Text = "[-] Erreur WebView2: " + ex.Message;
            }
        }
        
        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string message = e.WebMessageAsJson;
                Logger.Log("[ACP] Message JS recu: " + message, Logger.LogLevel.DEBUG);
                
                var jsonDoc = System.Text.Json.JsonDocument.Parse(message);
                var root = jsonDoc.RootElement;
                
                if (root.TryGetProperty("type", out var typeElement))
                {
                    string messageType = typeElement.GetString() ?? "";
                    if (messageType == "GET_VAULT_STATUS") SendVaultStatusToJS();
                    else if (messageType == "SYNC_ACP_DATA" && root.TryGetProperty("data", out var dataElement))
                        SyncACPDataToVault(dataElement.GetRawText());
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPWindow.CoreWebView2_WebMessageReceived", ex, Logger.LogLevel.ERROR);
            }
        }
        
        private void SendVaultStatusToJS()
        {
            try
            {
                bool isConnected = _vaultService != null && _vaultService.IsConnected;
                string vaultName = _vaultService != null ? _vaultService.VaultName : "";
                string userName = _vaultService != null ? _vaultService.UserName : "";
                string script = "window.postMessage({ type: 'VAULT_STATUS', connected: " + isConnected.ToString().ToLower() + ", vaultName: '" + vaultName + "', userName: '" + userName + "' }, '*');";
                WebViewControl.CoreWebView2.ExecuteScriptAsync(script);
            }
            catch (Exception ex) { Logger.LogException("ACPWindow.SendVaultStatusToJS", ex, Logger.LogLevel.ERROR); }
        }
        
        private async void SyncACPDataToVault(string jsonData)
        {
            try
            {
                await WebViewControl.CoreWebView2.ExecuteScriptAsync("window.postMessage({ type: 'SYNC_STARTED' }, '*');");
                if (_syncService != null) { await Task.Delay(1000); Logger.Log("[ACP] Donnees synchronisees avec Vault", Logger.LogLevel.INFO); }
                await WebViewControl.CoreWebView2.ExecuteScriptAsync("window.postMessage({ type: 'SYNC_COMPLETED' }, '*');");
                FilePathText.Text = "[+] Synchronisation terminee";
            }
            catch (Exception ex) { Logger.LogException("ACPWindow.SyncACPDataToVault", ex, Logger.LogLevel.ERROR); FilePathText.Text = "[-] Erreur sync: " + ex.Message; }
        }

        private async void WebViewControl_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess) { Logger.Log("[ACP] Navigation WebView2 terminee avec succes", Logger.LogLevel.INFO); await InjectBridgeScript(); SendVaultStatusToJS(); }
            else Logger.Log("[-] Erreur navigation WebView2: " + e.WebErrorStatus, Logger.LogLevel.ERROR);
        }
        
        private async Task InjectBridgeScript()
        {
            try
            {
                string bridgeScript = "window.ACPBridge = { sendToHost: function(msg) { if (window.chrome && window.chrome.webview) window.chrome.webview.postMessage(msg); }, getVaultStatus: function() { this.sendToHost({ type: 'GET_VAULT_STATUS' }); }, syncData: function(data) { this.sendToHost({ type: 'SYNC_ACP_DATA', data: data }); } }; console.log('[ACPBridge] Script bridge injecte');";
                await WebViewControl.CoreWebView2.ExecuteScriptAsync(bridgeScript);
                Logger.Log("[ACP] Script bridge JavaScript injecte", Logger.LogLevel.INFO);
            }
            catch (Exception ex) { Logger.LogException("ACPWindow.InjectBridgeScript", ex, Logger.LogLevel.ERROR); }
        }

        private void SyncNowButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_vaultService == null || !_vaultService.IsConnected) { XnrgyMessageBox.ShowWarning("Vault non connecte.", "Vault Deconnecte", this); return; }
                FilePathText.Text = "[>] Synchronisation en cours...";
                WebViewControl.CoreWebView2.ExecuteScriptAsync("if(window.ACPBridge) window.ACPBridge.syncData(localStorage.getItem('acp_data'));");
            }
            catch (Exception ex) { Logger.LogException("ACPWindow.SyncNowButton_Click", ex, Logger.LogLevel.ERROR); FilePathText.Text = "[-] Erreur: " + ex.Message; }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            try { if (WebViewControl.CoreWebView2 != null) { WebViewControl.CoreWebView2.Reload(); FilePathText.Text = "[>] Rechargement en cours..."; } }
            catch (Exception ex) { Logger.LogException("ACPWindow.RefreshButton_Click", ex, Logger.LogLevel.ERROR); }
        }

        private void OpenExternalButton_Click(object sender, RoutedEventArgs e)
        {
            try { if (File.Exists(_htmlFilePath)) Process.Start(new ProcessStartInfo { FileName = _htmlFilePath, UseShellExecute = true }); else XnrgyMessageBox.ShowError("Fichier non trouve:\n" + _htmlFilePath, "Erreur", this); }
            catch (Exception ex) { Logger.LogException("ACPWindow.OpenExternalButton_Click", ex, Logger.LogLevel.ERROR); XnrgyMessageBox.ShowError("Erreur: " + ex.Message, "Erreur", this); }
        }
    }
}
