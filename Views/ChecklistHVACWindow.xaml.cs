using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Views
{
    /// <summary>
    /// Checklist HVAC Window - Affiche la checklist HTML comme une application native
    /// Utilise WebView2 (Edge/Chromium) pour supporter React et JavaScript moderne
    /// </summary>
    public partial class ChecklistHVACWindow : Window
    {
        private readonly string _htmlFilePath;
        private readonly VaultSdkService _vaultService;

        public ChecklistHVACWindow(string htmlFilePath, VaultSdkService vaultService = null)
        {
            InitializeComponent();
            _htmlFilePath = htmlFilePath;
            _vaultService = vaultService;
            
            // Afficher statut Vault
            UpdateVaultStatus();
            
            // Initialiser WebView2 et charger le fichier HTML
            InitializeWebView();
        }
        
        /// <summary>
        /// Met a jour le statut Vault avec le format uniforme
        /// </summary>
        private void UpdateVaultStatus()
        {
            if (_vaultService != null && _vaultService.IsConnected)
            {
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(16, 124, 16)); // Vert
                StatusText.Text = $"🗄️ Vault : {_vaultService.VaultName}  /  👤 Utilisateur : {_vaultService.UserName}  /  📡 Statut : Connecte";
            }
            else
            {
                VaultStatusIndicator.Fill = new SolidColorBrush(Color.FromRgb(232, 17, 35)); // Rouge
                StatusText.Text = "🗄️ Vault : --  /  👤 Utilisateur : --  /  📡 Statut : Deconnecte";
            }
        }

        /// <summary>
        /// Initialise WebView2 avec le runtime Edge/Chromium
        /// </summary>
        private async void InitializeWebView()
        {
            try
            {
                FilePathText.Text = "Initialisation WebView2...";
                
                // S'assurer que WebView2 est prêt
                await WebViewControl.EnsureCoreWebView2Async(null);
                
                // Charger le fichier HTML
                LoadHtmlFile();
            }
            catch (Exception ex)
            {
                Logger.Log("Erreur WebView2: " + ex.Message, Logger.LogLevel.ERROR);
                FilePathText.Text = "Erreur WebView2";
                
                // Message d'erreur avec instruction d'installation
                MessageBox.Show(
                    "WebView2 Runtime n'est pas installé.\n\n" +
                    "Pour afficher la checklist, veuillez installer:\n" +
                    "Microsoft Edge WebView2 Runtime\n\n" +
                    "Téléchargez depuis:\n" +
                    "https://developer.microsoft.com/microsoft-edge/webview2/\n\n" +
                    "Erreur: " + ex.Message,
                    "WebView2 requis", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Warning);
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
                    MessageBox.Show("Fichier Checklist non trouvé:\n" + _htmlFilePath, 
                        "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Erreur chargement ChecklistHVAC: " + ex.Message, Logger.LogLevel.ERROR);
                FilePathText.Text = "Erreur de chargement";
                MessageBox.Show("Erreur: " + ex.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show("Erreur: " + ex.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void WebViewControl_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess)
            {
                FilePathText.Text = "[+] Pret - " + _htmlFilePath;
            }
            else
            {
                FilePathText.Text = "[-] Erreur de navigation";
                Logger.Log("WebView2 navigation error: " + e.WebErrorStatus, Logger.LogLevel.ERROR);
            }
        }
    }
}
