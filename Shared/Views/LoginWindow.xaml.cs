using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Animation;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Shared.Views
{
    public partial class LoginWindow : Window
    {
        private readonly VaultSdkService _vaultService;
        private Storyboard _spinnerStoryboard;
        private bool _autoConnectMode = false;
        private bool _isConnecting = false;
        private bool _closeAllowed = false;

        public LoginWindow(VaultSdkService vaultService, bool autoConnect = false)
        {
            InitializeComponent();
            _vaultService = vaultService;
            _autoConnectMode = autoConnect;
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Charger les identifiants sauvegardés
            LoadSavedCredentials();
            
            // Préparer l'animation du spinner
            _spinnerStoryboard = (Storyboard)FindResource("SpinnerAnimation");
            
            // Si mode auto-connect et credentials complets, tenter la connexion auto
            if (_autoConnectMode)
            {
                await TryAutoConnect();
            }
        }
        
        /// <summary>
        /// Tente une connexion automatique si les credentials sont sauvegardés
        /// </summary>
        private async Task TryAutoConnect()
        {
            try
            {
                var credentials = CredentialsManager.Load();
                
                // Vérifier si credentials complets
                if (credentials.SaveCredentials && 
                    !string.IsNullOrEmpty(credentials.Username) &&
                    !string.IsNullOrEmpty(credentials.Password) &&
                    !string.IsNullOrEmpty(credentials.Server) &&
                    !string.IsNullOrEmpty(credentials.VaultName))
                {
                    // Désactiver les contrôles et afficher le spinner
                    SetControlsEnabled(false);
                    ShowConnectionProgress("Connexion automatique en cours...");
                    
                    Logger.Log($"[>] Connexion automatique a {credentials.Server}/{credentials.VaultName}...", Logger.LogLevel.INFO);
                    
                    // Connexion asynchrone
                    bool success = await Task.Run(() => 
                        _vaultService.Connect(credentials.Server, credentials.VaultName, credentials.Username, credentials.Password));
                    
                    if (success)
                    {
                        ShowConnectionProgress("[+] Connexion reussie!");
                        Logger.Log($"[+] Connexion automatique reussie ({credentials.Username})", Logger.LogLevel.INFO);
                        
                        await Task.Delay(800); // Délai pour voir le succès
                        
                        _closeAllowed = true;  // Autoriser la fermeture apres succes
                        DialogResult = true;
                        Close();
                        return;
                    }
                    else
                    {
                        Logger.Log($"[!] Connexion automatique echouee - Intervention requise", Logger.LogLevel.WARNING);
                        ShowError("Connexion automatique échouée. Veuillez vérifier vos identifiants.");
                    }
                }
                else
                {
                    Logger.Log("[i] Pas de credentials sauvegardes - Intervention utilisateur requise", Logger.LogLevel.INFO);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur connexion auto: {ex.Message}", Logger.LogLevel.WARNING);
                ShowError($"Erreur: {ex.Message}");
            }
            finally
            {
                HideConnectionProgress();
                SetControlsEnabled(true);
            }
        }

        private void LoadSavedCredentials()
        {
            try
            {
                var credentials = CredentialsManager.Load();
                
                // Toujours charger serveur/vault
                ServerTextBox.Text = credentials.Server;
                VaultTextBox.Text = credentials.VaultName;
                
                // Si les credentials sont sauvegardés, les charger
                if (credentials.SaveCredentials && !string.IsNullOrEmpty(credentials.Username))
                {
                    UserTextBox.Text = credentials.Username;
                    PasswordBox.Password = credentials.Password;
                    SaveCredentialsCheckBox.IsChecked = true;
                    Logger.Log($"[+] Credentials charges pour {credentials.Username}", Logger.LogLevel.INFO);
                }
                else
                {
                    // Champs vides pour forcer la saisie
                    UserTextBox.Text = "";
                    PasswordBox.Password = "";
                    SaveCredentialsCheckBox.IsChecked = false;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur chargement credentials: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }

        private async void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            // Cacher l'erreur précédente
            ErrorBorder.Visibility = Visibility.Collapsed;
            ErrorMessage.Text = "";
            
            string server = ServerTextBox.Text.Trim();
            string vault = VaultTextBox.Text.Trim();
            string user = UserTextBox.Text.Trim();
            string password = PasswordBox.Password;

            if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(vault) || string.IsNullOrEmpty(user))
            {
                ShowError("Veuillez remplir tous les champs obligatoires.");
                return;
            }

            try
            {
                // Désactiver les contrôles et afficher le spinner
                SetControlsEnabled(false);
                ShowConnectionProgress("Connexion à Vault en cours...");
                
                Logger.Log($"[>] Tentative de connexion a {server}/{vault}...", Logger.LogLevel.INFO);
                
                // Connexion asynchrone pour ne pas bloquer l'UI
                bool success = await Task.Run(() => _vaultService.Connect(server, vault, user, password));
                
                if (success)
                {
                    ShowConnectionProgress("Connexion reussie!");
                    await Task.Delay(500); // Petit délai pour voir le succès
                    
                    Logger.Log($"[+] Connexion reussie a {server}/{vault}", Logger.LogLevel.INFO);
                    
                    // Sauvegarder les credentials
                    SaveCredentials(server, vault, user, password);
                    
                    _closeAllowed = true;  // Autoriser la fermeture apres succes
                    DialogResult = true;
                    Close();
                }
                else
                {
                    ShowError("Echec de la connexion. Verifiez vos identifiants.");
                    Logger.Log($"[-] Echec de connexion a {server}/{vault}", Logger.LogLevel.ERROR);
                }
            }
            catch (Exception ex)
            {
                ShowError($"Erreur: {ex.Message}");
                Logger.Log($"[-] Erreur connexion: {ex.Message}", Logger.LogLevel.ERROR);
            }
            finally
            {
                HideConnectionProgress();
                SetControlsEnabled(true);
            }
        }

        private void SaveCredentials(string server, string vault, string user, string password)
        {
            var credentials = new CredentialsManager.VaultCredentials
            {
                Server = server,
                VaultName = vault,
                Username = SaveCredentialsCheckBox.IsChecked == true ? user : "",
                Password = SaveCredentialsCheckBox.IsChecked == true ? password : "",
                SaveCredentials = SaveCredentialsCheckBox.IsChecked == true
            };
            CredentialsManager.Save(credentials);
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            UserTextBox.Text = "";
            PasswordBox.Password = "";
            SaveCredentialsCheckBox.IsChecked = false;
            CredentialsManager.Clear();
            ErrorBorder.Visibility = Visibility.Collapsed;
            Logger.Log("[i] Champs et credentials effacés", Logger.LogLevel.INFO);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// Empeche la fermeture de la fenetre pendant le processus de connexion
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Permettre la fermeture si autorisee (succes) ou si pas en cours de connexion
            if (_isConnecting && !_closeAllowed)
            {
                // Empecher la fermeture pendant le processus
                e.Cancel = true;
                XnrgyMessageBox.ShowWarning(
                    "La connexion est en cours.\nVeuillez attendre la fin du processus.",
                    "Fermeture impossible",
                    this);
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            _isConnecting = !enabled;  // Connexion en cours si controles desactives
            ServerTextBox.IsEnabled = enabled;
            VaultTextBox.IsEnabled = enabled;
            UserTextBox.IsEnabled = enabled;
            PasswordBox.IsEnabled = enabled;
            SaveCredentialsCheckBox.IsEnabled = enabled;
            ConnectButton.IsEnabled = enabled;
            CancelButton.IsEnabled = enabled;
            ClearButton.IsEnabled = enabled;
            ConnectButton.Content = enabled ? "[>] Connecter" : "Connexion...";
        }

        private void ShowConnectionProgress(string message)
        {
            ConnectionStatusText.Text = message;
            ConnectionSpinner.Visibility = Visibility.Visible;
            _spinnerStoryboard?.Begin(this, true);
        }

        private void HideConnectionProgress()
        {
            ConnectionSpinner.Visibility = Visibility.Collapsed;
            _spinnerStoryboard?.Stop(this);
        }

        private void ShowError(string message)
        {
            ErrorMessage.Text = message;
            ErrorBorder.Visibility = Visibility.Visible;
        }
    }
}
