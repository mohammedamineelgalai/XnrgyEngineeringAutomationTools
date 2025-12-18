using System;
using System.Windows;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Views
{
    public partial class LoginWindow : Window
    {
        private readonly VaultSdkService _vaultService;

        public LoginWindow(VaultSdkService vaultService)
        {
            InitializeComponent();
            _vaultService = vaultService;
            
            // Charger les identifiants sauvegardés
            LoadSavedCredentials();
        }

        private void LoadSavedCredentials()
        {
            try
            {
                // TODO: Charger depuis appsettings.json
                // Pour l'instant, utiliser les valeurs par défaut
            }
            catch { }
        }

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorMessage.Text = "";
            
            string server = ServerTextBox.Text.Trim();
            string vault = VaultTextBox.Text.Trim();
            string user = UserTextBox.Text.Trim();
            string password = PasswordBox.Password;

            if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(vault) || string.IsNullOrEmpty(user))
            {
                ErrorMessage.Text = "Veuillez remplir tous les champs.";
                return;
            }

            try
            {
                ConnectButton.IsEnabled = false;
                ConnectButton.Content = "Connexion...";
                
                Logger.Log($"🔌 Tentative de connexion à {server}/{vault}...", Logger.LogLevel.INFO);
                
                bool success = _vaultService.Connect(server, vault, user, password);
                
                if (success)
                {
                    Logger.Log($"✅ Connexion réussie à {server}/{vault}", Logger.LogLevel.INFO);
                    
                    // Sauvegarder si demandé
                    if (SaveCredentialsCheckBox.IsChecked == true)
                    {
                        // TODO: Sauvegarder dans appsettings.json
                    }
                    
                    DialogResult = true;
                    Close();
                }
                else
                {
                    ErrorMessage.Text = "Échec de la connexion. Vérifiez vos identifiants.";
                    Logger.Log($"❌ Échec de connexion à {server}/{vault}", Logger.LogLevel.ERROR);
                }
            }
            catch (Exception ex)
            {
                ErrorMessage.Text = $"Erreur: {ex.Message}";
                Logger.Log($"❌ Erreur connexion: {ex.Message}", Logger.LogLevel.ERROR);
            }
            finally
            {
                ConnectButton.IsEnabled = true;
                ConnectButton.Content = "🔌 Connecter";
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
