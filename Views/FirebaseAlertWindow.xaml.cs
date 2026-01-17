using System.Windows;
using XnrgyEngineeringAutomationTools.Models;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Views
{
    /// <summary>
    /// Fenetre d'alerte Firebase pour afficher les messages de controle a distance
    /// </summary>
    public partial class FirebaseAlertWindow : Window
    {
        public enum AlertType
        {
            KillSwitch,
            Maintenance,
            UpdateAvailable,
            ForceUpdate,
            UserDisabled,
            BroadcastInfo,
            BroadcastWarning,
            BroadcastError
        }

        public bool ShouldContinue { get; private set; }
        public bool ShouldDownload { get; private set; }

        private string _downloadUrl;

        public FirebaseAlertWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Configure et affiche l'alerte Kill Switch
        /// </summary>
        public static bool ShowKillSwitch(string message)
        {
            var window = new FirebaseAlertWindow();
            window._downloadUrl = null;
            window.ConfigureKillSwitch(message);
            window.ShowDialog();
            return false; // Toujours bloquer l'application
        }

        /// <summary>
        /// Configure et affiche l'alerte Utilisateur desactive
        /// </summary>
        public static bool ShowUserDisabled(string message)
        {
            var window = new FirebaseAlertWindow();
            window.ConfigureUserDisabled(message);
            window.ShowDialog();
            return false; // Toujours bloquer l'application
        }

        /// <summary>
        /// Configure et affiche l'alerte Maintenance
        /// </summary>
        public static bool ShowMaintenance(string message)
        {
            var window = new FirebaseAlertWindow();
            window.ConfigureMaintenance(message);
            window.ShowDialog();
            return false; // Toujours bloquer l'application
        }

        /// <summary>
        /// Configure et affiche l'alerte de mise a jour disponible
        /// </summary>
        public static (bool shouldContinue, bool shouldDownload) ShowUpdateAvailable(
            string currentVersion, string newVersion, string changelog, string downloadUrl, bool forceUpdate)
        {
            var window = new FirebaseAlertWindow();
            window.ConfigureUpdate(currentVersion, newVersion, changelog, downloadUrl, forceUpdate);
            window.ShowDialog();
            return (window.ShouldContinue, window.ShouldDownload);
        }

        /// <summary>
        /// Affiche un message broadcast
        /// Retourne true si l'application doit etre bloquee (type "error")
        /// </summary>
        public static bool ShowBroadcastMessage(string title, string message, string type)
        {
            var window = new FirebaseAlertWindow();
            window.ConfigureBroadcast(title, message, type);
            window.ShowDialog();
            
            // Bloquer seulement pour les messages de type "error"
            return type?.ToLowerInvariant() == "error" && !window.ShouldContinue;
        }

        private void ConfigureKillSwitch(string message)
        {
            AlertIcon.Text = "[X]";
            AlertTitle.Text = "Application Desactivee";
            AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(255, 100, 100));
            
            AlertMessage.Text = message ?? "Cette application a ete desactivee par l'administrateur.\n\n" +
                "Contactez le support technique pour plus d'informations.";
            
            PrimaryButton.Content = "Fermer";
            SecondaryButton.Visibility = Visibility.Collapsed;
            VersionInfoPanel.Visibility = Visibility.Collapsed;
        }

        private void ConfigureUserDisabled(string message)
        {
            AlertIcon.Text = "[!]";
            AlertTitle.Text = "Acces refuse";
            AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(255, 100, 100));
            
            AlertMessage.Text = message ?? "Votre compte a ete desactive par l'administrateur.\n\n" +
                "Contactez votre superviseur pour plus d'informations.";
            
            PrimaryButton.Content = "Fermer";
            SecondaryButton.Visibility = Visibility.Collapsed;
            VersionInfoPanel.Visibility = Visibility.Collapsed;
        }

        private void ConfigureMaintenance(string message)
        {
            AlertIcon.Text = "[~]";
            AlertTitle.Text = "Maintenance en cours";
            AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(255, 193, 7));
            
            AlertMessage.Text = message ?? "L'application est actuellement en maintenance.\n\n" +
                "Veuillez reessayer dans quelques minutes.";
            
            PrimaryButton.Content = "Fermer";
            SecondaryButton.Visibility = Visibility.Collapsed;
            VersionInfoPanel.Visibility = Visibility.Collapsed;
        }

        private void ConfigureBroadcast(string title, string message, string type)
        {
            type = type?.ToLowerInvariant() ?? "info";
            
            switch (type)
            {
                case "error":
                    AlertIcon.Text = "[-]";
                    AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(255, 100, 100));
                    PrimaryButton.Content = "Fermer";
                    SecondaryButton.Visibility = Visibility.Collapsed;
                    break;
                    
                case "warning":
                    AlertIcon.Text = "[!]";
                    AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(255, 193, 7));
                    PrimaryButton.Content = "Compris";
                    SecondaryButton.Visibility = Visibility.Collapsed;
                    ShouldContinue = true; // Warning ne bloque pas
                    break;
                    
                default: // info
                    AlertIcon.Text = "[i]";
                    AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0, 212, 255));
                    PrimaryButton.Content = "OK";
                    SecondaryButton.Visibility = Visibility.Collapsed;
                    ShouldContinue = true; // Info ne bloque pas
                    break;
            }
            
            AlertTitle.Text = title ?? "Message";
            AlertMessage.Text = message ?? "";
            VersionInfoPanel.Visibility = Visibility.Collapsed;
        }

        private void ConfigureUpdate(string currentVersion, string newVersion, 
            string changelog, string downloadUrl, bool forceUpdate)
        {
            _downloadUrl = downloadUrl;

            if (forceUpdate)
            {
                AlertIcon.Text = "[!]";
                AlertTitle.Text = "Mise a jour requise";
                AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(255, 100, 100));
                
                AlertMessage.Text = "Une mise a jour obligatoire est disponible.\n\n" +
                    "Vous devez telecharger la nouvelle version pour continuer a utiliser l'application.";
                
                PrimaryButton.Content = "Telecharger";
                SecondaryButton.Visibility = Visibility.Collapsed;
            }
            else
            {
                AlertIcon.Text = "[+]";
                AlertTitle.Text = "Mise a jour disponible";
                AlertTitle.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0, 212, 255));
                
                AlertMessage.Text = "Une nouvelle version est disponible.\n\n" +
                    "Souhaitez-vous telecharger la mise a jour maintenant?";
                
                PrimaryButton.Content = "Telecharger";
                SecondaryButton.Content = "Plus tard";
                SecondaryButton.Visibility = Visibility.Visible;
            }

            // Afficher les informations de version
            VersionInfoPanel.Visibility = Visibility.Visible;
            string versionText = $"Version actuelle: {currentVersion}\n" +
                                 $"Nouvelle version: {newVersion}";
            
            if (!string.IsNullOrEmpty(changelog))
            {
                versionText += $"\n\nChangements:\n{changelog}";
            }
            
            VersionDetails.Text = versionText;
        }

        private void PrimaryButton_Click(object sender, RoutedEventArgs e)
        {
            // Si c'est un bouton de telechargement
            if (PrimaryButton.Content.ToString() == "Telecharger" && !string.IsNullOrEmpty(_downloadUrl))
            {
                ShouldDownload = true;
                FirebaseRemoteConfigService.OpenDownloadUrl(_downloadUrl);
            }

            // Pour Kill Switch et Maintenance, ne pas continuer
            // Pour ForceUpdate, ne pas continuer
            // Pour Update optionnel avec telechargement, ne pas continuer
            ShouldContinue = false;
            
            Close();
        }

        private void SecondaryButton_Click(object sender, RoutedEventArgs e)
        {
            // Bouton "Plus tard" - continuer sans telecharger
            ShouldContinue = true;
            ShouldDownload = false;
            Close();
        }
    }
}
