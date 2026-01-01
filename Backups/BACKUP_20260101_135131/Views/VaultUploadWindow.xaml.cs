using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using XnrgyEngineeringAutomationTools.ViewModels;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Views
{
    public partial class VaultUploadWindow : Window
    {
        private readonly AppMainViewModel _viewModel;
        private readonly VaultSdkService _vaultService;
        private readonly InventorService _inventorService;

        public VaultUploadWindow(VaultSdkService vaultService, InventorService inventorService)
        {
            _vaultService = vaultService;
            _inventorService = inventorService;
            _viewModel = new AppMainViewModel();
            InitializeComponent();
            DataContext = _viewModel;
            InitializePasswordBox();
            Logger.Log("VaultUploadWindow initialisee", Logger.LogLevel.INFO);
        }

        private void InitializePasswordBox()
        {
            try
            {
                var passwordBox = FindName("PasswordBox") as PasswordBox;
                if (passwordBox != null)
                {
                    passwordBox.PasswordChanged += PasswordBox_PasswordChanged;
                }
            }
            catch { }
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            // Le PasswordBox ne supporte pas le binding direct
            // Le ViewModel gere le mot de passe autrement
        }

        private void DataGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            // Handler pour le DataGrid
        }

        private void FromInventorButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_inventorService != null && _inventorService.IsConnected)
                {
                    string path = _inventorService.GetActiveDocumentPath();
                    if (!string.IsNullOrEmpty(path))
                    {
                        string directory = System.IO.Path.GetDirectoryName(path);
                        // Le ViewModel expose ModulePath via Property
                    }
                    else
                    {
                        MessageBox.Show("Aucun document actif dans Inventor.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Inventor non connecte.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur: " + ex.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            Logger.Log("VaultUploadWindow fermee", Logger.LogLevel.INFO);
        }
    }
}