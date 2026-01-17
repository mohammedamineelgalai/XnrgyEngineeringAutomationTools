using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace XnrgyInstaller
{
    /// <summary>
    /// Code-behind pour InstallerWindow.xaml
    /// Installation multi-etapes XNRGY Engineering Automation Tools
    /// </summary>
    public partial class InstallerWindow : Window, INotifyPropertyChanged
    {
        #region Fields

        private int _currentPage = 1;
        private const int TOTAL_PAGES = 5;
        private readonly InstallationService _installService;
        private DispatcherTimer _spinnerTimer;
        private bool _isInstalling = false;

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Constructor

        public InstallerWindow()
        {
            InitializeComponent();
            _installService = new InstallationService();
            
            LoadLicenseText();
            UpdateDiskSpace();
            InitializeSpinnerAnimation();
            UpdateNavigationButtons();
        }

        #endregion

        #region Window Events

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Permet de deplacer la fenetre sans bordure
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
            {
                DragMove();
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isInstalling)
            {
                var result = System.Windows.MessageBox.Show(
                    "L'installation est en cours. Voulez-vous vraiment annuler ?",
                    "Confirmer l'annulation",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                
                if (result != MessageBoxResult.Yes)
                    return;
            }
            
            System.Windows.Application.Current.Shutdown();
        }

        #endregion

        #region Navigation

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage > 1 && !_isInstalling)
            {
                _currentPage--;
                ShowCurrentPage();
                UpdateNavigationButtons();
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateCurrentPage())
                return;

            if (_currentPage < TOTAL_PAGES)
            {
                _currentPage++;
                ShowCurrentPage();
                UpdateNavigationButtons();
            }
        }

        private async void InstallButton_Click(object sender, RoutedEventArgs e)
        {
            _currentPage = 4; // Page progres
            ShowCurrentPage();
            UpdateNavigationButtons();
            
            await StartInstallationAsync();
        }

        private void FinishButton_Click(object sender, RoutedEventArgs e)
        {
            // Lancer l'app si option cochee
            if (LaunchAfterInstall.IsChecked == true)
            {
                string exePath = Path.Combine(InstallPathTextBox.Text, "XnrgyEngineeringAutomationTools.exe");
                if (File.Exists(exePath))
                {
                    Process.Start(exePath);
                }
            }
            
            System.Windows.Application.Current.Shutdown();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            CloseButton_Click(sender, e);
        }

        private bool ValidateCurrentPage()
        {
            switch (_currentPage)
            {
                case 2: // Licence
                    if (AcceptLicenseCheckBox.IsChecked != true)
                    {
                        System.Windows.MessageBox.Show(
                            "Vous devez accepter le contrat de licence pour continuer.",
                            "Licence requise",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return false;
                    }
                    break;

                case 3: // Dossier
                    if (string.IsNullOrWhiteSpace(InstallPathTextBox.Text))
                    {
                        System.Windows.MessageBox.Show(
                            "Veuillez specifier un dossier d'installation.",
                            "Dossier requis",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return false;
                    }
                    
                    // Verifier espace disque
                    try
                    {
                        string root = Path.GetPathRoot(InstallPathTextBox.Text);
                        DriveInfo drive = new DriveInfo(root);
                        long requiredSpace = 200 * 1024 * 1024; // 200 MB
                        
                        if (drive.AvailableFreeSpace < requiredSpace)
                        {
                            System.Windows.MessageBox.Show(
                                $"Espace disque insuffisant sur {root}\nRequis: 200 MB\nDisponible: {drive.AvailableFreeSpace / (1024 * 1024)} MB",
                                "Espace insuffisant",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show(
                            $"Impossible de verifier l'espace disque: {ex.Message}",
                            "Erreur",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                        return false;
                    }
                    break;
            }
            
            return true;
        }

        private void ShowCurrentPage()
        {
            // Cacher toutes les pages
            WelcomePage.Visibility = Visibility.Collapsed;
            LicensePage.Visibility = Visibility.Collapsed;
            FolderPage.Visibility = Visibility.Collapsed;
            InstallProgressPage.Visibility = Visibility.Collapsed;
            CompletePage.Visibility = Visibility.Collapsed;

            // Afficher la page courante
            switch (_currentPage)
            {
                case 1:
                    WelcomePage.Visibility = Visibility.Visible;
                    break;
                case 2:
                    LicensePage.Visibility = Visibility.Visible;
                    break;
                case 3:
                    FolderPage.Visibility = Visibility.Visible;
                    UpdateDiskSpace();
                    break;
                case 4:
                    InstallProgressPage.Visibility = Visibility.Visible;
                    StartSpinnerAnimation();
                    break;
                case 5:
                    CompletePage.Visibility = Visibility.Visible;
                    FinalInstallPath.Text = InstallPathTextBox.Text;
                    StopSpinnerAnimation();
                    break;
            }
        }

        private void UpdateNavigationButtons()
        {
            // Reset visibility
            BackButton.Visibility = Visibility.Collapsed;
            NextButton.Visibility = Visibility.Collapsed;
            InstallButton.Visibility = Visibility.Collapsed;
            FinishButton.Visibility = Visibility.Collapsed;
            CancelButton.Visibility = Visibility.Collapsed;

            switch (_currentPage)
            {
                case 1: // Bienvenue
                    NextButton.Visibility = Visibility.Visible;
                    break;

                case 2: // Licence
                    BackButton.Visibility = Visibility.Visible;
                    NextButton.Visibility = Visibility.Visible;
                    NextButton.IsEnabled = AcceptLicenseCheckBox.IsChecked == true;
                    break;

                case 3: // Dossier
                    BackButton.Visibility = Visibility.Visible;
                    InstallButton.Visibility = Visibility.Visible;
                    break;

                case 4: // Installation
                    CancelButton.Visibility = Visibility.Visible;
                    break;

                case 5: // Complete
                    FinishButton.Visibility = Visibility.Visible;
                    break;
            }
        }

        #endregion

        #region License

        private void LoadLicenseText()
        {
            LicenseText.Text = @"CONTRAT DE LICENCE UTILISATEUR FINAL
XNRGY Engineering Automation Tools
=====================================

Copyright (c) 2026 XNRGY Climate Systems ULC
Tous droits reserves.

IMPORTANT - VEUILLEZ LIRE ATTENTIVEMENT:

Ce Contrat de Licence Utilisateur Final (""CLUF"") est un accord juridique 
entre vous (une personne physique ou une entite unique) et XNRGY Climate 
Systems ULC pour le logiciel XNRGY Engineering Automation Tools.

EN INSTALLANT, COPIANT OU UTILISANT CE LOGICIEL, VOUS ACCEPTEZ D'ETRE 
LIE PAR LES TERMES DE CE CONTRAT.

1. OCTROI DE LICENCE
--------------------
XNRGY Climate Systems ULC vous accorde une licence non exclusive et non 
transferable pour utiliser ce logiciel uniquement pour vos operations 
internes au sein de XNRGY Climate Systems.

2. RESTRICTIONS
---------------
Vous n'etes pas autorise a:
- Copier ou distribuer ce logiciel en dehors de l'organisation XNRGY
- Modifier, adapter ou creer des oeuvres derivees du logiciel
- Desassembler, decompiler ou tenter de decouvrir le code source
- Sous-licencier, louer ou preter le logiciel

3. PROPRIETE INTELLECTUELLE
---------------------------
Le logiciel est protege par les lois sur le droit d'auteur et les traites 
internationaux. XNRGY Climate Systems ULC conserve tous les droits de 
propriete intellectuelle sur le logiciel.

4. GARANTIE LIMITEE
-------------------
CE LOGICIEL EST FOURNI ""TEL QUEL"" SANS GARANTIE D'AUCUNE SORTE.

5. LIMITATION DE RESPONSABILITE
-------------------------------
EN AUCUN CAS XNRGY CLIMATE SYSTEMS ULC NE SERA RESPONSABLE DE TOUT 
DOMMAGE DIRECT, INDIRECT, ACCESSOIRE, SPECIAL OU CONSECUTIF.

6. RESILIATION
--------------
Cette licence est effective jusqu'a sa resiliation. Elle se terminera 
automatiquement si vous ne respectez pas les termes de ce contrat.

7. DROIT APPLICABLE
-------------------
Ce contrat est regi par les lois du Quebec, Canada.

=====================================
En cliquant sur ""J'accepte"", vous confirmez avoir lu et compris ce 
contrat et acceptez d'etre lie par ses termes.

Pour toute question, contactez:
mohammedamine.elgalai@xnrgy.com
XNRGY Climate Systems ULC
Quebec, Canada";
        }

        private void AcceptLicense_Changed(object sender, RoutedEventArgs e)
        {
            if (NextButton != null)
            {
                NextButton.IsEnabled = AcceptLicenseCheckBox.IsChecked == true;
            }
        }

        #endregion

        #region Folder Selection

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Selectionnez le dossier d'installation";
                dialog.SelectedPath = InstallPathTextBox.Text;
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    InstallPathTextBox.Text = dialog.SelectedPath;
                    UpdateDiskSpace();
                }
            }
        }

        private void UpdateDiskSpace()
        {
            try
            {
                string path = InstallPathTextBox.Text;
                if (string.IsNullOrEmpty(path))
                {
                    DiskSpaceText.Text = "Espace disponible: --";
                    return;
                }

                string root = Path.GetPathRoot(path);
                if (string.IsNullOrEmpty(root))
                {
                    DiskSpaceText.Text = "Espace disponible: --";
                    return;
                }

                DriveInfo drive = new DriveInfo(root);
                long freeSpaceMB = drive.AvailableFreeSpace / (1024 * 1024);
                long freeSpaceGB = freeSpaceMB / 1024;

                if (freeSpaceGB >= 1)
                {
                    DiskSpaceText.Text = $"Espace disponible: {freeSpaceGB:N1} GB sur {root}";
                }
                else
                {
                    DiskSpaceText.Text = $"Espace disponible: {freeSpaceMB:N0} MB sur {root}";
                }
                
                DiskSpaceText.Foreground = freeSpaceMB > 200 
                    ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(144, 238, 144))
                    : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 99, 71));
            }
            catch
            {
                DiskSpaceText.Text = "Espace disponible: Calcul impossible";
            }
        }

        #endregion

        #region Installation Process

        private async Task StartInstallationAsync()
        {
            _isInstalling = true;
            InstallProgress.Value = 0;

            var options = new InstallationOptions
            {
                InstallPath = InstallPathTextBox.Text,
                CreateDesktopShortcut = CreateDesktopShortcut.IsChecked == true,
                CreateStartMenuShortcut = CreateStartMenuShortcut.IsChecked == true
            };

            var progress = new Progress<InstallationProgress>(p =>
            {
                InstallProgress.Value = p.Percentage;
                ProgressPercentText.Text = $"{p.Percentage}%";
                CurrentActionText.Text = p.CurrentAction;
            });

            try
            {
                bool success = await _installService.InstallAsync(options, progress);

                _isInstalling = false;

                if (success)
                {
                    _currentPage = 5; // Page complete
                    ShowCurrentPage();
                    UpdateNavigationButtons();
                }
                else
                {
                    System.Windows.MessageBox.Show(
                        "L'installation a echoue. Consultez les logs pour plus de details.",
                        "Erreur d'installation",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    
                    _currentPage = 3; // Retour a la page dossier
                    ShowCurrentPage();
                    UpdateNavigationButtons();
                }
            }
            catch (Exception ex)
            {
                _isInstalling = false;
                System.Windows.MessageBox.Show(
                    $"Erreur lors de l'installation:\n{ex.Message}",
                    "Erreur",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                
                _currentPage = 3;
                ShowCurrentPage();
                UpdateNavigationButtons();
            }
        }

        #endregion

        #region Spinner Animation

        private void InitializeSpinnerAnimation()
        {
            _spinnerTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(16) // ~60 FPS
            };
            _spinnerTimer.Tick += SpinnerTimer_Tick;
        }

        private double _spinnerAngle = 0;

        private void SpinnerTimer_Tick(object sender, EventArgs e)
        {
            _spinnerAngle += 5;
            if (_spinnerAngle >= 360)
                _spinnerAngle = 0;
            
            SpinnerRotation.Angle = _spinnerAngle;
        }

        private void StartSpinnerAnimation()
        {
            _spinnerTimer?.Start();
        }

        private void StopSpinnerAnimation()
        {
            _spinnerTimer?.Stop();
        }

        #endregion

        #region INotifyPropertyChanged

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}
