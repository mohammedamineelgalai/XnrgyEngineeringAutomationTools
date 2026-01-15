using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using XnrgyEngineeringAutomationTools.Modules.OpenVaultProject.Models;
using XnrgyEngineeringAutomationTools.Modules.OpenVaultProject.Services;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Modules.OpenVaultProject.Views
{
    /// <summary>
    /// Fenetre pour telecharger et ouvrir des projets depuis Vault
    /// </summary>
    public partial class OpenVaultProjectWindow : Window
    {
        private readonly VaultSdkService _vaultService;
        private readonly InventorService _inventorService;
        private VaultDownloadService? _downloadService;
        
        private VaultProjectItem? _selectedProject;
        private VaultProjectItem? _selectedReference;
        private VaultProjectItem? _selectedModule;

        public OpenVaultProjectWindow(VaultSdkService vaultService, InventorService inventorService)
        {
            InitializeComponent();
            _vaultService = vaultService ?? throw new ArgumentNullException(nameof(vaultService));
            _inventorService = inventorService ?? throw new ArgumentNullException(nameof(inventorService));
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                _downloadService = new VaultDownloadService(_vaultService, _inventorService);
                _downloadService.OnProgress += OnDownloadProgress;
                _downloadService.OnFileProgress += OnFileProgress;

                AddLog("[+] Fenetre initialisee", "SUCCESS");
                AddLog($"[i] Workspace: {_downloadService.GetLocalWorkspacePath()}", "INFO");
                
                LoadProjects();
            }
            catch (Exception ex)
            {
                AddLog($"[-] Erreur initialisation: {ex.Message}", "ERROR");
                Logger.LogException("OpenVaultProject.Window_Loaded", ex, Logger.LogLevel.ERROR);
            }
        }

        private void LoadProjects()
        {
            if (_downloadService == null) return;

            AddLog("[>] Chargement des projets...", "START");
            TxtStatus.Text = "Chargement des projets...";

            try
            {
                var projects = _downloadService.GetProjects();
                LstProjects.ItemsSource = projects;
                TxtProjectCount.Text = $" ({projects.Count})";
                
                // Reset les autres colonnes
                LstReferences.ItemsSource = null;
                LstModules.ItemsSource = null;
                TxtRefCount.Text = " (0)";
                TxtModuleCount.Text = " (0)";
                BtnDownload.IsEnabled = false;

                AddLog($"[+] {projects.Count} projets charges", "SUCCESS");
                TxtStatus.Text = $"{projects.Count} projets disponibles - Selectionnez un projet";
            }
            catch (Exception ex)
            {
                AddLog($"[-] Erreur chargement projets: {ex.Message}", "ERROR");
                TxtStatus.Text = "Erreur lors du chargement";
            }
        }

        private void LstProjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _selectedProject = LstProjects.SelectedItem as VaultProjectItem;
            _selectedReference = null;
            _selectedModule = null;

            LstReferences.ItemsSource = null;
            LstModules.ItemsSource = null;
            TxtRefCount.Text = " (0)";
            TxtModuleCount.Text = " (0)";
            BtnDownload.IsEnabled = false;

            if (_selectedProject != null && _downloadService != null)
            {
                AddLog($"[>] Chargement des references pour {_selectedProject.Name}...", "INFO");
                TxtStatus.Text = $"Chargement des references pour {_selectedProject.Name}...";

                var references = _downloadService.GetReferences(_selectedProject.Path);
                LstReferences.ItemsSource = references;
                TxtRefCount.Text = $" ({references.Count})";

                AddLog($"[+] {references.Count} references trouvees", "SUCCESS");
                TxtStatus.Text = $"Projet {_selectedProject.Name} - {references.Count} references";
            }
        }

        private void LstReferences_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _selectedReference = LstReferences.SelectedItem as VaultProjectItem;
            _selectedModule = null;

            LstModules.ItemsSource = null;
            TxtModuleCount.Text = " (0)";
            BtnDownload.IsEnabled = false;

            if (_selectedReference != null && _downloadService != null)
            {
                AddLog($"[>] Chargement des modules pour {_selectedReference.Name}...", "INFO");
                TxtStatus.Text = $"Chargement des modules pour {_selectedReference.Name}...";

                var modules = _downloadService.GetModules(_selectedReference.Path);
                LstModules.ItemsSource = modules;
                TxtModuleCount.Text = $" ({modules.Count})";

                AddLog($"[+] {modules.Count} modules trouves", "SUCCESS");
                TxtStatus.Text = $"{_selectedProject?.Name}/{_selectedReference.Name} - {modules.Count} modules";
            }
        }

        private void LstModules_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _selectedModule = LstModules.SelectedItem as VaultProjectItem;
            BtnDownload.IsEnabled = _selectedModule != null;

            if (_selectedModule != null)
            {
                TxtStatus.Text = $"Module selectionne: {_selectedProject?.Name}/{_selectedReference?.Name}/{_selectedModule.Name}";
                AddLog($"[i] Module selectionne: {_selectedModule.Path}", "INFO");
            }
        }

        private async void BtnDownload_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedModule == null || _downloadService == null)
            {
                AddLog("[-] Aucun module selectionne", "ERROR");
                return;
            }

            // Desactiver les controles pendant le telechargement
            BtnDownload.IsEnabled = false;
            BtnRefresh.IsEnabled = false;
            LstProjects.IsEnabled = false;
            LstReferences.IsEnabled = false;
            LstModules.IsEnabled = false;

            try
            {
                AddLog($"[>] Debut du telechargement: {_selectedModule.Path}", "START");
                TxtStatus.Text = "Telechargement en cours...";

                bool success = await _downloadService.DownloadAndOpenModuleAsync(_selectedModule);

                if (success)
                {
                    AddLog("[+] Module telecharge et ouvert avec succes", "SUCCESS");
                    TxtStatus.Text = "Module ouvert dans Inventor";
                }
                else
                {
                    AddLog("[!] Telechargement termine avec des avertissements", "WARN");
                    TxtStatus.Text = "Telechargement termine (verifier le journal)";
                }
            }
            catch (Exception ex)
            {
                AddLog($"[-] Erreur: {ex.Message}", "ERROR");
                TxtStatus.Text = "Erreur lors du telechargement";
                Logger.LogException("OpenVaultProject.BtnDownload_Click", ex, Logger.LogLevel.ERROR);
            }
            finally
            {
                // Reactiver les controles
                BtnDownload.IsEnabled = true;
                BtnRefresh.IsEnabled = true;
                LstProjects.IsEnabled = true;
                LstReferences.IsEnabled = true;
                LstModules.IsEnabled = true;
                TxtFileProgress.Text = "";
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadProjects();
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            LstLog.Items.Clear();
            AddLog("[i] Journal efface", "INFO");
        }

        private void OnDownloadProgress(string message, string level)
        {
            Dispatcher.Invoke(() =>
            {
                AddLog(message, level);
                TxtStatus.Text = message;
            });
        }

        private void OnFileProgress(int current, int total)
        {
            Dispatcher.Invoke(() =>
            {
                TxtFileProgress.Text = $"{current}/{total} fichiers";
            });
        }

        private void AddLog(string message, string level)
        {
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                string text = $"[{timestamp}] {message}";

                var textBlock = new TextBlock
                {
                    Text = text,
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = 12,
                    Padding = new Thickness(8, 2, 8, 2),
                    TextWrapping = TextWrapping.Wrap
                };

                switch (level)
                {
                    case "ERROR":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(239, 68, 68)); // Rouge
                        textBlock.FontWeight = FontWeights.Bold;
                        break;
                    case "WARN":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(251, 191, 36)); // Orange
                        break;
                    case "SUCCESS":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(34, 197, 94)); // Vert
                        break;
                    case "START":
                        textBlock.Foreground = new SolidColorBrush(Color.FromRgb(59, 130, 246)); // Bleu
                        textBlock.FontWeight = FontWeights.SemiBold;
                        break;
                    default:
                        textBlock.Foreground = new SolidColorBrush(Colors.White);
                        break;
                }

                LstLog.Items.Add(textBlock);
                LstLog.ScrollIntoView(textBlock);

                // Limiter le nombre de lignes
                while (LstLog.Items.Count > 100)
                {
                    LstLog.Items.RemoveAt(0);
                }
            });
        }
    }
}
