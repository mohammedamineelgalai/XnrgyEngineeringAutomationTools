using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Shared.Views;
using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;

namespace XnrgyEngineeringAutomationTools.Modules.UploadTemplate.Views
{
    /// <summary>
    /// Fichier a uploader - avec selection et statut
    /// </summary>
    public class TemplateFileItem : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        private string _status = "En attente";
        private string _vaultPath = string.Empty;

        public string FileName { get; set; } = string.Empty;
        public string FileExtension { get; set; } = string.Empty;
        public string FileSizeFormatted { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string RelativePath { get; set; } = string.Empty;

        /// <summary>
        /// Chemin de destination dans Vault (ex: $/Engineering/Library/Template/SubFolder)
        /// </summary>
        public string VaultPath
        {
            get => _vaultPath;
            set { _vaultPath = value; OnPropertyChanged(); }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName ?? string.Empty));
        }
    }

    /// <summary>
    /// Upload Template vers Vault - Upload massif de fichiers
    /// Auteur: Mohammed Amine Elgalai - XNRGY Climate Systems
    /// Utilise la connexion partagee de MainWindow via VaultSdkService
    /// </summary>
    public partial class UploadTemplateWindow : Window
    {
        // ====================================================================
        // Services et connexion (partagee depuis MainWindow)
        // ====================================================================
        private readonly VaultSdkService? _vaultService;
        private VDF.Vault.Currency.Connections.Connection? _connection;
        private bool _isConnected = false;

        // ====================================================================
        // Upload
        // ====================================================================
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isUploading = false;
        private bool _isPaused = false;
        private ManualResetEventSlim _pauseEvent = new ManualResetEventSlim(true);
        private Dictionary<string, ACW.Folder> _folderCache = new Dictionary<string, ACW.Folder>();
        private string _uploadComment = string.Empty; // Commentaire pour l'upload

        // ====================================================================
        // DataGrid - Collection de fichiers
        // ====================================================================
        public ObservableCollection<TemplateFileItem> AllFiles { get; } = new();
        private List<TemplateFileItem> _allFilesList = new();
        private string _basePath = string.Empty;

        // ====================================================================
        // Statistiques
        // ====================================================================
        private int _uploadedFiles = 0;
        private int _skippedFiles = 0;
        private int _failedFiles = 0;
        private int _foldersCreated = 0;
        private DateTime _startTime;
        private TimeSpan _pausedTime = TimeSpan.Zero;
        private DateTime _pauseStartTime;
        private bool _createFoldersAutomatically = true;

        // ====================================================================
        // Exclusions
        // ====================================================================
        private static readonly string[] ExcludedExtensions = { ".bak", ".old", ".tmp", ".log", ".lck", ".dwl", ".dwl2", ".v" };
        private static readonly string[] ExcludedPrefixes = { "~$", "._", "Backup_", ".~" };
        private static readonly string[] ExcludedFolders = { "OldVersions", "Backup", ".vault", ".git", ".vs", "obj", "bin", "Workspace", "vltcache" };
        private static readonly string[] ExcludedFileNames = { "desktop.ini", "Thumbs.db", ".DS_Store" };

        // ====================================================================
        // Data Binding pour extensions (garde pour compatibilite)
        // ====================================================================
        public ObservableCollection<ExtensionInfo> ExtensionStats { get; set; } = new ObservableCollection<ExtensionInfo>();

        /// <summary>
        /// Constructeur avec service Vault partage
        /// </summary>
        public UploadTemplateWindow(VaultSdkService? vaultService)
        {
            InitializeComponent();
            
            // Bind DataGrid
            DgFiles.ItemsSource = AllFiles;
            
            _vaultService = vaultService;
            
            // Recuperer la connexion du service partage
            if (_vaultService != null && _vaultService.IsConnected)
            {
                _connection = _vaultService.Connection;
                _isConnected = _connection != null;
            }
            
            // S'abonner aux changements de theme
            MainWindow.ThemeChanged += OnThemeChanged;
            this.Closed += (s, e) => MainWindow.ThemeChanged -= OnThemeChanged;
            
            // Appliquer le theme actuel au demarrage
            ApplyTheme(MainWindow.CurrentThemeIsDark);
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
                MainGrid.Background = new SolidColorBrush(Color.FromRgb(30, 30, 46)); // #1E1E2E
            }
            else
            {
                // Theme CLAIR
                MainGrid.Background = new SolidColorBrush(Color.FromRgb(245, 247, 250)); // Bleu-gris tres clair
            }
        }

        // ====================================================================
        // Chargement fenetre
        // ====================================================================
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // [+] Initialiser _startTime pour eviter temps ecoule enorme au demarrage
            _startTime = DateTime.Now;
            _pausedTime = TimeSpan.Zero;
            TxtProgressTimeElapsed.Text = "00:00";
            TxtProgressTimeEstimated.Text = "00:00";
            
            // Note: Logo remplace par emoji dans le header moderne

            // Afficher le statut de connexion (format identique MainWindow)
            if (_isConnected)
            {
                RunVaultName.Text = $" Vault: {_vaultService?.VaultName}";
                RunUserName.Text = $" {_vaultService?.UserName}";
                RunStatus.Text = " Connecte";
                StatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#107C10"));
                Log($"[+] Connexion Vault heritee de l'application principale", LogLevel.SUCCESS);
                Log($"[i] Vault: {_vaultService?.VaultName} | User: {_vaultService?.UserName}", LogLevel.INFO);
                BtnStartUpload.IsEnabled = false; // Active apres scan
            }
            else
            {
                RunVaultName.Text = " Vault: --";
                RunUserName.Text = " --";
                RunStatus.Text = " Deconnecte";
                StatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E81123"));
                Log("[-] Aucune connexion Vault. Veuillez vous connecter via l'application principale.", LogLevel.ERROR);
                BtnStartUpload.IsEnabled = false;
            }
            
            Log("[i] Fenetre Upload Template initialisee", LogLevel.INFO);
            
            // Scan automatique au demarrage si le chemin existe
            string localPath = TxtLocalPath.Text.Trim();
            if (Directory.Exists(localPath))
            {
                Log("[>] Scan automatique au demarrage...", LogLevel.INFO);
                await Task.Delay(100); // Petit delai pour laisser l'UI se charger
                BtnScan_Click(sender, e);
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (_isUploading)
            {
                var result = XnrgyMessageBox.Show(
                    "Un upload est en cours. Voulez-vous vraiment annuler et fermer?",
                    "Upload en cours",
                    XnrgyMessageBoxType.Warning,
                    XnrgyMessageBoxButtons.YesNo,
                    this);

                if (result != XnrgyMessageBoxResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }

                _cancellationTokenSource?.Cancel();
            }
            // Note: Pas de deconnexion - connexion geree par MainWindow
        }

        // ====================================================================
        // Parcourir dossier
        // ====================================================================
        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Selectionnez le dossier a uploader";
                dialog.SelectedPath = TxtLocalPath.Text;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtLocalPath.Text = dialog.SelectedPath;
                    UpdateVaultPath();
                }
            }
        }

        private void UpdateVaultPath()
        {
            string localPath = TxtLocalPath.Text.Trim();
            if (localPath.StartsWith(@"C:\Vault", StringComparison.OrdinalIgnoreCase))
            {
                string relative = localPath.Substring(8).TrimStart('\\');
                TxtVaultPath.Text = string.IsNullOrEmpty(relative) ? "-> $/" : $"-> $/{relative.Replace("\\", "/")}";
            }
            else
            {
                TxtVaultPath.Text = "-> $/";
            }
        }

        // ====================================================================
        // Scanner les fichiers
        // ====================================================================
        private async void BtnScan_Click(object sender, RoutedEventArgs e)
        {
            string localPath = TxtLocalPath.Text.Trim();

            if (!Directory.Exists(localPath))
            {
                XnrgyMessageBox.ShowError($"Le dossier n'existe pas:\n{localPath}", "Erreur", this);
                return;
            }

            BtnScan.IsEnabled = false;
            _basePath = localPath;
            Log($"[?] Scan des fichiers dans {localPath}...", LogLevel.INFO);
            UpdateProgress("Scan en cours...", 0);

            // Capturer les options AVANT le Task.Run (thread UI)
            bool excludeOldVersions = ChkExcludeOldVersions.IsChecked == true;
            bool excludeTempFiles = ChkExcludeTempFiles.IsChecked == true;
            string baseVaultPath = TxtVaultPath.Text.Replace("-> ", "").TrimEnd('/');

            try
            {
                // Desabonner les anciens items avant de vider
                foreach (var item in _allFilesList)
                {
                    item.PropertyChanged -= FileItem_PropertyChanged;
                }
                AllFiles.Clear();
                _allFilesList.Clear();
                ExtensionStats.Clear();

                var fileItems = new List<TemplateFileItem>();

                await Task.Run(() =>
                {
                    var allFiles = Directory.EnumerateFiles(localPath, "*.*", SearchOption.AllDirectories)
                        .Where(f => !ShouldExcludeFile(f, excludeOldVersions, excludeTempFiles))
                        .ToList();

                    foreach (var filePath in allFiles)
                    {
                        var fileInfo = new FileInfo(filePath);
                        var relativePath = filePath.Substring(localPath.Length).TrimStart('\\');
                        var relativeDir = Path.GetDirectoryName(relativePath) ?? "";
                        
                        // Calculer le chemin Vault de destination pour ce fichier
                        string fileVaultPath = string.IsNullOrEmpty(relativeDir) 
                            ? baseVaultPath 
                            : $"{baseVaultPath}/{relativeDir.Replace("\\", "/")}";

                        fileItems.Add(new TemplateFileItem
                        {
                            FileName = fileInfo.Name,
                            FileExtension = fileInfo.Extension.ToLower(),
                            FileSizeFormatted = FormatFileSize(fileInfo.Length),
                            FullPath = filePath,
                            RelativePath = relativeDir,
                            VaultPath = fileVaultPath,
                            IsSelected = true,
                            Status = "En attente"
                        });
                    }
                    
                    // Trier par ordre alphabetique (nom de fichier)
                    fileItems = fileItems.OrderBy(f => f.FileName, StringComparer.OrdinalIgnoreCase).ToList();
                });

                // Ajouter au DataGrid sur le thread UI
                foreach (var item in fileItems)
                {
                    // S'abonner au changement de IsSelected pour mettre a jour les statistiques en temps reel
                    item.PropertyChanged += FileItem_PropertyChanged;
                    AllFiles.Add(item);
                    _allFilesList.Add(item);
                }

                // Mise a jour statistiques
                UpdateStatistics();
                
                // Peupler le filtre d'extension dynamiquement
                PopulateExtensionFilter();

                Log($"[+] Scan termine: {AllFiles.Count} fichiers trouves", LogLevel.SUCCESS);
                UpdateProgress($"{AllFiles.Count} fichiers prets a uploader", 0);

                BtnStartUpload.IsEnabled = _isConnected && AllFiles.Count > 0;
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur scan: {ex.Message}", LogLevel.ERROR);
                XnrgyMessageBox.ShowError($"Erreur lors du scan:\n{ex.Message}", "Erreur", this);
            }
            finally
            {
                BtnScan.IsEnabled = true;
            }
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double size = bytes;
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }
            return $"{size:0.##} {sizes[order]}";
        }

        private void UpdateStatistics()
        {
            int total = _allFilesList.Count;
            int selected = _allFilesList.Count(f => f.IsSelected);
            
            TxtTotalFiles.Text = total.ToString("N0");
            TxtSelectedFiles.Text = selected.ToString("N0");
            TxtUploaded.Text = _uploadedFiles.ToString("N0");
            TxtSkipped.Text = _skippedFiles.ToString("N0");
            TxtFailed.Text = _failedFiles.ToString("N0");
        }

        /// <summary>
        /// Gestionnaire pour mettre a jour les statistiques quand IsSelected change
        /// </summary>
        private void FileItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsSelected")
            {
                // Mise a jour sur le thread UI
                Dispatcher.BeginInvoke(new Action(() => UpdateStatistics()));
            }
        }

        /// <summary>
        /// Verifie si un fichier doit etre exclu (thread-safe)
        /// </summary>
        private bool ShouldExcludeFile(string filePath, bool excludeOldVersions, bool excludeTempFiles)
        {
            string fileName = Path.GetFileName(filePath);

            // Noms exclus
            if (ExcludedFileNames.Any(n => fileName.Equals(n, StringComparison.OrdinalIgnoreCase)))
                return true;

            // Extensions exclues (si option activee)
            if (excludeTempFiles)
            {
                if (ExcludedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                    return true;
            }

            // Prefixes exclus
            if (ExcludedPrefixes.Any(p => fileName.StartsWith(p)))
                return true;

            // Dossiers exclus (si option cochee)
            if (excludeOldVersions)
            {
                if (ExcludedFolders.Any(f => filePath.Contains($"\\{f}\\") || filePath.EndsWith($"\\{f}")))
                    return true;
            }

            return false;
        }

        // ====================================================================
        // Demarrer l'upload
        // ====================================================================
        private async void BtnStartUpload_Click(object sender, RoutedEventArgs e)
        {
            var selectedFiles = _allFilesList.Where(f => f.IsSelected).ToList();
            
            if (!_isConnected || selectedFiles.Count == 0)
            {
                XnrgyMessageBox.ShowError("Veuillez vous connecter et selectionner des fichiers d'abord.", "Erreur", this);
                return;
            }

            // Confirmation
            var confirm = XnrgyMessageBox.Show(
                $"Vous allez uploader {selectedFiles.Count} fichiers vers le Vault PRODUCTION.\n\n" +
                $"Source: {TxtLocalPath.Text}\n" +
                $"Destination: {TxtVaultPath.Text.Replace("-> ", "")}\n\n" +
                "Continuer?",
                "Confirmer Upload",
                XnrgyMessageBoxType.Warning,
                XnrgyMessageBoxButtons.YesNo,
                this);

            if (confirm != XnrgyMessageBoxResult.Yes)
                return;

            // Initialiser
            _isUploading = true;
            _isPaused = false;
            _pauseEvent.Set(); // S'assurer qu'on n'est pas en pause
            _cancellationTokenSource = new CancellationTokenSource();
            _uploadedFiles = 0;
            _skippedFiles = 0;
            _failedFiles = 0;
            _foldersCreated = 0;
            _folderCache.Clear();
            
            // [+] RESET du timer au début de l'opération
            _startTime = DateTime.Now;
            _pausedTime = TimeSpan.Zero;
            
            // [+] Reinitialiser les affichages de temps a zero
            TxtProgressTimeElapsed.Text = "00:00";
            TxtProgressTimeEstimated.Text = "00:00";
            TxtCurrentFile.Text = "";
            TxtProgressPercent.Text = "";
            ProgressBarFill.Width = 0;
            
            _createFoldersAutomatically = ChkCreateFolders.IsChecked == true;
            _uploadComment = TxtComment.Text.Trim(); // Recuperer le commentaire

            // UI - Activer boutons de controle
            BtnStartUpload.IsEnabled = false;
            BtnPause.IsEnabled = true;
            BtnStop.IsEnabled = true;
            BtnCancel.IsEnabled = false;
            BtnScan.IsEnabled = false;
            BtnBrowse.IsEnabled = false;
            BtnPause.Content = "[||] PAUSE";

            int totalSelected = selectedFiles.Count;
            
            Log("[>] Debut de l'upload...", LogLevel.INFO);
            Log($"[i] {totalSelected} fichiers selectionnes pour upload", LogLevel.INFO);
            Log($"[i] Destination Vault: {TxtVaultPath.Text.Replace("-> ", "")}", LogLevel.INFO);
            Log($"[i] Commentaire: {_uploadComment}", LogLevel.INFO);

            try
            {
                string localRoot = TxtLocalPath.Text.Trim();
                var token = _cancellationTokenSource.Token;
                int counter = 0;

                foreach (var fileItem in selectedFiles)
                {
                    // Verification multiple du token pour arret immediat
                    if (token.IsCancellationRequested || !_isUploading)
                    {
                        Log("[!] Upload annule par l'utilisateur", LogLevel.WARNING);
                        break;
                    }

                    // Attendre si en pause
                    _pauseEvent.Wait(token);
                    
                    // Re-verifier apres la pause
                    if (token.IsCancellationRequested || !_isUploading)
                    {
                        Log("[!] Upload annule par l'utilisateur", LogLevel.WARNING);
                        break;
                    }

                    counter++;
                    string filePath = fileItem.FullPath;
                    string fileName = fileItem.FileName;
                    string vaultFolder = GetVaultFolderPath(filePath, localRoot);

                    // Mise a jour statut dans DataGrid (AVANT upload)
                    Dispatcher.Invoke(() => { fileItem.Status = "Upload en cours..."; });

                    // Mise a jour progression
                    double progress = (counter * 100.0) / totalSelected;
                    var elapsed = DateTime.Now - _startTime;
                    double rate = counter / Math.Max(elapsed.TotalMinutes, 0.01);
                    double remaining = (totalSelected - counter) / Math.Max(rate, 1);

                    UpdateProgress(
                        $"{counter}/{totalSelected} fichiers",
                        progress,
                        "",
                        "",
                        fileName);

                    // Sauvegarder les compteurs avant l'upload pour detecter les changements
                    int uploadedBefore = _uploadedFiles;
                    int skippedBefore = _skippedFiles;
                    int failedBefore = _failedFiles;

                    // Upload - verifier le flag _isUploading aussi
                    bool success = false;
                    if (_isUploading && !token.IsCancellationRequested)
                    {
                        success = await Task.Run(() => UploadFile(filePath, vaultFolder), token);
                    }
                    else
                    {
                        break;
                    }

                    // Mise a jour statut (APRES upload) - determiner le statut selon les compteurs
                    Dispatcher.Invoke(() =>
                    {
                        if (_uploadedFiles > uploadedBefore)
                            fileItem.Status = "✅ Uploade";
                        else if (_skippedFiles > skippedBefore)
                            fileItem.Status = "⏸️ Ignore (existe)";
                        else if (_failedFiles > failedBefore)
                            fileItem.Status = "❌ Echec";
                        else
                            fileItem.Status = "❓ Inconnu";
                            
                        TxtUploaded.Text = _uploadedFiles.ToString("N0");
                        TxtSkipped.Text = _skippedFiles.ToString("N0");
                        TxtFailed.Text = _failedFiles.ToString("N0");
                        UpdateStatistics();
                    });
                }

                // Resume
                var totalTime = DateTime.Now - _startTime;
                Log($"[+] Upload termine en {totalTime.TotalMinutes:F1} minutes", LogLevel.SUCCESS);
                Log($"[i] Uploades: {_uploadedFiles} | Ignores: {_skippedFiles} | Echecs: {_failedFiles}", LogLevel.INFO);

                UpdateProgress($"Termine! {_uploadedFiles} fichiers uploades", 100);

                if (_failedFiles == 0)
                {
                    XnrgyMessageBox.ShowSuccess(
                        $"Upload termine avec succes!\n\n" +
                        $"Fichiers uploades: {_uploadedFiles}\n" +
                        $"Fichiers ignores: {_skippedFiles}\n" +
                        $"Dossiers crees: {_foldersCreated}\n" +
                        $"Temps total: {totalTime.TotalMinutes:F1} minutes",
                        "Upload termine", this);
                }
                else
                {
                    XnrgyMessageBox.Show(
                        $"Upload termine avec {_failedFiles} erreurs.\n\n" +
                        $"Fichiers uploades: {_uploadedFiles}\n" +
                        $"Fichiers ignores: {_skippedFiles}\n" +
                        $"Echecs: {_failedFiles}\n" +
                        $"Consultez le journal pour les details.",
                        "Upload termine",
                        XnrgyMessageBoxType.Warning,
                        XnrgyMessageBoxButtons.OK, this);
                }
            }
            catch (OperationCanceledException)
            {
                Log("[!] Upload annule", LogLevel.WARNING);
                UpdateProgress("Annule", 0);
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur: {ex.Message}", LogLevel.ERROR);
                XnrgyMessageBox.ShowError($"Erreur lors de l'upload:\n{ex.Message}", "Erreur", this);
            }
            finally
            {
                _isUploading = false;
                _isPaused = false;
                _pauseEvent.Set();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;

                // Restaurer UI
                BtnStartUpload.IsEnabled = _isConnected && _allFilesList.Count > 0;
                BtnPause.IsEnabled = false;
                BtnStop.IsEnabled = false;
                BtnCancel.IsEnabled = true;
                BtnScan.IsEnabled = true;
                BtnBrowse.IsEnabled = true;
                BtnPause.Content = "⏸️ PAUSE";
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource?.Cancel();
            BtnCancel.IsEnabled = false;
        }

        private void BtnPause_Click(object sender, RoutedEventArgs e)
        {
            if (_isPaused)
            {
                // Reprendre
                _isPaused = false;
                if (_pauseStartTime != default(DateTime))
                {
                    _pausedTime += DateTime.Now - _pauseStartTime;
                }
                _pauseEvent.Set();
                BtnPause.Content = "⏸️ PAUSE";
                Log("[>] Upload repris", LogLevel.INFO);
                UpdateProgress("Upload en cours...", -1); // -1 = garder le pourcentage actuel
            }
            else
            {
                // Mettre en pause
                _isPaused = true;
                _pauseEvent.Reset();
                BtnPause.Content = "▶️ REPRENDRE";
                Log("[!] Upload en pause", LogLevel.WARNING);
                UpdateProgress("En pause...", -1);
            }
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            if (_isUploading)
            {
                var confirm = XnrgyMessageBox.Show(
                    "Voulez-vous vraiment arreter l'upload?\nLes fichiers deja uploades seront conserves.",
                    "Arreter l'upload",
                    XnrgyMessageBoxType.Warning,
                    XnrgyMessageBoxButtons.YesNo,
                    this);

                if (confirm == XnrgyMessageBoxResult.Yes)
                {
                    // Forcer l'arret immediat
                    _isUploading = false;
                    _cancellationTokenSource?.Cancel();
                    _pauseEvent.Set(); // Debloquer si en pause
                    
                    // Restaurer UI immediatement
                    BtnStartUpload.IsEnabled = _isConnected && _allFilesList.Count > 0;
                    BtnPause.IsEnabled = false;
                    BtnStop.IsEnabled = false;
                    BtnCancel.IsEnabled = true;
                    BtnScan.IsEnabled = true;
                    BtnBrowse.IsEnabled = true;
                    BtnPause.Content = "[||] PAUSE";
                    
                    Log("[!] Upload arrete par l'utilisateur", LogLevel.WARNING);
                    UpdateProgress("Arrete par l'utilisateur", 0);
                    UpdateStatistics();
                }
            }
        }

        // ====================================================================
        // Upload d'un fichier
        // ====================================================================
        private bool UploadFile(string localFilePath, string vaultFolderPath)
        {
            if (_connection == null) return false;

            string fileName = Path.GetFileName(localFilePath);

            try
            {
                // Obtenir ou creer le dossier Vault
                var folder = EnsureVaultFolder(vaultFolderPath);
                if (folder == null)
                {
                    Interlocked.Increment(ref _failedFiles);
                    Log($"[-] Dossier introuvable: {vaultFolderPath}", LogLevel.ERROR);
                    return false;
                }

                // Verifier si le fichier existe deja
                bool skipExisting = false;
                Dispatcher.Invoke(() => { skipExisting = ChkSkipExisting.IsChecked == true; });
                
                if (skipExisting)
                {
                    try
                    {
                        var existingFiles = _connection.WebServiceManager.DocumentService
                            .GetLatestFilesByFolderId(folder.Id, false);

                        if (existingFiles != null && existingFiles.Any(f => f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                        {
                            Interlocked.Increment(ref _skippedFiles);
                            Log($"[~] {fileName}: Fichier existe deja - ignore", LogLevel.WARNING);
                            return true;
                        }
                    }
                    catch { }
                }

                // Upload
                var fileInfo = new FileInfo(localFilePath);

                using (var stream = File.OpenRead(localFilePath))
                {
                    var vdfFolder = new VDF.Vault.Currency.Entities.Folder(_connection, folder);

                    var result = _connection.FileManager.AddFile(
                        vdfFolder,
                        fileName,
                        _uploadComment, // Utiliser le commentaire du TextBox
                        fileInfo.LastWriteTimeUtc,
                        null,
                        null,
                        ACW.FileClassification.None,
                        false,
                        stream
                    );

                    if (result != null)
                    {
                        Interlocked.Increment(ref _uploadedFiles);
                        Log($"[+] {fileName}: Upload reussi", LogLevel.SUCCESS);
                        return true;
                    }
                }

                Interlocked.Increment(ref _failedFiles);
                Log($"[-] {fileName}: Echec de l'upload (resultat null)", LogLevel.ERROR);
                return false;
            }
            catch (Exception ex)
            {
                // Fichier existant? (traduire le message)
                if (ex.Message.Contains("already exists") || ex.Message.Contains("1008"))
                {
                    Interlocked.Increment(ref _skippedFiles);
                    Log($"[~] {fileName}: Fichier existe deja dans Vault - ignore", LogLevel.WARNING);
                    return true;
                }

                // Traduire les erreurs courantes en francais
                string errorMsg = TranslateVaultError(ex.Message);
                
                Interlocked.Increment(ref _failedFiles);
                Log($"[-] {fileName}: {errorMsg}", LogLevel.ERROR);
                return false;
            }
        }
        
        /// <summary>
        /// Traduit les messages d'erreur Vault SDK en francais
        /// </summary>
        private string TranslateVaultError(string englishMessage)
        {
            if (englishMessage.Contains("The calling thread cannot access this object"))
                return "Erreur de thread - acces interdit depuis ce thread";
            if (englishMessage.Contains("already exists"))
                return "Le fichier existe deja";
            if (englishMessage.Contains("1008"))
                return "Fichier deja present dans Vault (erreur 1008)";
            if (englishMessage.Contains("1003"))
                return "Fichier en cours de traitement par Job Processor (erreur 1003)";
            if (englishMessage.Contains("1013"))
                return "Fichier doit etre checke out d'abord (erreur 1013)";
            if (englishMessage.Contains("Access denied") || englishMessage.Contains("permission"))
                return "Acces refuse - verifiez vos permissions";
            if (englishMessage.Contains("connection") || englishMessage.Contains("network"))
                return "Erreur de connexion reseau";
            if (englishMessage.Contains("timeout"))
                return "Delai d'attente depasse (timeout)";
                
            // Si pas de traduction, retourner le message original
            return englishMessage;
        }

        private ACW.Folder EnsureVaultFolder(string vaultPath)
        {
            if (_connection == null) return null;

            // Cache
            if (_folderCache.TryGetValue(vaultPath, out var cachedFolder))
                return cachedFolder;

            try
            {
                var folder = _connection.WebServiceManager.DocumentService.GetFolderByPath(vaultPath);
                _folderCache[vaultPath] = folder;
                return folder;
            }
            catch
            {
                // Creer le dossier automatiquement (si option activee)
                if (!_createFoldersAutomatically)
                    return null;
                    
                try
                {
                    int lastSlash = vaultPath.LastIndexOf('/');
                    if (lastSlash <= 0)
                    {
                        var rootFolder = _connection.WebServiceManager.DocumentService.GetFolderByPath("$");
                        _folderCache["$"] = rootFolder;
                        return rootFolder;
                    }

                    string parentPath = vaultPath.Substring(0, lastSlash);
                    string folderName = vaultPath.Substring(lastSlash + 1);

                    var parentFolder = EnsureVaultFolder(parentPath);
                    if (parentFolder == null) return null;

                    var newFolder = _connection.WebServiceManager.DocumentService.AddFolder(folderName, parentFolder.Id, false);
                    Interlocked.Increment(ref _foldersCreated);
                    _folderCache[vaultPath] = newFolder;
                    return newFolder;
                }
                catch
                {
                    // Peut-etre deja cree
                    try
                    {
                        var folder = _connection.WebServiceManager.DocumentService.GetFolderByPath(vaultPath);
                        _folderCache[vaultPath] = folder;
                        return folder;
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
        }

        private string GetVaultFolderPath(string localFilePath, string localRoot)
        {
            string fileDir = Path.GetDirectoryName(localFilePath) ?? "";
            const string vaultRoot = @"C:\Vault";

            if (fileDir.StartsWith(vaultRoot, StringComparison.OrdinalIgnoreCase))
            {
                string relative = fileDir.Substring(vaultRoot.Length).TrimStart('\\');
                return string.IsNullOrEmpty(relative) ? "$" : "$/" + relative.Replace("\\", "/");
            }

            if (fileDir.StartsWith(localRoot, StringComparison.OrdinalIgnoreCase))
            {
                string localRootRelative = localRoot.StartsWith(vaultRoot, StringComparison.OrdinalIgnoreCase)
                    ? localRoot.Substring(vaultRoot.Length).TrimStart('\\')
                    : "";

                string vaultBase = string.IsNullOrEmpty(localRootRelative) ? "$" : "$/" + localRootRelative.Replace("\\", "/");
                string relative = fileDir.Substring(localRoot.Length).TrimStart('\\');

                return string.IsNullOrEmpty(relative) ? vaultBase : vaultBase + "/" + relative.Replace("\\", "/");
            }

            return "$";
        }

        // ====================================================================
        // Journal et UI
        // ====================================================================
        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            TxtLog.Document.Blocks.Clear();
        }

        private void Log(string message, LogLevel level)
        {
            // Ecrire aussi dans le fichier log principal de l'application
            var logLevel = level switch
            {
                LogLevel.SUCCESS => Logger.LogLevel.INFO,
                LogLevel.ERROR => Logger.LogLevel.ERROR,
                LogLevel.WARNING => Logger.LogLevel.WARNING,
                _ => Logger.LogLevel.INFO
            };
            Logger.Log($"[UploadTemplate] {message}", logLevel);
            
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                
                // Utilise JournalColorService pour les couleurs uniformisees
                var serviceLevel = level switch
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
                var paragraph = new System.Windows.Documents.Paragraph();
                paragraph.Margin = new Thickness(0, 2, 0, 2);
                
                // Timestamp en gris (depuis le service)
                var timestampRun = new System.Windows.Documents.Run($"[{timestamp}] ")
                {
                    Foreground = JournalColorService.TimestampBrush
                };
                paragraph.Inlines.Add(timestampRun);
                
                // Message avec couleur
                var messageRun = new System.Windows.Documents.Run($"{prefix} {message}")
                {
                    Foreground = new SolidColorBrush(color)
                };
                paragraph.Inlines.Add(messageRun);
                
                TxtLog.Document.Blocks.Add(paragraph);
                TxtLog.ScrollToEnd();
            });
        }

        private void UpdateStatus(string text, string color)
        {
            Dispatcher.Invoke(() =>
            {
                StatusText.Text = text;
                StatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));
            });
        }

        private void UpdateProgress(string status, double percent, string speed = "", string remaining = "", string currentFile = "")
        {
            Dispatcher.Invoke(() =>
            {
                TxtProgressStatus.Text = status;
                
                // Afficher le fichier actuellement en traitement
                if (!string.IsNullOrEmpty(currentFile))
                {
                    TxtCurrentFile.Text = currentFile;
                }
                
                // Calcul du temps écoulé et estimé
                TimeSpan elapsed = DateTime.Now - _startTime - _pausedTime;
                TimeSpan? estimatedTotal = null;
                if (percent > 0 && percent < 100)
                {
                    double estimatedSeconds = elapsed.TotalSeconds * 100 / percent;
                    estimatedTotal = TimeSpan.FromSeconds(estimatedSeconds);
                }
                
                // Formatage du temps
                string elapsedStr = FormatTimeSpan(elapsed);
                string estimatedStr = estimatedTotal.HasValue 
                    ? FormatTimeSpan(estimatedTotal.Value)
                    : "00:00";
                TxtProgressTimeElapsed.Text = elapsedStr;
                TxtProgressTimeEstimated.Text = estimatedStr;
                
                // Si percent < 0, garder le pourcentage actuel (utilisé pour pause)
                if (percent >= 0)
                {
                    TxtProgressPercent.Text = percent > 0 ? $"{percent:F0}%" : "";
                    
                    // Mettre a jour la largeur de la barre de progression (style Creer Module)
                    if (ProgressBarFill.Parent is Grid parentGrid)
                    {
                        double maxWidth = parentGrid.ActualWidth > 0 ? parentGrid.ActualWidth : 400;
                        double fillWidth = (percent / 100.0) * maxWidth;
                        ProgressBarFill.Width = fillWidth;
                    }
                }
            });
        }
        
        private string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalHours >= 1)
                return $"{(int)ts.TotalHours}h {ts.Minutes:D2}m {ts.Seconds:D2}s";
            else if (ts.TotalMinutes >= 1)
                return $"{ts.Minutes}m {ts.Seconds:D2}s";
            else
                return $"{ts.Seconds}s";
        }

        // ====================================================================
        // Selection et filtres DataGrid
        // ====================================================================
        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var file in AllFiles)
                file.IsSelected = true;
            UpdateStatistics();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var file in AllFiles)
                file.IsSelected = false;
            UpdateStatistics();
        }

        private void DataGrid_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                if (sender is DataGrid dataGrid)
                {
                    // Verifier si on a clique directement sur la checkbox
                    var originalSource = e.OriginalSource as FrameworkElement;
                    if (originalSource != null)
                    {
                        var parent = originalSource;
                        while (parent != null)
                        {
                            if (parent is CheckBox)
                                return;
                            parent = VisualTreeHelper.GetParent(parent) as FrameworkElement;
                        }
                    }

                    // Toggle la selection de l'item clique
                    if (dataGrid.SelectedItem is TemplateFileItem fileItem)
                    {
                        fileItem.IsSelected = !fileItem.IsSelected;
                        UpdateStatistics();
                    }
                }
            }
            catch { }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void CmbExtension_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void CmbState_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        /// <summary>
        /// Peuple le filtre d'extension dynamiquement avec les extensions presentes dans les fichiers charges
        /// </summary>
        private void PopulateExtensionFilter()
        {
            if (_allFilesList == null || _allFilesList.Count == 0 || CmbExtension == null)
                return;

            // Extraire les extensions uniques et les trier
            var uniqueExtensions = _allFilesList
                .Select(f => f.FileExtension?.ToLowerInvariant() ?? "")
                .Where(ext => !string.IsNullOrWhiteSpace(ext))
                .Distinct()
                .OrderBy(ext => ext)
                .ToList();

            // Garder "Tous" et ajouter les extensions dynamiques
            CmbExtension.Items.Clear();
            CmbExtension.Items.Add(new ComboBoxItem { Content = "Tous", IsSelected = true });

            foreach (var ext in uniqueExtensions)
            {
                CmbExtension.Items.Add(new ComboBoxItem { Content = ext });
            }

            CmbExtension.SelectedIndex = 0;
            Log($"[i] Filtre extension: {uniqueExtensions.Count} extensions detectees", LogLevel.INFO);
        }

        private void ApplyFilters()
        {
            if (_allFilesList == null || _allFilesList.Count == 0)
                return;

            string searchText = TxtSearch?.Text?.ToLower() ?? "";
            string extensionFilter = (CmbExtension?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Tous";
            string stateFilter = (CmbState?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Tous";

            AllFiles.Clear();

            foreach (var file in _allFilesList)
            {
                bool matchSearch = string.IsNullOrEmpty(searchText) || 
                                   file.FileName.ToLower().Contains(searchText) ||
                                   file.RelativePath.ToLower().Contains(searchText);

                bool matchExtension = extensionFilter == "Tous" || 
                                      file.FileExtension.Equals(extensionFilter, StringComparison.OrdinalIgnoreCase);

                // Filtre etat (comme VaultUploadModuleWindow)
                bool matchState = true;
                switch (stateFilter)
                {
                    case "En attente":
                        matchState = file.Status == "En attente";
                        break;
                    case "Uploade":
                        matchState = file.Status?.Contains("Uploade") == true || file.Status?.Contains("[+]") == true || file.Status?.Contains("✅") == true;
                        break;
                    case "Ignore":
                        matchState = file.Status?.Contains("Ignore") == true || file.Status?.Contains("Existe") == true;
                        break;
                    case "Erreur":
                        matchState = file.Status?.Contains("Erreur") == true || file.Status?.Contains("[-]") == true || file.Status?.Contains("❌") == true;
                        break;
                }

                if (matchSearch && matchExtension && matchState)
                {
                    AllFiles.Add(file);
                }
            }
        }

        private enum LogLevel { INFO, SUCCESS, WARNING, ERROR }
    }

    // ====================================================================
    // Classe pour les stats d'extension
    // ====================================================================
    public class ExtensionInfo
    {
        public string Extension { get; set; }
        public int Count { get; set; }
        public double Percentage { get; set; }
    }
}
