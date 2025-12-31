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
using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;

namespace XnrgyEngineeringAutomationTools.Views
{
    /// <summary>
    /// Fichier a uploader - avec selection et statut
    /// </summary>
    public class TemplateFileItem : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        private string _status = "En attente";

        public string FileName { get; set; } = string.Empty;
        public string FileExtension { get; set; } = string.Empty;
        public string FileSizeFormatted { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string RelativePath { get; set; } = string.Empty;

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
        }

        // ====================================================================
        // Chargement fenetre
        // ====================================================================
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Note: Logo remplace par emoji dans le header moderne

            // Afficher le statut de connexion (format identique VaultUploadModule)
            if (_isConnected)
            {
                StatusText.Text = $"Vault: {_vaultService?.VaultName}";
                TxtVaultUser.Text = $"Utilisateur: {_vaultService?.UserName}";
                StatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#107C10"));
                Log($"[+] Connexion Vault heritee de l'application principale", LogLevel.SUCCESS);
                Log($"[i] Vault: {_vaultService?.VaultName} | User: {_vaultService?.UserName}", LogLevel.INFO);
                BtnStartUpload.IsEnabled = false; // Active apres scan
            }
            else
            {
                StatusText.Text = "Vault: Non connecte";
                TxtVaultUser.Text = "Veuillez vous connecter";
                StatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E81123"));
                Log("[-] Aucune connexion Vault. Veuillez vous connecter via l'application principale.", LogLevel.ERROR);
                BtnStartUpload.IsEnabled = false;
            }
            
            Log("[i] Fenetre Upload Template initialisee", LogLevel.INFO);
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

            try
            {
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

                        fileItems.Add(new TemplateFileItem
                        {
                            FileName = fileInfo.Name,
                            FileExtension = fileInfo.Extension.ToLower(),
                            FileSizeFormatted = FormatFileSize(fileInfo.Length),
                            FullPath = filePath,
                            RelativePath = relativeDir,
                            IsSelected = true,
                            Status = "En attente"
                        });
                    }
                });

                // Ajouter au DataGrid sur le thread UI
                foreach (var item in fileItems)
                {
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
            _startTime = DateTime.Now;
            _createFoldersAutomatically = ChkCreateFolders.IsChecked == true;

            // UI - Activer boutons de controle
            BtnStartUpload.IsEnabled = false;
            BtnPause.IsEnabled = true;
            BtnStop.IsEnabled = true;
            BtnCancel.IsEnabled = false;
            BtnScan.IsEnabled = false;
            BtnBrowse.IsEnabled = false;
            BtnPause.Content = "⏸️ PAUSE";

            Log("[>] Debut de l'upload...", LogLevel.INFO);

            try
            {
                string localRoot = TxtLocalPath.Text.Trim();
                var token = _cancellationTokenSource.Token;
                int counter = 0;
                int totalSelected = selectedFiles.Count;

                foreach (var fileItem in selectedFiles)
                {
                    if (token.IsCancellationRequested)
                    {
                        Log("[!] Upload annule par l'utilisateur", LogLevel.WARNING);
                        break;
                    }

                    // Attendre si en pause
                    _pauseEvent.Wait(token);

                    counter++;
                    string filePath = fileItem.FullPath;
                    string fileName = fileItem.FileName;
                    string vaultFolder = GetVaultFolderPath(filePath, localRoot);

                    // Mise a jour statut dans DataGrid
                    Dispatcher.Invoke(() => { fileItem.Status = "Upload en cours..."; });

                    // Mise a jour progression
                    double progress = (counter * 100.0) / totalSelected;
                    var elapsed = DateTime.Now - _startTime;
                    double rate = counter / Math.Max(elapsed.TotalMinutes, 0.01);
                    double remaining = (totalSelected - counter) / Math.Max(rate, 1);

                    UpdateProgress(
                        $"[{counter}/{totalSelected}] {fileName}",
                        progress);

                    // Upload
                    bool success = await Task.Run(() => UploadFile(filePath, vaultFolder), token);

                    // Mise a jour statut
                    Dispatcher.Invoke(() =>
                    {
                        if (success)
                            fileItem.Status = "Uploade";
                        else if (_skippedFiles > 0 && fileItem.Status == "Upload en cours...")
                            fileItem.Status = "Ignore (existe)";
                        else
                            fileItem.Status = "Echec";
                            
                        TxtUploaded.Text = _uploadedFiles.ToString("N0");
                        TxtSkipped.Text = _skippedFiles.ToString("N0");
                        TxtFailed.Text = _failedFiles.ToString("N0");
                        UpdateStatistics();
                    });
                }

                // Resume
                var totalTime = DateTime.Now - _startTime;
                Log($"[+] Upload termine en {totalTime.TotalMinutes:F1} minutes", LogLevel.SUCCESS);
                Log($"    Uploades: {_uploadedFiles} | Ignores: {_skippedFiles} | Echecs: {_failedFiles}", LogLevel.INFO);

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
                    _cancellationTokenSource?.Cancel();
                    _pauseEvent.Set(); // Debloquer si en pause
                    Log("[!] Upload arrete par l'utilisateur", LogLevel.WARNING);
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
                    Dispatcher.Invoke(() => Log($"[-] Dossier introuvable: {vaultFolderPath}", LogLevel.ERROR));
                    return false;
                }

                // Verifier si le fichier existe deja
                if (ChkSkipExisting.IsChecked == true)
                {
                    try
                    {
                        var existingFiles = _connection.WebServiceManager.DocumentService
                            .GetLatestFilesByFolderId(folder.Id, false);

                        if (existingFiles != null && existingFiles.Any(f => f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                        {
                            Interlocked.Increment(ref _skippedFiles);
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
                        "Upload Template - XNRGY",
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
                        return true;
                    }
                }

                Interlocked.Increment(ref _failedFiles);
                return false;
            }
            catch (Exception ex)
            {
                // Fichier existant?
                if (ex.Message.Contains("already exists") || ex.Message.Contains("1008"))
                {
                    Interlocked.Increment(ref _skippedFiles);
                    return true;
                }

                Interlocked.Increment(ref _failedFiles);
                Dispatcher.Invoke(() => Log($"[-] {fileName}: {ex.Message}", LogLevel.ERROR));
                return false;
            }
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
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                
                // Determiner la couleur selon le niveau
                Color color = level switch
                {
                    LogLevel.SUCCESS => (Color)ColorConverter.ConvertFromString("#2ECC71"), // Vert
                    LogLevel.ERROR => (Color)ColorConverter.ConvertFromString("#E74C3C"),   // Rouge
                    LogLevel.WARNING => (Color)ColorConverter.ConvertFromString("#F39C12"), // Orange
                    _ => (Color)ColorConverter.ConvertFromString("#B8D4E8")                 // Bleu clair (Info)
                };
                
                // Determiner le prefix selon le niveau (sans emoji)
                string prefix = level switch
                {
                    LogLevel.SUCCESS => "[+]",
                    LogLevel.ERROR => "[-]",
                    LogLevel.WARNING => "[!]",
                    _ => "[>]"
                };
                
                // Creer le paragraph avec couleur
                var paragraph = new System.Windows.Documents.Paragraph();
                paragraph.Margin = new Thickness(0, 2, 0, 2);
                
                // Timestamp en gris
                var timestampRun = new System.Windows.Documents.Run($"[{timestamp}] ")
                {
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#888888"))
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

        private void UpdateProgress(string status, double percent, string speed = "", string remaining = "")
        {
            Dispatcher.Invoke(() =>
            {
                TxtProgressStatus.Text = status;
                
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
                        ProgressBarShine.Width = fillWidth;
                    }
                }
            });
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
                        matchState = file.Status?.Contains("Uploade") == true || file.Status?.Contains("[+]") == true;
                        break;
                    case "Ignore":
                        matchState = file.Status?.Contains("Ignore") == true || file.Status?.Contains("Existe") == true;
                        break;
                    case "Erreur":
                        matchState = file.Status?.Contains("Erreur") == true || file.Status?.Contains("[-]") == true;
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
