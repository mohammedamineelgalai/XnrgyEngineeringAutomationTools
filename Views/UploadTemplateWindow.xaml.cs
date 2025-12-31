using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
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
        private List<string> _filesToUpload = new List<string>();
        private Dictionary<string, ACW.Folder> _folderCache = new Dictionary<string, ACW.Folder>();

        // ====================================================================
        // Statistiques
        // ====================================================================
        private int _totalFiles = 0;
        private int _totalFolders = 0;
        private int _uploadedFiles = 0;
        private int _skippedFiles = 0;
        private int _failedFiles = 0;
        private int _foldersCreated = 0;
        private DateTime _startTime;

        // ====================================================================
        // Exclusions
        // ====================================================================
        private static readonly string[] ExcludedExtensions = { ".bak", ".old", ".tmp", ".log", ".lck", ".dwl", ".dwl2", ".v" };
        private static readonly string[] ExcludedPrefixes = { "~$", "._", "Backup_", ".~" };
        private static readonly string[] ExcludedFolders = { "OldVersions", "Backup", ".vault", ".git", ".vs", "obj", "bin", "Workspace", "vltcache" };
        private static readonly string[] ExcludedFileNames = { "desktop.ini", "Thumbs.db", ".DS_Store" };

        // ====================================================================
        // Data Binding pour extensions
        // ====================================================================
        public ObservableCollection<ExtensionInfo> ExtensionStats { get; set; } = new ObservableCollection<ExtensionInfo>();

        /// <summary>
        /// Constructeur avec service Vault partage
        /// </summary>
        public UploadTemplateWindow(VaultSdkService? vaultService)
        {
            InitializeComponent();
            LstExtensions.ItemsSource = ExtensionStats;
            
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
            // Charger le logo
            try
            {
                string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "xnrgy_logo.png");
                if (File.Exists(logoPath))
                {
                    LogoImage.Source = new BitmapImage(new Uri(logoPath));
                }
            }
            catch { }

            // Afficher le statut de connexion
            if (_isConnected)
            {
                UpdateStatus($"Connecte - {_vaultService?.VaultName}", "#107C10");
                Log($"[+] Connexion Vault heritee de l'application principale", LogLevel.SUCCESS);
                Log($"[i] Vault: {_vaultService?.VaultName} | User: {_vaultService?.UserName}", LogLevel.INFO);
                BtnStartUpload.IsEnabled = true;
            }
            else
            {
                UpdateStatus("Non connecte - Vault requis", "#E81123");
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
                TxtVaultPath.Text = string.IsNullOrEmpty(relative) ? "Destination Vault: $/" : $"Destination Vault: $/{relative.Replace("\\", "/")}";
            }
            else
            {
                TxtVaultPath.Text = "Destination Vault: $/";
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
            Log($"[?] Scan des fichiers dans {localPath}...", LogLevel.INFO);
            UpdateProgress("Scan en cours...", 0);

            try
            {
                _filesToUpload.Clear();
                ExtensionStats.Clear();

                await Task.Run(() =>
                {
                    var allFiles = Directory.EnumerateFiles(localPath, "*.*", SearchOption.AllDirectories)
                        .Where(f => !ShouldExclude(f))
                        .ToList();

                    _filesToUpload = allFiles;
                });

                _totalFiles = _filesToUpload.Count;
                _totalFolders = _filesToUpload.Select(f => Path.GetDirectoryName(f)).Distinct().Count();

                // Statistiques par extension
                var byExtension = _filesToUpload
                    .GroupBy(f => Path.GetExtension(f).ToLower())
                    .OrderByDescending(g => g.Count())
                    .Take(10);

                double maxCount = byExtension.Any() ? byExtension.Max(g => g.Count()) : 1;

                foreach (var grp in byExtension)
                {
                    ExtensionStats.Add(new ExtensionInfo
                    {
                        Extension = grp.Key,
                        Count = grp.Count(),
                        Percentage = (grp.Count() / maxCount) * 100
                    });
                }

                // Mise a jour UI
                TxtTotalFiles.Text = _totalFiles.ToString("N0");
                TxtTotalFolders.Text = _totalFolders.ToString("N0");
                TxtUploaded.Text = "0";
                TxtSkipped.Text = "0";
                TxtFailed.Text = "0";
                TxtFoldersCreated.Text = "0";

                Log($"[+] Scan termine: {_totalFiles} fichiers dans {_totalFolders} dossiers", LogLevel.SUCCESS);
                UpdateProgress($"{_totalFiles} fichiers prets a uploader", 0);

                BtnStartUpload.IsEnabled = _isConnected && _totalFiles > 0;
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

        private bool ShouldExclude(string filePath)
        {
            string fileName = Path.GetFileName(filePath);

            // Noms exclus
            if (ExcludedFileNames.Any(n => fileName.Equals(n, StringComparison.OrdinalIgnoreCase)))
                return true;

            // Extensions exclues
            if (ExcludedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                return true;

            // Prefixes exclus
            if (ExcludedPrefixes.Any(p => fileName.StartsWith(p)))
                return true;

            // Dossiers exclus (si option cochee)
            if (ChkExcludeOldVersions.IsChecked == true)
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
            if (!_isConnected || _filesToUpload.Count == 0)
            {
                XnrgyMessageBox.ShowError("Veuillez vous connecter et scanner les fichiers d'abord.", "Erreur", this);
                return;
            }

            // Confirmation
            var confirm = XnrgyMessageBox.Show(
                $"Vous allez uploader {_totalFiles} fichiers vers le Vault PRODUCTION.\n\n" +
                $"Source: {TxtLocalPath.Text}\n" +
                $"Destination: {TxtVaultPath.Text.Replace("Destination Vault: ", "")}\n\n" +
                "Continuer?",
                "Confirmer Upload",
                XnrgyMessageBoxType.Warning,
                XnrgyMessageBoxButtons.YesNo,
                this);

            if (confirm != XnrgyMessageBoxResult.Yes)
                return;

            // Initialiser
            _isUploading = true;
            _cancellationTokenSource = new CancellationTokenSource();
            _uploadedFiles = 0;
            _skippedFiles = 0;
            _failedFiles = 0;
            _foldersCreated = 0;
            _folderCache.Clear();
            _startTime = DateTime.Now;

            // UI
            BtnStartUpload.IsEnabled = false;
            BtnCancel.IsEnabled = true;
            BtnScan.IsEnabled = false;
            BtnBrowse.IsEnabled = false;

            Log("[>] Debut de l'upload...", LogLevel.INFO);

            try
            {
                string localRoot = TxtLocalPath.Text.Trim();
                var token = _cancellationTokenSource.Token;
                int counter = 0;

                foreach (var filePath in _filesToUpload)
                {
                    if (token.IsCancellationRequested)
                    {
                        Log("[!] Upload annule par l'utilisateur", LogLevel.WARNING);
                        break;
                    }

                    counter++;
                    string fileName = Path.GetFileName(filePath);
                    string vaultFolder = GetVaultFolderPath(filePath, localRoot);

                    // Mise a jour progression
                    double progress = (counter * 100.0) / _totalFiles;
                    var elapsed = DateTime.Now - _startTime;
                    double rate = counter / Math.Max(elapsed.TotalMinutes, 0.01);
                    double remaining = (_totalFiles - counter) / Math.Max(rate, 1);

                    UpdateProgress(
                        $"[{counter}/{_totalFiles}] {fileName}",
                        progress,
                        $"{rate:F1} fichiers/min",
                        $"Reste: ~{remaining:F1} min");

                    TxtCurrentFile.Text = filePath;

                    // Upload
                    bool success = await Task.Run(() => UploadFile(filePath, vaultFolder), token);

                    // Mise a jour stats
                    Dispatcher.Invoke(() =>
                    {
                        TxtUploaded.Text = _uploadedFiles.ToString("N0");
                        TxtSkipped.Text = _skippedFiles.ToString("N0");
                        TxtFailed.Text = _failedFiles.ToString("N0");
                        TxtFoldersCreated.Text = _foldersCreated.ToString("N0");
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
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;

                // Restaurer UI
                BtnStartUpload.IsEnabled = _isConnected && _totalFiles > 0;
                BtnCancel.IsEnabled = false;
                BtnScan.IsEnabled = true;
                BtnBrowse.IsEnabled = true;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource?.Cancel();
            BtnCancel.IsEnabled = false;
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
                // Creer le dossier
                if (ChkCreateFolders.IsChecked != true)
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
            TxtLog.Clear();
        }

        private void Log(string message, LogLevel level)
        {
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                TxtLog.AppendText($"[{timestamp}] {message}\n");
                LogScrollViewer.ScrollToEnd();
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
                TxtProgressPercent.Text = $"{percent:F0}%";
                ProgressBar.Value = percent;
                TxtSpeed.Text = speed;
                TxtTimeRemaining.Text = remaining;
            });
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
