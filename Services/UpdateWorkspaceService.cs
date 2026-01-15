#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service de mise a jour du workspace local depuis Vault.
    /// Telecharge les dossiers critiques, copie les plugins et execute les installations silencieuses.
    /// </summary>
    public class UpdateWorkspaceService
    {
        #region Configuration Constants

        // Chemins Vault source (utiliser / comme separateur pour Vault)
        private static readonly string[] VaultFolderPaths = new[]
        {
            "$/Engineering/Inventor_Standards/",
            "$/Engineering/Library/Cabinet/",
            "$/Engineering/Library/Xnrgy_M99/",
            "$/Engineering/Library/Xnrgy_Module/"
        };

        // Chemins locaux destination correspondants
        private static readonly string[] LocalFolderPaths = new[]
        {
            @"C:\Vault\Engineering\Inventor_Standards\",
            @"C:\Vault\Engineering\Library\Cabinet\",
            @"C:\Vault\Engineering\Library\Xnrgy_M99\",
            @"C:\Vault\Engineering\Library\Xnrgy_Module\"
        };

        // Plugins a copier (depuis Application_Plugins)
        private static readonly (string SourceSubPath, string DestFolder)[] PluginCopyPaths = new[]
        {
            (@"Automation_Standard\Application_Plugins\SIBL_XNRGY_ADDINS_2026", "SIBL_XNRGY_ADDINS_2026"),
            (@"Automation_Standard\Application_Plugins\XNRGY_ADDINS_2026", "XNRGY_ADDINS_2026")
        };

        // Dossier ApplicationPlugins d'Autodesk
        private const string APPLICATION_PLUGINS_PATH = @"C:\ProgramData\Autodesk\ApplicationPlugins";

        // Dossiers a exclure lors de la copie des plugins
        private static readonly string[] ExcludedFolders = new[]
        {
            "Xnrgy_Software",
            "Automation_Data"
        };

        // Dossier contenant les installateurs (scan automatique)
        private const string INSTALLERS_FOLDER = @"C:\Vault\Engineering\Inventor_Standards\Automation_Standard\Application_Plugins\XNRGY_ADDINS_2026\Xnrgy_Software";
        
        // Arguments par defaut pour les installateurs Inno Setup
        private const string SILENT_INSTALL_ARGS = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-";

        #endregion

        #region Progress Reporting

        public event EventHandler<UpdateProgressEventArgs>? ProgressChanged;
        public event EventHandler<UpdateLogEventArgs>? LogMessage;
        public event EventHandler<UpdateStepEventArgs>? StepChanged;

        private void ReportProgress(int percent, string status, string? currentFile = null)
        {
            ProgressChanged?.Invoke(this, new UpdateProgressEventArgs(percent, status, currentFile));
        }

        private void Log(string message, LogLevel level = LogLevel.INFO)
        {
            LogMessage?.Invoke(this, new UpdateLogEventArgs(message, level));
            Logger.Log(message, level == LogLevel.ERROR ? Logger.LogLevel.ERROR 
                : level == LogLevel.WARNING ? Logger.LogLevel.WARNING 
                : Logger.LogLevel.INFO);
        }

        private void UpdateStep(int stepNumber, StepStatus status, string? message = null)
        {
            StepChanged?.Invoke(this, new UpdateStepEventArgs(stepNumber, status, message));
        }

        #endregion

        #region Main Execution

        /// <summary>
        /// Execute la mise a jour complete du workspace
        /// </summary>
        /// <param name="connection">Connexion Vault active</param>
        /// <param name="cancellationToken">Token d'annulation</param>
        /// <returns>Resultat de la mise a jour</returns>
        public async Task<UpdateWorkspaceResult> ExecuteFullUpdateAsync(
            VDF.Vault.Currency.Connections.Connection connection,
            CancellationToken cancellationToken = default)
        {
            var result = new UpdateWorkspaceResult();
            var stopwatch = Stopwatch.StartNew();

            try
            {
                Log("[>] Demarrage de la mise a jour du workspace...");

                // Verifier si Inventor est en cours (bloquerait les fichiers)
                var inventorProcess = Process.GetProcessesByName("Inventor").FirstOrDefault();
                if (inventorProcess != null)
                {
                    Log("[!] Inventor detecte en cours d'execution", LogLevel.WARNING);
                    Log("[>] Fermeture d'Inventor pour permettre la mise a jour...");
                    
                    try
                    {
                        // Demander fermeture propre via COM si possible
                        try
                        {
                            var inventorApp = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                            inventorApp.Quit();
                            Log("[+] Commande de fermeture envoyee a Inventor");
                        }
                        catch
                        {
                            // Si COM echoue, fermer le processus
                            inventorProcess.CloseMainWindow();
                        }
                        
                        // Attendre la fermeture (max 30 secondes)
                        bool closed = await Task.Run(() => inventorProcess.WaitForExit(30000), cancellationToken);
                        if (!closed)
                        {
                            inventorProcess.Kill();
                            await Task.Delay(2000, cancellationToken);
                        }
                        Log("[+] Inventor ferme avec succes");
                    }
                    catch (Exception ex)
                    {
                        Log($"[!] Impossible de fermer Inventor: {ex.Message}", LogLevel.WARNING);
                        Log("[!] Certains fichiers pourraient etre verouilles", LogLevel.WARNING);
                    }
                }

                // Etape 1: Verifier la connexion Vault
                UpdateStep(1, StepStatus.InProgress, "Verification de la connexion...");
                if (connection == null)
                {
                    UpdateStep(1, StepStatus.Failed, "Connexion Vault requise");
                    result.Success = false;
                    result.ErrorMessage = "Aucune connexion Vault active";
                    return result;
                }
                UpdateStep(1, StepStatus.Completed, "Connexion verifiee");
                Log("[+] Connexion Vault verifiee");

                cancellationToken.ThrowIfCancellationRequested();

                // Etapes 2-5: Telecharger les 4 dossiers Vault
                for (int i = 0; i < VaultFolderPaths.Length; i++)
                {
                    int stepNumber = i + 2; // Etapes 2, 3, 4, 5
                    var vaultPath = VaultFolderPaths[i];
                    var localPath = LocalFolderPaths[i];
                    var folderName = Path.GetFileName(vaultPath.TrimEnd('/'));

                    UpdateStep(stepNumber, StepStatus.InProgress, $"Telechargement {folderName}...");
                    
                    try
                    {
                        var downloadResult = await DownloadVaultFolderAsync(
                            connection, vaultPath, localPath, cancellationToken);
                        
                        if (downloadResult.Success)
                        {
                            UpdateStep(stepNumber, StepStatus.Completed, $"{downloadResult.FileCount} fichiers");
                            result.DownloadedFiles += downloadResult.FileCount;
                            Log($"[+] {folderName}: {downloadResult.FileCount} fichiers telecharges");
                        }
                        else
                        {
                            UpdateStep(stepNumber, StepStatus.Warning, downloadResult.ErrorMessage ?? "Erreur partielle");
                            Log($"[!] {folderName}: {downloadResult.ErrorMessage}", LogLevel.WARNING);
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStep(stepNumber, StepStatus.Failed, ex.Message);
                        Log($"[-] Erreur telechargement {folderName}: {ex.Message}", LogLevel.ERROR);
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                }

                // Etapes 6-7: Copier les plugins (necessite droits admin)
                bool isAdmin = new System.Security.Principal.WindowsPrincipal(
                    System.Security.Principal.WindowsIdentity.GetCurrent())
                    .IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);

                if (!isAdmin)
                {
                    Log("[!] L'application n'a pas les droits administrateur", LogLevel.WARNING);
                    Log("[!] La copie des plugins et l'installation peuvent echouer", LogLevel.WARNING);
                    Log("[i] Relancez l'application en tant qu'administrateur pour cette fonctionnalite");
                }

                for (int i = 0; i < PluginCopyPaths.Length; i++)
                {
                    int stepNumber = i + 6; // Etapes 6, 7
                    var (sourceSubPath, destFolder) = PluginCopyPaths[i];

                    UpdateStep(stepNumber, StepStatus.InProgress, $"Copie {destFolder}...");
                    
                    try
                    {
                        var copyResult = await CopyPluginFolderAsync(
                            sourceSubPath, destFolder, cancellationToken);
                        
                        if (copyResult.Success)
                        {
                            UpdateStep(stepNumber, StepStatus.Completed, $"{copyResult.FileCount} fichiers");
                            result.CopiedPluginFiles += copyResult.FileCount;
                            Log($"[+] {destFolder}: {copyResult.FileCount} fichiers copies");
                        }
                        else
                        {
                            UpdateStep(stepNumber, StepStatus.Warning, copyResult.ErrorMessage ?? "Erreur");
                            Log($"[!] {destFolder}: {copyResult.ErrorMessage}", LogLevel.WARNING);
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStep(stepNumber, StepStatus.Failed, ex.Message);
                        Log($"[-] Erreur copie {destFolder}: {ex.Message}", LogLevel.ERROR);
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                }

                // Etape 8: Installations silencieuses
                UpdateStep(8, StepStatus.InProgress, "Installation des outils...");
                try
                {
                    var installResult = await RunSilentInstallersAsync(cancellationToken);
                    
                    if (installResult.Success)
                    {
                        UpdateStep(8, StepStatus.Completed, $"{installResult.SuccessCount}/{installResult.TotalCount} installes");
                        result.InstalledTools = installResult.SuccessCount;
                        Log($"[+] {installResult.SuccessCount}/{installResult.TotalCount} outils installes avec succes");
                    }
                    else
                    {
                        UpdateStep(8, StepStatus.Warning, installResult.ErrorMessage ?? "Erreurs d'installation");
                        Log($"[!] Installation: {installResult.ErrorMessage}", LogLevel.WARNING);
                    }
                }
                catch (Exception ex)
                {
                    UpdateStep(8, StepStatus.Failed, ex.Message);
                    Log($"[-] Erreur installations: {ex.Message}", LogLevel.ERROR);
                }

                // Etape 9: Finalisation
                UpdateStep(9, StepStatus.InProgress, "Finalisation...");
                await Task.Delay(500, cancellationToken); // Court delai pour s'assurer que tout est termine
                UpdateStep(9, StepStatus.Completed, "Mise a jour terminee");

                stopwatch.Stop();
                result.Success = true;
                result.Duration = stopwatch.Elapsed;
                Log($"[+] Mise a jour terminee en {stopwatch.Elapsed.TotalSeconds:F1}s");
                
                ReportProgress(100, "Mise a jour terminee!");
            }
            catch (OperationCanceledException)
            {
                Log("[!] Mise a jour annulee par l'utilisateur", LogLevel.WARNING);
                result.Success = false;
                result.ErrorMessage = "Annule par l'utilisateur";
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur critique: {ex.Message}", LogLevel.ERROR);
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        #endregion

        #region Vault Download

        /// <summary>
        /// Telecharge recursivement un dossier Vault vers un chemin local
        /// </summary>
        private async Task<DownloadResult> DownloadVaultFolderAsync(
            VDF.Vault.Currency.Connections.Connection connection,
            string vaultFolderPath,
            string localFolderPath,
            CancellationToken cancellationToken)
        {
            var result = new DownloadResult();

            return await Task.Run(() =>
            {
                try
                {
                    Log($"   [>] Telechargement: {vaultFolderPath}");

                    // Obtenir le dossier Vault
                    var folder = connection.WebServiceManager.DocumentService.GetFolderByPath(vaultFolderPath);
                    if (folder == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Dossier non trouve: {vaultFolderPath}";
                        return result;
                    }

                    // Collecter tous les fichiers recursivement
                    var allFiles = new List<ACW.File>();
                    var allFolders = new List<ACW.Folder>();
                    GetAllFilesRecursive(connection, folder, allFiles, allFolders);

                    Log($"   [i] {allFiles.Count} fichiers dans {allFolders.Count} dossiers");

                    if (allFiles.Count == 0)
                    {
                        result.Success = true;
                        result.FileCount = 0;
                        return result;
                    }

                    cancellationToken.ThrowIfCancellationRequested();

                    // Creer les dossiers locaux
                    EnsureLocalFolderStructure(connection, vaultFolderPath, localFolderPath, allFolders);

                    // Telecharger les fichiers par batch
                    int batchSize = 50;
                    int downloadedCount = 0;
                    int totalFiles = allFiles.Count;

                    for (int i = 0; i < totalFiles; i += batchSize)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var batch = allFiles.Skip(i).Take(batchSize).ToList();
                        
                        try
                        {
                            var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection, false);
                            
                            // Options: Download = telecharger sans checkout
                            foreach (var file in batch)
                            {
                                var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(connection, file);
                                downloadSettings.AddFileToAcquire(fileIteration, 
                                    VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                            }

                            var downloadResult = connection.FileManager.AcquireFiles(downloadSettings);
                            
                            if (downloadResult?.FileResults != null)
                            {
                                foreach (var fileResult in downloadResult.FileResults)
                                {
                                    if (fileResult.Status == VDF.Vault.Results.FileAcquisitionResult.AcquisitionStatus.Success)
                                    {
                                        downloadedCount++;
                                    }
                                }
                            }

                            int progressPercent = (int)((i + batch.Count) * 100.0 / totalFiles);
                            ReportProgress(progressPercent, $"Telechargement {downloadedCount}/{totalFiles}...", 
                                batch.FirstOrDefault()?.Name);
                        }
                        catch (Exception batchEx)
                        {
                            Log($"   [!] Erreur batch: {batchEx.Message}", LogLevel.WARNING);
                        }
                    }

                    result.Success = true;
                    result.FileCount = downloadedCount;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }

                return result;
            }, cancellationToken);
        }

        /// <summary>
        /// Obtient tous les fichiers recursivement depuis un dossier Vault
        /// </summary>
        private void GetAllFilesRecursive(
            VDF.Vault.Currency.Connections.Connection connection, 
            ACW.Folder folder, 
            List<ACW.File> allFiles, 
            List<ACW.Folder> allFolders)
        {
            try
            {
                allFolders.Add(folder);

                // Obtenir les fichiers de ce dossier
                var files = connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, false);
                if (files != null && files.Length > 0)
                {
                    allFiles.AddRange(files);
                }

                // Obtenir les sous-dossiers
                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(folder.Id, false);
                if (subFolders != null && subFolders.Length > 0)
                {
                    foreach (var subFolder in subFolders)
                    {
                        GetAllFilesRecursive(connection, subFolder, allFiles, allFolders);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"   [!] Erreur enumeration {folder.FullName}: {ex.Message}", LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Cree la structure de dossiers locaux correspondant a Vault
        /// </summary>
        private void EnsureLocalFolderStructure(
            VDF.Vault.Currency.Connections.Connection connection,
            string vaultRootPath,
            string localRootPath,
            List<ACW.Folder> folders)
        {
            try
            {
                // Creer le dossier racine s'il n'existe pas
                if (!Directory.Exists(localRootPath))
                {
                    Directory.CreateDirectory(localRootPath);
                }

                // Creer chaque sous-dossier
                foreach (var folder in folders)
                {
                    // Calculer le chemin relatif depuis la racine Vault
                    var vaultFullPath = folder.FullName; // ex: $/Engineering/Library/Cabinet/SubFolder
                    
                    // Normaliser les chemins
                    vaultRootPath = vaultRootPath.TrimEnd('/');
                    var relativePath = vaultFullPath.Substring(vaultRootPath.Length).TrimStart('/');
                    
                    if (!string.IsNullOrEmpty(relativePath))
                    {
                        // Convertir separateurs
                        relativePath = relativePath.Replace("/", "\\");
                        var localPath = Path.Combine(localRootPath, relativePath);
                        
                        if (!Directory.Exists(localPath))
                        {
                            Directory.CreateDirectory(localPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"   [!] Erreur creation dossiers: {ex.Message}", LogLevel.WARNING);
            }
        }

        #endregion

        #region Plugin Copy

        /// <summary>
        /// Copie un dossier plugin vers ApplicationPlugins avec exclusions
        /// </summary>
        private async Task<CopyResult> CopyPluginFolderAsync(
            string sourceSubPath,
            string destFolderName,
            CancellationToken cancellationToken)
        {
            var result = new CopyResult();

            return await Task.Run(() =>
            {
                try
                {
                    // Construire les chemins complets
                    string sourcePath = Path.Combine(
                        @"C:\Vault\Engineering\Inventor_Standards",
                        sourceSubPath);
                    
                    string destPath = Path.Combine(APPLICATION_PLUGINS_PATH, destFolderName);

                    Log($"   [>] Copie: {sourceSubPath} -> {destFolderName}");

                    if (!Directory.Exists(sourcePath))
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Source non trouvee: {sourcePath}";
                        return result;
                    }

                    // Creer le dossier destination
                    if (!Directory.Exists(destPath))
                    {
                        Directory.CreateDirectory(destPath);
                    }

                    // Copier recursivement avec exclusions
                    int filesCopied = CopyDirectoryWithExclusions(sourcePath, destPath, cancellationToken);

                    result.Success = true;
                    result.FileCount = filesCopied;
                }
                catch (UnauthorizedAccessException uaEx)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Acces refuse: {uaEx.Message}";
                    Log($"   [-] Acces refuse lors de la copie. Verifier les permissions admin.", LogLevel.ERROR);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }

                return result;
            }, cancellationToken);
        }

        /// <summary>
        /// Copie recursivement un dossier en excluant certains sous-dossiers
        /// </summary>
        private int CopyDirectoryWithExclusions(string sourceDir, string destDir, CancellationToken cancellationToken)
        {
            int filesCopied = 0;

            // Creer le dossier destination
            Directory.CreateDirectory(destDir);

            // Copier les fichiers
            foreach (var file in Directory.GetFiles(sourceDir))
            {
                cancellationToken.ThrowIfCancellationRequested();

                var fileName = Path.GetFileName(file);
                var destFile = Path.Combine(destDir, fileName);

                try
                {
                    // Supprimer le fichier existant s'il est read-only
                    if (File.Exists(destFile))
                    {
                        var fi = new FileInfo(destFile);
                        if (fi.IsReadOnly)
                        {
                            fi.IsReadOnly = false;
                        }
                    }

                    File.Copy(file, destFile, true);
                    filesCopied++;
                }
                catch (Exception ex)
                {
                    Log($"   [!] Erreur copie {fileName}: {ex.Message}", LogLevel.WARNING);
                }
            }

            // Copier les sous-dossiers (avec exclusions)
            foreach (var subDir in Directory.GetDirectories(sourceDir))
            {
                var dirName = Path.GetFileName(subDir);
                
                // Verifier si ce dossier doit etre exclu
                if (ExcludedFolders.Any(ef => ef.Equals(dirName, StringComparison.OrdinalIgnoreCase)))
                {
                    Log($"   [i] Dossier exclu: {dirName}");
                    continue;
                }

                var destSubDir = Path.Combine(destDir, dirName);
                filesCopied += CopyDirectoryWithExclusions(subDir, destSubDir, cancellationToken);
            }

            return filesCopied;
        }

        #endregion

        #region Silent Installers

        /// <summary>
        /// Execute les installateurs silencieux - SCALABLE
        /// Scanne automatiquement le dossier Xnrgy_Software pour tous les .exe
        /// </summary>
        private async Task<InstallResult> RunSilentInstallersAsync(CancellationToken cancellationToken)
        {
            var result = new InstallResult { Success = true };
            int successCount = 0;
            var errors = new List<string>();

            // Enlever les attributs ReadOnly des fichiers existants dans ApplicationPlugins
            // Car Vault les telecharge souvent en ReadOnly
            try
            {
                var pluginPaths = new[]
                {
                    @"C:\ProgramData\Autodesk\ApplicationPlugins\XNRGY_ADDINS_2026",
                    @"C:\ProgramData\Autodesk\ApplicationPlugins\SIBL_XNRGY_ADDINS_2026"
                };
                
                foreach (var pluginPath in pluginPaths)
                {
                    if (Directory.Exists(pluginPath))
                    {
                        foreach (var file in Directory.GetFiles(pluginPath, "*.*", SearchOption.AllDirectories))
                        {
                            try
                            {
                                var fi = new FileInfo(file);
                                if (fi.IsReadOnly)
                                {
                                    fi.IsReadOnly = false;
                                }
                            }
                            catch { /* Ignorer les erreurs individuelles */ }
                        }
                    }
                }
                Log("   [i] Attributs ReadOnly enleves des plugins existants");
            }
            catch (Exception ex)
            {
                Log($"   [!] Impossible d'enlever ReadOnly: {ex.Message}", LogLevel.WARNING);
            }

            // Scanner dynamiquement le dossier Xnrgy_Software pour tous les .exe
            var installers = new List<string>();
            
            if (Directory.Exists(INSTALLERS_FOLDER))
            {
                // Chercher tous les Setup*.exe dans les sous-dossiers
                foreach (var subDir in Directory.GetDirectories(INSTALLERS_FOLDER))
                {
                    var exeFiles = Directory.GetFiles(subDir, "*.exe", SearchOption.TopDirectoryOnly)
                        .Where(f => !Path.GetFileName(f).StartsWith("unins", StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    
                    installers.AddRange(exeFiles);
                }
                
                Log($"   [i] {installers.Count} installateur(s) detecte(s) dans Xnrgy_Software");
            }
            else
            {
                Log($"   [!] Dossier introuvable: {INSTALLERS_FOLDER}", LogLevel.WARNING);
            }

            // Executer chaque installateur
            foreach (var fullPath in installers)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var installerName = Path.GetFileName(fullPath);
                Log($"   [>] Installation: {installerName}");

                try
                {
                    ReportProgress(0, $"Installation de {installerName}...", installerName);

                    // Verifier si on est deja admin
                    bool isAdmin = new System.Security.Principal.WindowsPrincipal(
                        System.Security.Principal.WindowsIdentity.GetCurrent())
                        .IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);

                    var startInfo = new ProcessStartInfo
                    {
                        FileName = fullPath,
                        Arguments = SILENT_INSTALL_ARGS,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        CreateNoWindow = true
                    };

                    if (isAdmin)
                    {
                        // Deja admin - lancer directement sans shell
                        startInfo.UseShellExecute = false;
                        startInfo.RedirectStandardOutput = true;
                        startInfo.RedirectStandardError = true;
                    }
                    else
                    {
                        // Pas admin - demander elevation via shell
                        startInfo.UseShellExecute = true;
                        startInfo.Verb = "runas";
                    }

                    var process = new Process
                    {
                        StartInfo = startInfo,
                        EnableRaisingEvents = true
                    };

                    try
                    {
                        process.Start();
                    }
                    catch (System.ComponentModel.Win32Exception w32ex) when (w32ex.NativeErrorCode == 1223)
                    {
                        // L'utilisateur a annule la demande UAC
                        Log($"   [!] Installation annulee par l'utilisateur: {installerName}", LogLevel.WARNING);
                        errors.Add($"Annule: {installerName}");
                        continue;
                    }

                    // Attendre avec timeout de 2 minutes
                    bool exited = await Task.Run(() => process.WaitForExit(120000), cancellationToken);

                    if (!exited)
                    {
                        process.Kill();
                        Log($"   [!] Timeout pour {installerName}", LogLevel.WARNING);
                        errors.Add($"Timeout: {installerName}");
                    }
                    else if (process.ExitCode == 0 || process.ExitCode == 3010) // 3010 = reboot required
                    {
                        successCount++;
                        Log($"   [+] {installerName} installe (code: {process.ExitCode})");
                    }
                    else if (process.ExitCode == 5)
                    {
                        // Code 5 = Acces refuse - fichier verrouille ou droits insuffisants
                        Log($"   [!] Acces refuse pour {installerName} - fermer Inventor et reessayer", LogLevel.WARNING);
                        errors.Add($"Acces refuse: {installerName}");
                    }
                    else if (process.ExitCode == 1602 || process.ExitCode == 1603)
                    {
                        // Erreurs MSI courantes
                        Log($"   [!] Installation annulee ou echouee: {installerName}", LogLevel.WARNING);
                        errors.Add($"Echec MSI: {installerName}");
                    }
                    else
                    {
                        Log($"   [!] {installerName} termine avec code: {process.ExitCode}", LogLevel.WARNING);
                        errors.Add($"Code {process.ExitCode}: {installerName}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"   [-] Erreur installation {installerName}: {ex.Message}", LogLevel.ERROR);
                    errors.Add($"Erreur: {installerName}");
                }

                // Court delai entre installations
                await Task.Delay(1000, cancellationToken);
            }

            result.SuccessCount = successCount;
            result.TotalCount = installers.Count;
            result.Success = errors.Count == 0;
            if (errors.Count > 0)
            {
                result.ErrorMessage = string.Join("; ", errors);
            }

            return result;
        }

        #endregion

        #region Result Classes

        public class UpdateWorkspaceResult
        {
            public bool Success { get; set; }
            public string? ErrorMessage { get; set; }
            public int DownloadedFiles { get; set; }
            public int CopiedPluginFiles { get; set; }
            public int InstalledTools { get; set; }
            public TimeSpan Duration { get; set; }
        }

        private class DownloadResult
        {
            public bool Success { get; set; }
            public int FileCount { get; set; }
            public string? ErrorMessage { get; set; }
        }

        private class CopyResult
        {
            public bool Success { get; set; }
            public int FileCount { get; set; }
            public string? ErrorMessage { get; set; }
        }

        private class InstallResult
        {
            public bool Success { get; set; }
            public int SuccessCount { get; set; }
            public int TotalCount { get; set; }
            public string? ErrorMessage { get; set; }
        }

        #endregion

        #region Event Args

        public enum LogLevel { INFO, WARNING, ERROR, SUCCESS }
        public enum StepStatus { Pending, InProgress, Completed, Failed, Warning, Skipped }

        public class UpdateProgressEventArgs : EventArgs
        {
            public int Percent { get; }
            public string Status { get; }
            public string? CurrentFile { get; }

            public UpdateProgressEventArgs(int percent, string status, string? currentFile = null)
            {
                Percent = percent;
                Status = status;
                CurrentFile = currentFile;
            }
        }

        public class UpdateLogEventArgs : EventArgs
        {
            public string Message { get; }
            public LogLevel Level { get; }

            public UpdateLogEventArgs(string message, LogLevel level)
            {
                Message = message;
                Level = level;
            }
        }

        public class UpdateStepEventArgs : EventArgs
        {
            public int StepNumber { get; }
            public StepStatus Status { get; }
            public string? Message { get; }

            public UpdateStepEventArgs(int stepNumber, StepStatus status, string? message = null)
            {
                StepNumber = stepNumber;
                Status = status;
                Message = message;
            }
        }

        #endregion
    }
}
