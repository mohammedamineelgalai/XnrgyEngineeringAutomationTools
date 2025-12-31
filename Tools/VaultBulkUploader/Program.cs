// ============================================================================
// VaultBulkUploader - Upload massif C:\Vault vers $/ (Vault PROD)
// Auteur: Mohammed Amine Elgalai - XNRGY Climate Systems
// Date: 2025-12-30
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;

namespace VaultBulkUploader
{
    class Program
    {
        // ====================================================================
        // Configuration PRODUCTION
        // ====================================================================
        private const string VaultServer = "VAULTPOC";
        private const string VaultName = "PROD_XNGRY";  // Vault de production (note: typo admin XNGRY)
        private const string VaultUser = "mohammedamine.elgalai";
        private const string VaultPassword = "Vtr8aPz2*";
        private const string LocalVaultRoot = @"C:\Vault";

        // Exclusions
        private static readonly string[] ExcludedExtensions = { ".bak", ".old", ".tmp", ".log", ".lck", ".dwl", ".dwl2", ".v" };
        private static readonly string[] ExcludedPrefixes = { "~$", "._", "Backup_", ".~" };
        private static readonly string[] ExcludedFolders = { "OldVersions", "Backup", ".vault", ".git", ".vs", "obj", "bin", "Workspace", "vltcache" };
        private static readonly string[] ExcludedFileNames = { "desktop.ini", "Thumbs.db", ".DS_Store" };

        // Statistiques
        private static int _filesTotal = 0;
        private static int _filesUploaded = 0;
        private static int _filesSkipped = 0;
        private static int _filesFailed = 0;
        private static int _foldersCreated = 0;
        private static readonly List<string> _failedFiles = new List<string>();
        private static readonly Dictionary<string, ACW.Folder> _folderCache = new Dictionary<string, ACW.Folder>();

        // Connexion
        private static VDF.Vault.Currency.Connections.Connection? _connection;

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("================================================================");
            Console.WriteLine("  VAULT BULK UPLOADER - XNRGY");
            Console.WriteLine($"  Vault: {VaultName} sur {VaultServer}");
            Console.WriteLine("================================================================");
            Console.ResetColor();
            Console.WriteLine();

            string localRoot = LocalVaultRoot;

            // Verifier argument ligne de commande
            if (args.Length > 0 && Directory.Exists(args[0]))
            {
                localRoot = args[0];
                Log($"Dossier specifie: {localRoot}", "INFO");
            }

            // Verifier que le dossier existe
            if (!Directory.Exists(localRoot))
            {
                Log($"Dossier introuvable: {localRoot}", "ERROR");
                WaitForExit();
                return;
            }

            // Scanner les fichiers
            Log($"Scan des fichiers dans {localRoot}...", "INFO");
            var files = GetFilesToUpload(localRoot);
            _filesTotal = files.Count;

            Log($"Fichiers a traiter: {_filesTotal}", "INFO");
            Console.WriteLine();

            // Afficher apercu par extension
            var byExtension = files.GroupBy(f => Path.GetExtension(f).ToLower())
                                   .OrderByDescending(g => g.Count())
                                   .Take(10);
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("  Apercu par extension:");
            foreach (var grp in byExtension)
            {
                Console.WriteLine($"    {grp.Key}: {grp.Count()} fichiers");
            }
            Console.ResetColor();
            Console.WriteLine();

            if (_filesTotal == 0)
            {
                Log("Aucun fichier a uploader", "WARNING");
                WaitForExit();
                return;
            }

            // Confirmation PROD
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("================================================================");
            Console.WriteLine($"  [!] ATTENTION: VAULT PRODUCTION ({VaultName})");
            Console.WriteLine("================================================================");
            Console.ResetColor();
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"  Vous allez uploader {_filesTotal} fichiers vers PRODUCTION!");
            Console.WriteLine($"  Structure: {localRoot} -> $/...");
            Console.ResetColor();
            Console.WriteLine();
            Console.Write("Tapez 'PROD' pour confirmer: ");
            var confirm = Console.ReadLine();
            
            if (confirm != "PROD")
            {
                Log("Operation annulee", "WARNING");
                WaitForExit();
                return;
            }

            Console.WriteLine();

            // Connexion au Vault
            if (!Connect())
            {
                WaitForExit();
                return;
            }

            Console.WriteLine();
            Log($"Debut de l'upload vers {VaultName}...", "INFO");
            Console.WriteLine("------------------------------------------------------------");

            var startTime = DateTime.Now;
            int counter = 0;
            DateTime lastProgressTime = startTime;

            foreach (var filePath in files)
            {
                counter++;
                
                // Calculer le chemin Vault
                string vaultFolderPath = GetVaultFolderPath(filePath, localRoot);
                string fileName = Path.GetFileName(filePath);

                // Afficher progression tous les 50 fichiers
                if (counter % 50 == 0 || (DateTime.Now - lastProgressTime).TotalSeconds >= 10)
                {
                    var elapsed = (DateTime.Now - startTime).TotalMinutes;
                    var rate = counter / Math.Max(elapsed, 0.01);
                    var remaining = (_filesTotal - counter) / Math.Max(rate, 1);

                    Console.WriteLine();
                    Log($"Progression: {counter}/{_filesTotal} ({counter * 100 / _filesTotal}%)", "INFO");
                    Log($"  Uploades: {_filesUploaded} | Ignores: {_filesSkipped} | Echecs: {_filesFailed}", "INFO");
                    Log($"  Vitesse: {rate:F1} fichiers/min | Reste: ~{remaining:F1} min", "INFO");
                    Console.WriteLine();
                    
                    lastProgressTime = DateTime.Now;
                }

                // Upload
                bool success = UploadFile(filePath, vaultFolderPath);
                
                if (!success && _filesFailed <= 20)
                {
                    Log($"  [-] {fileName} -> {vaultFolderPath}", "ERROR");
                }
            }

            // Deconnexion
            Disconnect();

            // Resume
            var totalTime = DateTime.Now - startTime;
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("================================================================");
            Console.WriteLine($"  RESUME UPLOAD VERS {VaultName}");
            Console.WriteLine("================================================================");
            Console.ResetColor();
            Console.WriteLine();
            
            Log($"Temps total: {totalTime.TotalMinutes:F1} minutes", "INFO");
            Log($"Total fichiers: {_filesTotal}", "INFO");
            Log($"Uploades avec succes: {_filesUploaded}", "SUCCESS");
            Log($"Ignores (deja existants): {_filesSkipped}", "WARNING");
            Log($"Echecs: {_filesFailed}", "ERROR");
            Log($"Dossiers crees: {_foldersCreated}", "INFO");
            Console.WriteLine();

            // Sauvegarder les echecs
            if (_failedFiles.Count > 0)
            {
                string failLogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, 
                    $"Upload-FailedFiles_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(failLogPath, _failedFiles);
                Log($"Liste des echecs sauvegardee: {failLogPath}", "INFO");
            }

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("================================================================");
            if (_filesFailed == 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("  [+] UPLOAD TERMINE AVEC SUCCES!");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"  [!] UPLOAD TERMINE AVEC {_filesFailed} ERREURS");
            }
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("================================================================");
            Console.ResetColor();
            Console.WriteLine();

            WaitForExit();
        }

        // ====================================================================
        // Connexion Vault
        // ====================================================================
        private static bool Connect()
        {
            try
            {
                Log($"Connexion a {VaultName} sur {VaultServer}...", "INFO");

                var result = VDF.Vault.Library.ConnectionManager.LogIn(
                    VaultServer,
                    VaultName,
                    VaultUser,
                    VaultPassword,
                    VDF.Vault.Currency.Connections.AuthenticationFlags.Standard,
                    null
                );

                if (result.Success && result.Connection != null)
                {
                    _connection = result.Connection;
                    Log($"Connecte au Vault '{VaultName}' (User: {VaultUser})", "SUCCESS");
                    return true;
                }
                else
                {
                    Log($"Echec connexion: {result.ErrorMessages?.FirstOrDefault()}", "ERROR");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur connexion: {ex.Message}", "ERROR");
                return false;
            }
        }

        private static void Disconnect()
        {
            try
            {
                if (_connection != null)
                {
                    VDF.Vault.Library.ConnectionManager.LogOut(_connection);
                    Log("Deconnecte du Vault", "INFO");
                }
            }
            catch { }
        }

        // ====================================================================
        // Upload de fichier
        // ====================================================================
        private static bool UploadFile(string localFilePath, string vaultFolderPath)
        {
            if (_connection == null) return false;

            string fileName = Path.GetFileName(localFilePath);

            try
            {
                // Obtenir ou creer le dossier Vault
                var folder = EnsureVaultFolder(vaultFolderPath);
                if (folder == null)
                {
                    _filesFailed++;
                    _failedFiles.Add($"[-] {localFilePath} - Dossier Vault introuvable: {vaultFolderPath}");
                    return false;
                }

                // Verifier si le fichier existe deja
                try
                {
                    var existingFiles = _connection.WebServiceManager.DocumentService
                        .GetLatestFilesByFolderId(folder.Id, false);
                    
                    if (existingFiles != null && existingFiles.Any(f => f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                    {
                        _filesSkipped++;
                        return true; // Skip silencieusement
                    }
                }
                catch { }

                // Upload
                var fileInfo = new FileInfo(localFilePath);
                
                using (var stream = File.OpenRead(localFilePath))
                {
                    var vdfFolder = new VDF.Vault.Currency.Entities.Folder(_connection, folder);

                    var result = _connection.FileManager.AddFile(
                        vdfFolder,
                        fileName,
                        "Upload initial - XNRGY Template",
                        fileInfo.LastWriteTimeUtc,
                        null, // associations
                        null, // bom
                        ACW.FileClassification.None,
                        false, // hidden
                        stream
                    );

                    if (result != null)
                    {
                        _filesUploaded++;
                        return true;
                    }
                }

                _filesFailed++;
                _failedFiles.Add($"[-] {localFilePath} - Upload retourne null");
                return false;
            }
            catch (Exception ex)
            {
                // Verifier si c'est juste un fichier existant
                if (ex.Message.Contains("already exists") || ex.Message.Contains("1008"))
                {
                    _filesSkipped++;
                    return true;
                }

                _filesFailed++;
                _failedFiles.Add($"[-] {localFilePath} - {ex.Message}");
                return false;
            }
        }

        // ====================================================================
        // Gestion des dossiers Vault
        // ====================================================================
        private static ACW.Folder? EnsureVaultFolder(string vaultPath)
        {
            if (_connection == null) return null;

            // Verifier le cache
            if (_folderCache.TryGetValue(vaultPath, out var cachedFolder))
            {
                return cachedFolder;
            }

            try
            {
                // Verifier si le dossier existe
                var folder = _connection.WebServiceManager.DocumentService.GetFolderByPath(vaultPath);
                _folderCache[vaultPath] = folder;
                return folder;
            }
            catch
            {
                // Dossier n'existe pas - le creer
                try
                {
                    int lastSlash = vaultPath.LastIndexOf('/');
                    if (lastSlash <= 0)
                    {
                        // Racine $/ - retourner directement
                        try
                        {
                            var rootFolder = _connection.WebServiceManager.DocumentService.GetFolderByPath("$");
                            _folderCache["$"] = rootFolder;
                            return rootFolder;
                        }
                        catch
                        {
                            return null;
                        }
                    }

                    string parentPath = vaultPath.Substring(0, lastSlash);
                    string folderName = vaultPath.Substring(lastSlash + 1);

                    // Creer le parent recursivement
                    var parentFolder = EnsureVaultFolder(parentPath);
                    if (parentFolder == null) return null;

                    // Creer le dossier
                    var newFolder = _connection.WebServiceManager.DocumentService.AddFolder(folderName, parentFolder.Id, false);
                    _foldersCreated++;
                    _folderCache[vaultPath] = newFolder;
                    return newFolder;
                }
                catch
                {
                    // Peut-etre deja cree - reessayer
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

        // ====================================================================
        // Helpers
        // ====================================================================
        private static List<string> GetFilesToUpload(string rootPath)
        {
            var files = new List<string>();

            try
            {
                foreach (var file in Directory.EnumerateFiles(rootPath, "*.*", SearchOption.AllDirectories))
                {
                    if (!ShouldExclude(file))
                    {
                        files.Add(file);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur scan: {ex.Message}", "ERROR");
            }

            return files;
        }

        private static bool ShouldExclude(string filePath)
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

            // Dossiers exclus
            if (ExcludedFolders.Any(f => filePath.Contains($"\\{f}\\") || filePath.EndsWith($"\\{f}")))
                return true;

            return false;
        }

        private static string GetVaultFolderPath(string localFilePath, string localRoot)
        {
            // Obtenir le dossier du fichier
            string fileDir = Path.GetDirectoryName(localFilePath) ?? "";

            // Convertir C:\Vault\X\Y en $/X/Y
            if (fileDir.StartsWith(LocalVaultRoot, StringComparison.OrdinalIgnoreCase))
            {
                string relativePath = fileDir.Substring(LocalVaultRoot.Length).TrimStart('\\');
                if (string.IsNullOrEmpty(relativePath))
                    return "$";
                return "$/" + relativePath.Replace("\\", "/");
            }

            // Si localRoot different de C:\Vault
            if (fileDir.StartsWith(localRoot, StringComparison.OrdinalIgnoreCase))
            {
                // Calculer le chemin Vault base sur localRoot
                string localRootRelative = localRoot.StartsWith(LocalVaultRoot, StringComparison.OrdinalIgnoreCase)
                    ? localRoot.Substring(LocalVaultRoot.Length).TrimStart('\\')
                    : "";

                string vaultBase = string.IsNullOrEmpty(localRootRelative) 
                    ? "$" 
                    : "$/" + localRootRelative.Replace("\\", "/");

                string relativePath = fileDir.Substring(localRoot.Length).TrimStart('\\');
                if (string.IsNullOrEmpty(relativePath))
                    return vaultBase;
                return vaultBase + "/" + relativePath.Replace("\\", "/");
            }

            return "$";
        }

        private static void Log(string message, string level)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            
            ConsoleColor color = level switch
            {
                "SUCCESS" => ConsoleColor.Green,
                "ERROR" => ConsoleColor.Red,
                "WARNING" => ConsoleColor.Yellow,
                "INFO" => ConsoleColor.Cyan,
                "DEBUG" => ConsoleColor.DarkGray,
                _ => ConsoleColor.White
            };

            string prefix = level switch
            {
                "SUCCESS" => "[+]",
                "ERROR" => "[-]",
                "WARNING" => "[!]",
                "INFO" => "[i]",
                "DEBUG" => "[>]",
                _ => "[>]"
            };

            Console.ForegroundColor = color;
            Console.WriteLine($"[{timestamp}] {prefix} {message}");
            Console.ResetColor();

            // Log fichier
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, 
                    $"VaultBulkUploader_{DateTime.Now:yyyyMMdd}.log");
                File.AppendAllText(logPath, $"[{timestamp}] [{level}] {message}\r\n");
            }
            catch { }
        }

        private static void WaitForExit()
        {
            Console.WriteLine();
            Console.Write("Appuyez sur Entree pour fermer...");
            Console.ReadLine();
        }
    }
}
