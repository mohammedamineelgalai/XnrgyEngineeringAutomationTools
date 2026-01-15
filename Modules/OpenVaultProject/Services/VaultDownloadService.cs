#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XnrgyEngineeringAutomationTools.Modules.OpenVaultProject.Models;
using XnrgyEngineeringAutomationTools.Services;
using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;

namespace XnrgyEngineeringAutomationTools.Modules.OpenVaultProject.Services
{
    /// <summary>
    /// Service pour telecharger et ouvrir des projets depuis Vault
    /// </summary>
    public class VaultDownloadService
    {
        private readonly XnrgyEngineeringAutomationTools.Services.VaultSdkService _vaultService;
        private readonly XnrgyEngineeringAutomationTools.Services.InventorService _inventorService;
        
        // Chemin de base dans Vault pour les projets
        private const string VAULT_PROJECTS_PATH = "$/Engineering/Projects";
        
        // Workspace local
        private readonly string _workspacePath;
        
        public event Action<string, string>? OnProgress;
        public event Action<int, int>? OnFileProgress;

        public VaultDownloadService(
            XnrgyEngineeringAutomationTools.Services.VaultSdkService vaultService,
            XnrgyEngineeringAutomationTools.Services.InventorService inventorService)
        {
            _vaultService = vaultService ?? throw new ArgumentNullException(nameof(vaultService));
            _inventorService = inventorService ?? throw new ArgumentNullException(nameof(inventorService));
            
            // Determiner le workspace depuis Vault settings
            _workspacePath = GetWorkspacePath();
        }

        private string GetWorkspacePath()
        {
            try
            {
                var connection = _vaultService.Connection;
                if (connection != null)
                {
                    var workingFolder = connection.WorkingFoldersManager.GetWorkingFolder("$");
                    return workingFolder.FullPath;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Impossible de recuperer le workspace Vault: {ex.Message}", Logger.LogLevel.WARNING);
            }
            
            // Fallback vers le chemin standard
            return @"C:\Vault";
        }

        /// <summary>
        /// Liste les projets disponibles dans Vault
        /// </summary>
        public List<VaultProjectItem> GetProjects()
        {
            var projects = new List<VaultProjectItem>();
            
            try
            {
                var connection = _vaultService.Connection;
                if (connection == null)
                {
                    Logger.Log("[-] Non connecte a Vault", Logger.LogLevel.ERROR);
                    return projects;
                }

                OnProgress?.Invoke("Chargement des projets...", "INFO");

                // Obtenir le dossier Projects
                var projectsFolder = connection.WebServiceManager.DocumentService.GetFolderByPath(VAULT_PROJECTS_PATH);
                if (projectsFolder == null)
                {
                    Logger.Log($"[-] Dossier non trouve: {VAULT_PROJECTS_PATH}", Logger.LogLevel.ERROR);
                    return projects;
                }

                // Lister les sous-dossiers (projets)
                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(projectsFolder.Id, false);
                if (subFolders != null)
                {
                    foreach (var folder in subFolders)
                    {
                        // Ne garder que les dossiers qui ressemblent a des numeros de projet (ex: 10359)
                        if (folder.Name.All(char.IsDigit) && folder.Name.Length >= 4)
                        {
                            projects.Add(new VaultProjectItem
                            {
                                Name = folder.Name,
                                Path = $"{VAULT_PROJECTS_PATH}/{folder.Name}",
                                Type = "Project",
                                EntityId = folder.Id,
                                LastModified = folder.CreateDate
                            });
                        }
                    }
                }

                OnProgress?.Invoke($"{projects.Count} projets trouves", "SUCCESS");
                Logger.Log($"[+] {projects.Count} projets trouves dans Vault", Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur lors du chargement des projets: {ex.Message}", Logger.LogLevel.ERROR);
                OnProgress?.Invoke($"Erreur: {ex.Message}", "ERROR");
            }

            return projects.OrderByDescending(p => p.Name).ToList();
        }

        /// <summary>
        /// Liste les references d'un projet
        /// </summary>
        public List<VaultProjectItem> GetReferences(string projectPath)
        {
            var references = new List<VaultProjectItem>();
            
            try
            {
                var connection = _vaultService.Connection;
                if (connection == null) return references;

                var folder = connection.WebServiceManager.DocumentService.GetFolderByPath(projectPath);
                if (folder == null) return references;

                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(folder.Id, false);
                if (subFolders != null)
                {
                    foreach (var subFolder in subFolders)
                    {
                        // References: REF01, REF02, etc.
                        if (subFolder.Name.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
                        {
                            references.Add(new VaultProjectItem
                            {
                                Name = subFolder.Name,
                                Path = $"{projectPath}/{subFolder.Name}",
                                Type = "Reference",
                                EntityId = subFolder.Id,
                                LastModified = subFolder.CreateDate
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur lors du chargement des references: {ex.Message}", Logger.LogLevel.ERROR);
            }

            return references.OrderBy(r => r.Name).ToList();
        }

        /// <summary>
        /// Liste les modules d'une reference
        /// </summary>
        public List<VaultProjectItem> GetModules(string referencePath)
        {
            var modules = new List<VaultProjectItem>();
            
            try
            {
                var connection = _vaultService.Connection;
                if (connection == null) return modules;

                var folder = connection.WebServiceManager.DocumentService.GetFolderByPath(referencePath);
                if (folder == null) return modules;

                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(folder.Id, false);
                if (subFolders != null)
                {
                    foreach (var subFolder in subFolders)
                    {
                        // Modules: M01, M02, etc.
                        if (subFolder.Name.StartsWith("M", StringComparison.OrdinalIgnoreCase) && 
                            subFolder.Name.Length >= 2 &&
                            subFolder.Name.Substring(1).All(char.IsDigit))
                        {
                            modules.Add(new VaultProjectItem
                            {
                                Name = subFolder.Name,
                                Path = $"{referencePath}/{subFolder.Name}",
                                Type = "Module",
                                EntityId = subFolder.Id,
                                LastModified = subFolder.CreateDate
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur lors du chargement des modules: {ex.Message}", Logger.LogLevel.ERROR);
            }

            return modules.OrderBy(m => m.Name).ToList();
        }

        /// <summary>
        /// Telecharge recursivement un module depuis Vault et l'ouvre dans Inventor
        /// </summary>
        public async Task<bool> DownloadAndOpenModuleAsync(VaultProjectItem module)
        {
            try
            {
                if (module == null || module.Type != "Module")
                {
                    Logger.Log("[-] Element invalide: doit etre un module", Logger.LogLevel.ERROR);
                    return false;
                }

                OnProgress?.Invoke($"Telechargement de {module.Path}...", "START");
                Logger.Log($"[>] Debut telechargement module: {module.Path}", Logger.LogLevel.INFO);

                var connection = _vaultService.Connection;
                if (connection == null)
                {
                    Logger.Log("[-] Non connecte a Vault", Logger.LogLevel.ERROR);
                    return false;
                }

                // 1. Obtenir la liste de tous les fichiers du module
                var files = await Task.Run(() => GetAllFilesInFolder(module.Path));
                if (files.Count == 0)
                {
                    Logger.Log($"[!] Aucun fichier trouve dans {module.Path}", Logger.LogLevel.WARNING);
                    OnProgress?.Invoke("Aucun fichier trouve", "WARN");
                    return false;
                }

                OnProgress?.Invoke($"{files.Count} fichiers a telecharger", "INFO");
                Logger.Log($"[i] {files.Count} fichiers a telecharger", Logger.LogLevel.INFO);

                // 2. Telecharger les fichiers
                int downloaded = 0;
                int failed = 0;
                string? masterFile = null;

                foreach (var file in files)
                {
                    OnFileProgress?.Invoke(downloaded + failed + 1, files.Count);
                    OnProgress?.Invoke($"Telechargement: {file.Name}", "INFO");

                    try
                    {
                        var localPath = await Task.Run(() => DownloadFile(file));
                        if (!string.IsNullOrEmpty(localPath))
                        {
                            downloaded++;
                            
                            // Identifier le fichier master (.iam ou .ipt principal)
                            if (masterFile == null && 
                                (file.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) ||
                                 file.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase)))
                            {
                                // Le master est generalement celui qui porte le nom du module
                                var moduleName = module.Name;
                                if (file.Name.IndexOf(moduleName, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    masterFile = localPath;
                                }
                            }
                        }
                        else
                        {
                            failed++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"   [-] Erreur: {file.Name} - {ex.Message}", Logger.LogLevel.ERROR);
                        failed++;
                    }
                }

                OnProgress?.Invoke($"Telechargement termine: {downloaded} reussi(s), {failed} echec(s)", 
                    failed == 0 ? "SUCCESS" : "WARN");
                Logger.Log($"[=] Telechargement: {downloaded} reussi(s), {failed} echec(s)", Logger.LogLevel.INFO);

                // 3. Ouvrir le fichier master dans Inventor si trouve
                if (!string.IsNullOrEmpty(masterFile) && File.Exists(masterFile))
                {
                    OnProgress?.Invoke($"Ouverture dans Inventor: {Path.GetFileName(masterFile)}", "START");
                    Logger.Log($"[>] Ouverture du master: {masterFile}", Logger.LogLevel.INFO);

                    bool opened = await Task.Run(() => OpenDocumentInInventor(masterFile));
                    if (opened)
                    {
                        OnProgress?.Invoke("Module ouvert dans Inventor", "SUCCESS");
                        Logger.Log("[+] Module ouvert avec succes dans Inventor", Logger.LogLevel.INFO);
                        return true;
                    }
                    else
                    {
                        OnProgress?.Invoke("Erreur ouverture Inventor", "ERROR");
                        Logger.Log("[-] Erreur lors de l'ouverture dans Inventor", Logger.LogLevel.ERROR);
                    }
                }
                else
                {
                    OnProgress?.Invoke("Fichiers telecharges (pas de master identifie)", "WARN");
                    Logger.Log("[!] Fichiers telecharges mais pas de master identifie", Logger.LogLevel.WARNING);
                }

                return downloaded > 0;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur lors du telechargement: {ex.Message}", Logger.LogLevel.ERROR);
                OnProgress?.Invoke($"Erreur: {ex.Message}", "ERROR");
                return false;
            }
        }

        /// <summary>
        /// Obtient tous les fichiers d'un dossier Vault (recursif)
        /// </summary>
        private List<ACW.File> GetAllFilesInFolder(string vaultPath)
        {
            var allFiles = new List<ACW.File>();
            
            try
            {
                var connection = _vaultService.Connection;
                if (connection == null) return allFiles;

                var folder = connection.WebServiceManager.DocumentService.GetFolderByPath(vaultPath);
                if (folder == null) return allFiles;

                GetFilesRecursive(folder.Id, allFiles);
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur GetAllFilesInFolder: {ex.Message}", Logger.LogLevel.ERROR);
            }

            return allFiles;
        }

        private void GetFilesRecursive(long folderId, List<ACW.File> allFiles)
        {
            var connection = _vaultService.Connection;
            if (connection == null) return;

            try
            {
                // Obtenir les fichiers du dossier
                var files = connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folderId, false);
                if (files != null)
                {
                    foreach (var file in files)
                    {
                        // Filtrer les fichiers inutiles
                        var ext = Path.GetExtension(file.Name).ToLowerInvariant();
                        if (!ext.EndsWith(".bak") && !ext.EndsWith(".old") && 
                            !ext.EndsWith(".tmp") && !ext.EndsWith(".lck") &&
                            !file.Name.StartsWith("~$") && !file.Name.StartsWith("._"))
                        {
                            allFiles.Add(file);
                        }
                    }
                }

                // Recursion dans les sous-dossiers
                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(folderId, false);
                if (subFolders != null)
                {
                    foreach (var subFolder in subFolders)
                    {
                        // Ignorer les dossiers de backup
                        if (!subFolder.Name.Equals("OldVersions", StringComparison.OrdinalIgnoreCase) &&
                            !subFolder.Name.Equals("Backup", StringComparison.OrdinalIgnoreCase))
                        {
                            GetFilesRecursive(subFolder.Id, allFiles);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   [!] Erreur recursion dossier {folderId}: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Telecharge un fichier depuis Vault vers le workspace local
        /// </summary>
        private string? DownloadFile(ACW.File file)
        {
            try
            {
                var connection = _vaultService.Connection;
                if (connection == null) return null;

                // Obtenir le FileIteration pour le telechargement
                var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(connection, file);
                
                // Creer les parametres de telechargement
                var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection, false);
                downloadSettings.AddFileToAcquire(fileIteration, 
                    VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                
                // Telecharger le fichier
                var downloadResult = connection.FileManager.AcquireFiles(downloadSettings);
                
                if (downloadResult.FileResults != null && 
                    downloadResult.FileResults.Any(r => r.LocalPath != null))
                {
                    var localPath = downloadResult.FileResults.First().LocalPath.FullPath;
                    Logger.Log($"   [+] Telecharge: {Path.GetFileName(localPath)}", Logger.LogLevel.DEBUG);
                    return localPath;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   [-] Erreur telechargement {file.Name}: {ex.Message}", Logger.LogLevel.ERROR);
            }

            return null;
        }

        /// <summary>
        /// Retourne le chemin du workspace local
        /// </summary>
        public string GetLocalWorkspacePath() => _workspacePath;

        /// <summary>
        /// Ouvre un document dans Inventor via l'API COM
        /// </summary>
        private bool OpenDocumentInInventor(string filePath)
        {
            try
            {
                if (!_inventorService.IsConnected)
                {
                    Logger.Log("[!] Inventor n'est pas connecte, tentative de connexion...", Logger.LogLevel.WARNING);
                    if (!_inventorService.TryConnect())
                    {
                        Logger.Log("[-] Impossible de se connecter a Inventor", Logger.LogLevel.ERROR);
                        return false;
                    }
                }

                // Obtenir l'instance Inventor via reflexion (acces interne)
                var inventorAppField = _inventorService.GetType()
                    .GetField("_inventorApp", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                
                if (inventorAppField == null)
                {
                    Logger.Log("[-] Impossible d'acceder a l'instance Inventor", Logger.LogLevel.ERROR);
                    return false;
                }

                dynamic? inventorApp = inventorAppField.GetValue(_inventorService);
                if (inventorApp == null)
                {
                    Logger.Log("[-] Instance Inventor null", Logger.LogLevel.ERROR);
                    return false;
                }

                // Ouvrir le document en mode visible
                inventorApp.Documents.Open(filePath, true);
                Logger.Log($"[+] Document ouvert: {Path.GetFileName(filePath)}", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur ouverture document: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }
    }
}
