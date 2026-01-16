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

                // 0. OBLIGATOIRE: Sauvegarder et fermer tous les documents avant switch IPJ
                OnProgress?.Invoke("Sauvegarde des documents ouverts...", "INFO");
                await Task.Run(() => SaveAllAndCloseAllDocuments());

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

                // 2. Telecharger les fichiers - collecter les .iam et .ipj a la racine
                int downloaded = 0;
                int failed = 0;
                string? masterIam = null;
                string? projectIpj = null;
                var rootIamFiles = new List<(string Name, string LocalPath)>();

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
                            
                            // Collecter les fichiers .ipj (projet Inventor)
                            if (file.Name.EndsWith(".ipj", StringComparison.OrdinalIgnoreCase))
                            {
                                // Priorite au .ipj a la racine du module
                                if (projectIpj == null || IsRootLevelFile(file, module.Path))
                                {
                                    projectIpj = localPath;
                                    Logger.Log($"   [i] Projet IPJ trouve: {file.Name}", Logger.LogLevel.DEBUG);
                                }
                            }
                            
                            // Collecter les .iam a la racine pour identifier le master
                            if (file.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase))
                            {
                                if (IsRootLevelFile(file, module.Path))
                                {
                                    rootIamFiles.Add((file.Name, localPath));
                                    Logger.Log($"   [i] IAM racine trouve: {file.Name}", Logger.LogLevel.DEBUG);
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

                // 3. Identifier le fichier master (.iam principal)
                masterIam = IdentifyMasterAssembly(rootIamFiles, module);

                OnProgress?.Invoke($"Telechargement termine: {downloaded} reussi(s), {failed} echec(s)", 
                    failed == 0 ? "SUCCESS" : "WARN");
                Logger.Log($"[=] Telechargement: {downloaded} reussi(s), {failed} echec(s)", Logger.LogLevel.INFO);

                // 4. Switch vers le projet IPJ du module
                // L'IPJ est dans le MEME dossier que le master .iam (format: 123450101.ipj)
                string? ipjToUse = projectIpj;
                
                // Si pas d'IPJ trouve dans le telechargement, chercher dans le dossier du module
                if (string.IsNullOrEmpty(ipjToUse) || !File.Exists(ipjToUse))
                {
                    if (!string.IsNullOrEmpty(masterIam))
                    {
                        string? moduleDir = Path.GetDirectoryName(masterIam);
                        if (!string.IsNullOrEmpty(moduleDir))
                        {
                            ipjToUse = FindModuleIpj(moduleDir);
                        }
                    }
                }
                
                if (!string.IsNullOrEmpty(ipjToUse) && File.Exists(ipjToUse))
                {
                    OnProgress?.Invoke($"Switch vers projet: {Path.GetFileName(ipjToUse)}", "INFO");
                    Logger.Log($"[>] Switch vers IPJ: {ipjToUse}", Logger.LogLevel.INFO);
                    await Task.Run(() => SwitchToProjectIpj(ipjToUse));
                }
                else
                {
                    Logger.Log("[!] Aucun IPJ trouve dans le module, ouverture sans switch...", Logger.LogLevel.WARNING);
                }

                // 5. Ouvrir le fichier master dans Inventor si trouve
                if (!string.IsNullOrEmpty(masterIam) && File.Exists(masterIam))
                {
                    OnProgress?.Invoke($"Ouverture dans Inventor: {Path.GetFileName(masterIam)}", "START");
                    Logger.Log($"[>] Ouverture du master: {masterIam}", Logger.LogLevel.INFO);

                    bool opened = await Task.Run(() => OpenDocumentInInventor(masterIam));
                    if (opened)
                    {
                        // 6. Appliquer le nettoyage et zoom all
                        OnProgress?.Invoke("Preparation de la vue...", "INFO");
                        await Task.Run(() => PrepareViewAfterOpen());
                        
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
                    Logger.Log($"[i] IAM racine trouves: {rootIamFiles.Count}", Logger.LogLevel.DEBUG);
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

        /// <summary>
        /// Verifie si un fichier est a la racine du module (pas dans un sous-dossier)
        /// </summary>
        private bool IsRootLevelFile(ACW.File file, string modulePath)
        {
            try
            {
                // Comparer le chemin du dossier du fichier avec le chemin du module
                var connection = _vaultService.Connection;
                if (connection == null) return false;

                var folder = connection.WebServiceManager.DocumentService.GetFolderById(file.FolderId);
                if (folder == null) return false;

                // Le fichier est a la racine si son dossier parent est le module
                return folder.FullName.Equals(modulePath, StringComparison.OrdinalIgnoreCase) ||
                       folder.FullName.Replace("$", "").TrimStart('/').Equals(
                           modulePath.Replace("$", "").TrimStart('/'), StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Identifie le fichier master .iam parmi les fichiers a la racine
        /// Logique: Le master est generalement le plus gros .iam ou celui qui a un numero de projet
        /// </summary>
        private string? IdentifyMasterAssembly(List<(string Name, string LocalPath)> rootIamFiles, VaultProjectItem module)
        {
            if (rootIamFiles.Count == 0) return null;
            
            // Si un seul .iam, c'est le master
            if (rootIamFiles.Count == 1)
            {
                Logger.Log($"[+] Master identifie (unique): {rootIamFiles[0].Name}", Logger.LogLevel.INFO);
                return rootIamFiles[0].LocalPath;
            }

            // Extraire le numero de projet du chemin (ex: $/Engineering/Projects/12345/REF01/M01 -> 12345)
            var pathParts = module.Path.Split('/');
            string? projectNumber = null;
            for (int i = 0; i < pathParts.Length; i++)
            {
                if (pathParts[i].Equals("Projects", StringComparison.OrdinalIgnoreCase) && i + 1 < pathParts.Length)
                {
                    projectNumber = pathParts[i + 1];
                    break;
                }
            }

            // Chercher un .iam qui contient le numero de projet (ex: 123450101.iam)
            if (!string.IsNullOrEmpty(projectNumber))
            {
                foreach (var iam in rootIamFiles)
                {
                    if (iam.Name.StartsWith(projectNumber, StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.Log($"[+] Master identifie (numero projet {projectNumber}): {iam.Name}", Logger.LogLevel.INFO);
                        return iam.LocalPath;
                    }
                }
            }

            // Sinon, prendre le plus gros fichier .iam
            var largestIam = rootIamFiles
                .Where(f => File.Exists(f.LocalPath))
                .OrderByDescending(f => new FileInfo(f.LocalPath).Length)
                .FirstOrDefault();

            if (!string.IsNullOrEmpty(largestIam.LocalPath))
            {
                Logger.Log($"[+] Master identifie (plus gros fichier): {largestIam.Name}", Logger.LogLevel.INFO);
                return largestIam.LocalPath;
            }

            // Fallback: premier .iam
            Logger.Log($"[+] Master identifie (premier): {rootIamFiles[0].Name}", Logger.LogLevel.INFO);
            return rootIamFiles[0].LocalPath;
        }

        /// <summary>
        /// Sauvegarde et ferme tous les documents ouverts dans Inventor
        /// OBLIGATOIRE avant de switcher l'IPJ
        /// </summary>
        private void SaveAllAndCloseAllDocuments()
        {
            try
            {
                var inventorAppField = _inventorService.GetType()
                    .GetField("_inventorApp", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                
                if (inventorAppField == null) return;

                dynamic? inventorApp = inventorAppField.GetValue(_inventorService);
                if (inventorApp == null) return;

                int docCount = inventorApp.Documents.Count;
                if (docCount == 0)
                {
                    Logger.Log("[i] Aucun document ouvert", Logger.LogLevel.DEBUG);
                    return;
                }

                Logger.Log($"[>] Sauvegarde et fermeture de {docCount} document(s)...", Logger.LogLevel.INFO);
                OnProgress?.Invoke($"Sauvegarde de {docCount} document(s)...", "INFO");

                // Activer le mode silencieux
                bool origSilent = inventorApp.SilentOperation;
                bool origUserDisabled = inventorApp.UserInterfaceManager.UserInteractionDisabled;

                try
                {
                    inventorApp.SilentOperation = true;
                    inventorApp.UserInterfaceManager.UserInteractionDisabled = true;

                    // Sauvegarder tous les documents modifies
                    foreach (dynamic doc in inventorApp.Documents)
                    {
                        try
                        {
                            if (doc.IsModifiable && doc.Dirty)
                            {
                                doc.Save();
                                Logger.Log($"   [+] Sauvegarde: {System.IO.Path.GetFileName(doc.FullFileName)}", Logger.LogLevel.DEBUG);
                            }
                        }
                        catch { }
                    }

                    // Fermer tous les documents (avec sauvegarde au cas ou)
                    while (inventorApp.Documents.Count > 0)
                    {
                        try
                        {
                            inventorApp.Documents[1].Close(true); // true = skip save prompt (deja sauvegarde)
                        }
                        catch { break; }
                    }

                    Logger.Log($"[+] {docCount} document(s) ferme(s)", Logger.LogLevel.INFO);
                }
                finally
                {
                    // Restaurer les modes
                    inventorApp.SilentOperation = origSilent;
                    inventorApp.UserInterfaceManager.UserInteractionDisabled = origUserDisabled;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur SaveAllAndClose: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Trouve le fichier IPJ du module (format: 123450101.ipj ou 12345-01-M01.ipj)
        /// L'IPJ est TOUJOURS dans le meme dossier que le fichier .iam master
        /// </summary>
        private string? FindModuleIpj(string moduleFolder)
        {
            try
            {
                if (!Directory.Exists(moduleFolder))
                {
                    Logger.Log($"[-] Dossier module inexistant: {moduleFolder}", Logger.LogLevel.ERROR);
                    return null;
                }

                // Chercher tous les .ipj dans le dossier du module
                var ipjFiles = Directory.GetFiles(moduleFolder, "*.ipj", SearchOption.TopDirectoryOnly);
                
                if (ipjFiles.Length == 0)
                {
                    Logger.Log($"[!] Aucun IPJ trouve dans: {moduleFolder}", Logger.LogLevel.WARNING);
                    return null;
                }

                if (ipjFiles.Length == 1)
                {
                    Logger.Log($"[+] IPJ trouve: {Path.GetFileName(ipjFiles[0])}", Logger.LogLevel.INFO);
                    return ipjFiles[0];
                }

                // Si plusieurs IPJ, prendre celui qui correspond au format attendu
                // Format nouveau: 123450101.ipj (chiffres uniquement)
                // Format ancien: 12345-01-M01.ipj
                foreach (var ipj in ipjFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(ipj);
                    // Priorite au format nouveau (chiffres uniquement, 8-10 caracteres)
                    if (fileName.All(char.IsDigit) && fileName.Length >= 8)
                    {
                        Logger.Log($"[+] IPJ Master trouve (format nouveau): {Path.GetFileName(ipj)}", Logger.LogLevel.INFO);
                        return ipj;
                    }
                }

                // Sinon prendre le premier
                Logger.Log($"[+] IPJ trouve: {Path.GetFileName(ipjFiles[0])}", Logger.LogLevel.INFO);
                return ipjFiles[0];
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur recherche IPJ: {ex.Message}", Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// Switch vers le projet IPJ specifie
        /// Logique identique a EquipmentCopyDesignService.SwitchProject
        /// </summary>
        private bool SwitchToProjectIpj(string ipjPath)
        {
            try
            {
                if (!File.Exists(ipjPath))
                {
                    Logger.Log($"[-] Fichier IPJ inexistant: {ipjPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                var inventorAppField = _inventorService.GetType()
                    .GetField("_inventorApp", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                
                if (inventorAppField == null)
                {
                    Logger.Log("[-] Impossible d'acceder a Inventor", Logger.LogLevel.ERROR);
                    return false;
                }

                dynamic? inventorApp = inventorAppField.GetValue(_inventorService);
                if (inventorApp == null)
                {
                    Logger.Log("[-] Instance Inventor null", Logger.LogLevel.ERROR);
                    return false;
                }

                var designProjectManager = inventorApp.DesignProjectManager;
                var projectsCollection = designProjectManager.DesignProjects;

                // 1. Verifier si ce projet est deja actif
                var activeProject = designProjectManager.ActiveDesignProject;
                if (activeProject != null && 
                    activeProject.FullFileName.Equals(ipjPath, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Log($"[i] IPJ deja actif: {Path.GetFileName(ipjPath)}", Logger.LogLevel.INFO);
                    return true;
                }

                // 2. Chercher si le projet existe deja dans la collection
                dynamic? targetProject = null;
                foreach (dynamic project in projectsCollection)
                {
                    if (project.FullFileName.Equals(ipjPath, StringComparison.OrdinalIgnoreCase))
                    {
                        targetProject = project;
                        break;
                    }
                }

                // 3. Si pas trouve, l'ajouter a la collection
                if (targetProject == null)
                {
                    Logger.Log($"[i] Chargement du projet: {Path.GetFileName(ipjPath)}", Logger.LogLevel.DEBUG);
                    targetProject = projectsCollection.AddExisting(ipjPath);
                }

                // 4. Activer le projet avec delai (comme CopyDesign)
                if (targetProject != null)
                {
                    targetProject.Activate();
                    System.Threading.Thread.Sleep(1000); // Attendre 1s comme dans CopyDesign
                    Logger.Log($"[+] IPJ active: {Path.GetFileName(ipjPath)}", Logger.LogLevel.INFO);
                    return true;
                }
                else
                {
                    Logger.Log("[-] Impossible de charger le projet IPJ", Logger.LogLevel.ERROR);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur switch IPJ: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Ouvre un module local (sans telechargement Vault)
        /// SEQUENCE: SaveAll + CloseAll + Switch IPJ + Open + PrepareView
        /// </summary>
        public bool OpenLocalModule(string masterIamPath)
        {
            try
            {
                Logger.Log($"[>] Ouverture module local: {Path.GetFileName(masterIamPath)}", Logger.LogLevel.INFO);
                
                // 1. OBLIGATOIRE: Sauvegarder et fermer tous les documents avant switch IPJ
                OnProgress?.Invoke("Sauvegarde des documents ouverts...", "INFO");
                SaveAllAndCloseAllDocuments();
                
                // 2. Trouver l'IPJ du module (dans le MEME dossier que le .iam)
                // Format: 123450101.ipj ou 12345-01-M01.ipj
                string? moduleFolder = Path.GetDirectoryName(masterIamPath);
                if (string.IsNullOrEmpty(moduleFolder))
                {
                    Logger.Log("[-] Impossible de determiner le dossier du module", Logger.LogLevel.ERROR);
                    return false;
                }
                
                OnProgress?.Invoke("Switch vers le projet IPJ...", "INFO");
                
                // 3. Trouver et activer l'IPJ du module
                string? ipjPath = FindModuleIpj(moduleFolder);
                if (!string.IsNullOrEmpty(ipjPath))
                {
                    if (!SwitchToProjectIpj(ipjPath))
                    {
                        Logger.Log("[!] Echec switch IPJ, tentative d'ouverture quand meme...", Logger.LogLevel.WARNING);
                    }
                }
                else
                {
                    Logger.Log("[!] Aucun IPJ trouve dans le module, ouverture sans switch...", Logger.LogLevel.WARNING);
                }

                OnProgress?.Invoke($"Ouverture: {Path.GetFileName(masterIamPath)}", "INFO");
                
                // 4. Ouvrir le fichier dans Inventor (mode visible)
                var inventorAppField = _inventorService.GetType()
                    .GetField("_inventorApp", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                
                if (inventorAppField == null)
                {
                    Logger.Log("[-] Impossible d'acceder a Inventor", Logger.LogLevel.ERROR);
                    return false;
                }

                dynamic? inventorApp = inventorAppField.GetValue(_inventorService);
                if (inventorApp == null)
                {
                    Logger.Log("[-] Inventor n'est pas disponible", Logger.LogLevel.ERROR);
                    return false;
                }

                // Ouvrir le document (mode visible)
                dynamic doc = inventorApp.Documents.Open(masterIamPath, true);
                if (doc == null)
                {
                    Logger.Log("[-] Echec ouverture du document", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"[+] Document ouvert: {Path.GetFileName(masterIamPath)}", Logger.LogLevel.INFO);

                OnProgress?.Invoke("Preparation de la vue...", "INFO");
                
                // 5. Preparer la vue (cacher references, vue ISO, zoom)
                System.Threading.Thread.Sleep(500);
                PrepareViewAfterOpen();

                OnProgress?.Invoke("Module local ouvert avec succes", "SUCCESS");
                Logger.Log("[+] Module local ouvert avec succes", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogException("OpenLocalModule", ex, Logger.LogLevel.ERROR);
                OnProgress?.Invoke($"Erreur: {ex.Message}", "ERROR");
                return false;
            }
        }

        /// <summary>
        /// Prepare la vue apres ouverture: cache les elements de reference, vue ISO, zoom all
        /// </summary>
        private void PrepareViewAfterOpen()
        {
            try
            {
                var inventorAppField = _inventorService.GetType()
                    .GetField("_inventorApp", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                
                if (inventorAppField == null) return;

                dynamic? inventorApp = inventorAppField.GetValue(_inventorService);
                if (inventorApp == null) return;

                // Attendre que le document soit pret
                System.Threading.Thread.Sleep(1000);

                dynamic? activeDoc = inventorApp.ActiveDocument;
                if (activeDoc == null) return;

                // 1. Cacher les elements de reference via ObjectVisibility (methode globale efficace)
                try
                {
                    dynamic asmDocVis = activeDoc;
                    // Plans de travail utilisateur
                    asmDocVis.ObjectVisibility.UserWorkPlanes = false;
                    // Axes de travail utilisateur
                    asmDocVis.ObjectVisibility.UserWorkAxes = false;
                    // Points de travail utilisateur
                    asmDocVis.ObjectVisibility.UserWorkPoints = false;
                    // Plans d'origine
                    asmDocVis.ObjectVisibility.OriginWorkPlanes = false;
                    // Axes d'origine
                    asmDocVis.ObjectVisibility.OriginWorkAxes = false;
                    // Points d'origine
                    asmDocVis.ObjectVisibility.OriginWorkPoints = false;
                    // Esquisses 2D
                    asmDocVis.ObjectVisibility.Sketches = false;
                    // Esquisses 3D
                    asmDocVis.ObjectVisibility.Sketches3D = false;
                    
                    Logger.Log("  [+] ObjectVisibility: References et esquisses cachees (global)", Logger.LogLevel.DEBUG);
                }
                catch (Exception ex)
                {
                    Logger.Log($"  [!] ObjectVisibility: {ex.Message}", Logger.LogLevel.DEBUG);
                    
                    // Fallback: methode recursive
                    try
                    {
                        if (activeDoc.DocumentType == 12290) // kAssemblyDocumentObject
                        {
                            HideWorkFeaturesRecursive(activeDoc);
                        }
                        else if (activeDoc.DocumentType == 12289) // kPartDocumentObject
                        {
                            HideWorkFeatures(activeDoc.ComponentDefinition);
                        }
                        Logger.Log("  [+] Elements de reference caches (fallback)", Logger.LogLevel.DEBUG);
                    }
                    catch { }
                }

                // 2. Vue isometrique via commande native (comme CreateModule)
                try
                {
                    dynamic cmdManager = inventorApp.CommandManager;
                    dynamic controlDefs = cmdManager.ControlDefinitions;
                    dynamic cmdIso = controlDefs["AppIsometricViewCmd"];
                    cmdIso.Execute();
                    Logger.Log("  [+] Vue ISO appliquee (commande native)", Logger.LogLevel.DEBUG);
                }
                catch (Exception ex)
                {
                    // Fallback: via Camera
                    try
                    {
                        dynamic activeView = inventorApp.ActiveView;
                        if (activeView != null)
                        {
                            dynamic camera = activeView.Camera;
                            camera.ViewOrientationType = 10764; // kIsoTopRightViewOrientation
                            camera.Apply();
                            Logger.Log("  [+] Vue isometrique appliquee (camera)", Logger.LogLevel.DEBUG);
                        }
                    }
                    catch
                    {
                        Logger.Log($"  [!] Vue ISO: {ex.Message}", Logger.LogLevel.DEBUG);
                    }
                }

                // 3. Zoom All
                try
                {
                    dynamic cmdManager = inventorApp.CommandManager;
                    dynamic controlDefs = cmdManager.ControlDefinitions;
                    dynamic cmdZoom = controlDefs["AppZoomAllCmd"];
                    cmdZoom.Execute();
                    Logger.Log("  [+] Zoom All applique", Logger.LogLevel.DEBUG);
                }
                catch (Exception ex)
                {
                    try
                    {
                        dynamic activeView = inventorApp.ActiveView;
                        activeView?.Fit();
                        Logger.Log("  [+] Zoom All (Fit fallback)", Logger.LogLevel.DEBUG);
                    }
                    catch
                    {
                        Logger.Log($"  [!] Zoom All: {ex.Message}", Logger.LogLevel.DEBUG);
                    }
                }

                OnProgress?.Invoke("Vue preparee (ISO + Zoom All)", "SUCCESS");
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur preparation vue: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Cache les WorkFeatures dans un composant (fallback si ObjectVisibility ne fonctionne pas)
        /// </summary>
        private void HideWorkFeatures(dynamic compDef)
        {
            try
            {
                foreach (dynamic wp in compDef.WorkPlanes)
                {
                    try { wp.Visible = false; } catch { }
                }
                foreach (dynamic wa in compDef.WorkAxes)
                {
                    try { wa.Visible = false; } catch { }
                }
                foreach (dynamic wpt in compDef.WorkPoints)
                {
                    try { wpt.Visible = false; } catch { }
                }
                // Esquisses 2D
                try
                {
                    foreach (dynamic sk in compDef.Sketches)
                    {
                        try { sk.Visible = false; } catch { }
                    }
                }
                catch { }
                // Esquisses 3D
                try
                {
                    foreach (dynamic sk3d in compDef.Sketches3D)
                    {
                        try { sk3d.Visible = false; } catch { }
                    }
                }
                catch { }
            }
            catch { }
        }

        /// <summary>
        /// Cache les WorkFeatures recursivement dans un assemblage (fallback)
        /// </summary>
        private void HideWorkFeaturesRecursive(dynamic asmDoc)
        {
            try
            {
                var compDef = asmDoc.ComponentDefinition;
                HideWorkFeatures(compDef);

                foreach (dynamic occ in compDef.Occurrences)
                {
                    try
                    {
                        // Cacher dans les sous-composants
                        if (occ.DefinitionDocumentType == 12290) // Assembly
                        {
                            HideWorkFeaturesRecursive(occ.Definition.Document);
                        }
                        else
                        {
                            HideWorkFeatures(occ.Definition);
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
