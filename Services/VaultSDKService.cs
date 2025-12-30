#nullable enable
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;
using ACET = Autodesk.Connectivity.Explorer.ExtensibilityTools;
using XnrgyEngineeringAutomationTools.Models;

using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Structure pour stocker les proprietes e appliquer en differe (apres Job Processor)
    /// </summary>
    public class PendingPropertyUpdate
    {
        public long FileMasterId { get; set; }
        public string? ProjectNumber { get; set; }
        public string? Reference { get; set; }
        public string? Module { get; set; }
        public string FileName { get; set; } = "";
        public int RetryCount { get; set; } = 0;
    }

    public class VaultSdkService
    {
        private VDF.Vault.Currency.Connections.Connection? _connection;
        private readonly Dictionary<string, long> _propertyDefIds = new();
        private long? _baseCategoryId = null; // ID de la Category "Base"
        
        // Proprietes de connexion
        public string? VaultName { get; private set; }
        public string? UserName { get; private set; }
        public string? ServerName { get; private set; }
        
        // -------------------------------------------------------------------------------
        // FILE D'ATTENTE pour les proprietes a appliquer apres que le Job Processor ait fini
        // -------------------------------------------------------------------------------
        private readonly List<PendingPropertyUpdate> _pendingPropertyUpdates = new();
        private readonly object _pendingLock = new object();

        public bool IsConnected => _connection != null;
        
        /// <summary>
        /// Nombre de proprietes en attente d'application
        /// </summary>
        public int PendingPropertyCount 
        { 
            get 
            { 
                lock (_pendingLock) 
                { 
                    return _pendingPropertyUpdates.Count; 
                } 
            } 
        }

        /// <summary>
        /// Ajoute une mise e jour de proprietes e la file d'attente (pour application apres Job Processor)
        /// </summary>
        public void QueuePropertyUpdate(long fileMasterId, string? projectNumber, string? reference, string? module, string fileName)
        {
            lock (_pendingLock)
            {
                _pendingPropertyUpdates.Add(new PendingPropertyUpdate
                {
                    FileMasterId = fileMasterId,
                    ProjectNumber = projectNumber,
                    Reference = reference,
                    Module = module,
                    FileName = fileName
                });
                Logger.Log($"   [+] Proprietes ajoutees a la file d'attente: {fileName} (MasterId: {fileMasterId})", Logger.LogLevel.DEBUG);
            }
        }

        /// <summary>
        /// Vide la file d'attente des proprietes
        /// </summary>
        public void ClearPendingPropertyUpdates()
        {
            lock (_pendingLock)
            {
                _pendingPropertyUpdates.Clear();
                Logger.Log($"   [OK] File d'attente des proprietes videe", Logger.LogLevel.DEBUG);
            }
        }

        /// <summary>
        /// Applique toutes les proprietes en attente (a appeler APRES tous les uploads)
        /// Le Job Processor aura eu le temps de traiter tous les fichiers
        /// </summary>
        /// <param name="waitBeforeStart">Delai initial d'attente en secondes (par defaut 0s - pas d'attente)</param>
        /// <returns>(succes, echecs)</returns>
        public (int successCount, int failedCount) ApplyPendingPropertyUpdates(int waitBeforeStart = 0)
        {
            List<PendingPropertyUpdate> updates;
            lock (_pendingLock)
            {
                if (_pendingPropertyUpdates.Count == 0)
                {
                    Logger.Log("   [i] Aucune propriete en attente", Logger.LogLevel.INFO);
                    return (0, 0);
                }
                
                updates = new List<PendingPropertyUpdate>(_pendingPropertyUpdates);
                _pendingPropertyUpdates.Clear();
            }

            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);
            Logger.Log($"[>] APPLICATION DES PROPRIETES EN BATCH ({updates.Count} fichiers)", Logger.LogLevel.INFO);
            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);
            
            // Attendre que le Job Processor ait fini avec les fichiers
            if (waitBeforeStart > 0)
            {
                Logger.Log($"   [~] Attente de {waitBeforeStart}s pour le Job Processor...", Logger.LogLevel.INFO);
                System.Threading.Thread.Sleep(waitBeforeStart * 1000);
            }

            int successCount = 0;
            int failedCount = 0;

            foreach (var update in updates)
            {
                Logger.Log($"   [>] [{successCount + failedCount + 1}/{updates.Count}] {update.FileName}...", Logger.LogLevel.INFO);
                
                bool success = SetFilePropertiesDirectly(update.FileMasterId, update.ProjectNumber, update.Reference, update.Module);
                
                if (success)
                {
                    successCount++;
                    Logger.Log($"      [+] Proprietes appliquees", Logger.LogLevel.INFO);
                }
                else
                {
                    failedCount++;
                    Logger.Log($"      [-] Echec (Job Processor occupe ou autre erreur)", Logger.LogLevel.WARNING);
                }
            }

            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);
            Logger.Log($"[=] RESULTAT: {successCount} reussi(s), {failedCount} echec(s)", Logger.LogLevel.INFO);
            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);

            return (successCount, failedCount);
        }

        /// <summary>
        /// Soumet un Job de synchronisation des proprietes (Vault -> iProperties) via le Job Processor.
        /// Le Job Processor doit etre configure avec le handler "Autodesk.Vault.SyncProperties".
        /// </summary>
        /// <param name="fileVersionId">ID de la version du fichier (FileId, pas MasterId)</param>
        /// <param name="fileName">Nom du fichier (pour le log)</param>
        /// <returns>True si le job a ete soumis avec succes</returns>
        public bool SubmitSyncPropertiesJob(long fileVersionId, string fileName)
        {
            if (_connection == null)
            {
                Logger.Log($"   [!] [JOB] Non connecte a Vault", Logger.LogLevel.WARNING);
                return false;
            }

            try
            {
                // Le type de job pour synchroniser les proprietes Vault -> iProperties
                // C'est le handler standard de Vault pour "Property Sync"
                const string JOB_TYPE = "Autodesk.Vault.SyncProperties";
                
                // Le job attend FileVersionId (pas FileMasterId)
                // Erreur si on utilise FileMasterId: "Missing required parameter. Requires either FileVersionIds or FileVersionId."
                var jobParams = new ACW.JobParam[]
                {
                    new ACW.JobParam { Name = "FileVersionId", Val = fileVersionId.ToString() }
                };

                // Soumettre le job avec priorite normale (10)
                var job = _connection.WebServiceManager.JobService.AddJob(
                    JOB_TYPE,
                    $"Sync properties for {fileName}",
                    jobParams,
                    10  // Priorite (1 = haute, 10 = normale)
                );

                if (job != null && job.Id > 0)
                {
                    Logger.Log($"   [+] [JOB] Sync properties soumis (JobId: {job.Id}) pour {fileName}", Logger.LogLevel.INFO);
                    return true;
                }
                else
                {
                    Logger.Log($"   [!] [JOB] Job retourne null ou ID invalide", Logger.LogLevel.WARNING);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   [!] [JOB] Erreur soumission: {ex.Message}", Logger.LogLevel.WARNING);
                // Ne pas faire echouer le process principal si le job echoue
                return false;
            }
        }

        public bool Connect(string server, string vaultName, string username, string password)
        {
            try
            {
                Logger.Log($"[>] Tentative de connexion a Vault...", Logger.LogLevel.INFO);
                Logger.Log($"   Serveur: {server}", Logger.LogLevel.DEBUG);
                Logger.Log($"   Vault: {vaultName}", Logger.LogLevel.DEBUG);
                Logger.Log($"   Utilisateur: {username}", Logger.LogLevel.DEBUG);

                var result = VDF.Vault.Library.ConnectionManager.LogIn(
                    server,
                    vaultName,
                    username,
                    password,
                    VDF.Vault.Currency.Connections.AuthenticationFlags.Standard,
                    null
                );

                if (result.Success)
                {
                    _connection = result.Connection;
                    
                    // Stocker les infos de connexion
                    VaultName = vaultName;
                    UserName = username;
                    ServerName = server;
                    
                    Logger.Log($"[+] Connecte au Vault '{vaultName}' sur '{server}'", Logger.LogLevel.INFO);
                    Logger.Log($"   Utilisateur: {username}", Logger.LogLevel.INFO);
                    Logger.Log($"   Dossier racine: {_connection.FolderManager.RootFolder.FullName}", Logger.LogLevel.INFO);
                    
                    // Log des versions pour diagnostiquer les incompatibilites
                    try
                    {
                        // Version du SDK client (via Assembly)
                        var sdkAssembly = typeof(VDF.Vault.Currency.Connections.Connection).Assembly;
                        var sdkVersion = sdkAssembly.GetName().Version;
                        Logger.Log($"   [i] SDK Client Version: {sdkVersion}", Logger.LogLevel.INFO);
                        
                        // Version du serveur (via WebServiceManager)
                        if (_connection.WebServiceManager != null)
                        {
                            try
                            {
                                // Essayer d'obtenir la version du serveur via AdminService
                                var adminService = _connection.WebServiceManager.AdminService;
                                if (adminService != null && adminService.Session != null)
                                {
                                    // La session contient parfois des infos de version
                                    Logger.Log($"   [i] Serveur Vault: Informations de session disponibles", Logger.LogLevel.DEBUG);
                                }
                            }
                            catch (Exception verEx)
                            {
                                Logger.Log($"   [!] Impossible de recuperer la version serveur: {verEx.Message}", Logger.LogLevel.DEBUG);
                            }
                        }
                        
                        // Version de l'application
                        var appAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                        var appVersion = appAssembly.GetName().Version;
                        Logger.Log($"   [i] Application Version: {appVersion}", Logger.LogLevel.DEBUG);
                    }
                    catch (Exception verEx)
                    {
                        Logger.Log($"   [!] Erreur lors de la recuperation des versions: {verEx.Message}", Logger.LogLevel.DEBUG);
                    }
                    
                    LoadPropertyDefinitions();
                    LoadCategories();
                    
                    // Verifier si writeback active (info seulement)
                    try
                    {
                        bool writebackEnabled = _connection?.WebServiceManager?.DocumentService
                            ?.GetEnableItemPropertyWritebackToFiles() ?? false;
                        Logger.Log($"   Writeback vers fichiers: {(writebackEnabled ? "[+] ACTIVE" : "[!] DESACTIVE")}", 
                            Logger.LogLevel.INFO);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"   [!] Impossible de verifier le writeback: {ex.Message}", Logger.LogLevel.DEBUG);
                    }
                    
                    // NOTE: ListRecentJobs() supprime - etait du diagnostic non necessaire
                    
                    return true;
                }
                else
                {
                    string errorMsg = "Echec d'authentification";
                    if (result.ErrorMessages != null && result.ErrorMessages.Count > 0)
                    {
                        errorMsg = result.ErrorMessages.First().Value;
                    }
                    Logger.Log($"[-] Echec connexion: {errorMsg}", Logger.LogLevel.ERROR);
                    
                    if (result.ErrorMessages != null)
                    {
                        foreach (var err in result.ErrorMessages)
                        {
                            Logger.Log($"   Error [{err.Key}]: {err.Value}", Logger.LogLevel.ERROR);
                        }
                    }
                    
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("Connect", ex, Logger.LogLevel.FATAL);
                _connection = null;
                return false;
            }
        }

        /// <summary>
        /// Verifie si l'utilisateur actuellement connecte est un administrateur Vault.
        /// Verifie via l'API Vault:
        /// 1. Si l'utilisateur a le Role "Administrator" assigne directement
        /// 2. Si l'utilisateur appartient au groupe "Admin_Designer" (groupe admin XNRGY)
        /// </summary>
        /// <returns>True si l'utilisateur est administrateur Vault, false sinon</returns>
        public bool IsCurrentUserAdmin()
        {
            try
            {
                if (!IsConnected || _connection == null)
                    return false;

                var adminService = _connection.WebServiceManager.AdminService;
                if (adminService?.Session?.User == null)
                {
                    Logger.Log("[!] Session AdminService non disponible", Logger.LogLevel.DEBUG);
                    return false;
                }

                // Obtenir l'ID de l'utilisateur connecte
                long userId = adminService.Session.User.Id;
                // Utiliser le login name (mohammedamine.elgalai) pas le display name
                string userLoginName = adminService.Session.User.Name ?? UserName ?? "Unknown";
                
                Logger.Log($"[?] Verification admin pour: '{userLoginName}' (ID: {userId})", Logger.LogLevel.DEBUG);

                // Methode 1: Verifier si l'utilisateur est un utilisateur systeme (ex: Administrator built-in)
                if (adminService.Session.User.IsSys)
                {
                    Logger.Log($"[+] Utilisateur '{userLoginName}' est un utilisateur systeme (admin)", Logger.LogLevel.DEBUG);
                    return true;
                }

                // Methode 2: Obtenir UserInfo avec les Roles ET Groupes de l'utilisateur via l'API Vault
                try
                {
                    var userInfo = adminService.GetUserInfoByUserId(userId);
                    
                    // --- VERIFICATION DES ROLES ---
                    if (userInfo?.Roles != null && userInfo.Roles.Length > 0)
                    {
                        Logger.Log($"[i] Roles de '{userLoginName}':", Logger.LogLevel.DEBUG);
                        foreach (var role in userInfo.Roles)
                        {
                            Logger.Log($"    - '{role.Name}' (Id: {role.Id})", Logger.LogLevel.DEBUG);
                        }

                        // Verifier si l'utilisateur a le role "Administrator" assigne directement
                        bool hasAdminRole = userInfo.Roles.Any(r => 
                            r.Name.Equals("Administrator", StringComparison.OrdinalIgnoreCase));

                        if (hasAdminRole)
                        {
                            Logger.Log($"[+] Utilisateur '{userLoginName}' a le role Administrator", Logger.LogLevel.INFO);
                            return true;
                        }
                    }

                    // --- VERIFICATION DES GROUPES ---
                    if (userInfo?.Groups != null && userInfo.Groups.Length > 0)
                    {
                        Logger.Log($"[i] Groupes de '{userLoginName}':", Logger.LogLevel.DEBUG);
                        foreach (var group in userInfo.Groups)
                        {
                            Logger.Log($"    - '{group.Name}' (IsSys: {group.IsSys})", Logger.LogLevel.DEBUG);
                        }

                        // Groupes admin reconnus:
                        // 1. "Administrators" - groupe systeme Vault
                        // 2. "Admin_Designer" - groupe personnalise XNRGY pour les admins/designers seniors
                        var adminGroupNames = new[] { "Administrators", "Admin_Designer" };

                        bool isInAdminGroup = userInfo.Groups.Any(g => 
                            adminGroupNames.Any(adminName => g.Name.Equals(adminName, StringComparison.OrdinalIgnoreCase)));

                        if (isInAdminGroup)
                        {
                            var matchedGroup = userInfo.Groups.First(g => 
                                adminGroupNames.Any(adminName => g.Name.Equals(adminName, StringComparison.OrdinalIgnoreCase)));
                            Logger.Log($"[+] Utilisateur '{userLoginName}' est membre du groupe admin '{matchedGroup.Name}'", Logger.LogLevel.INFO);
                            return true;
                        }
                    }
                    else
                    {
                        Logger.Log($"[i] Utilisateur '{userLoginName}' n'appartient a aucun groupe", Logger.LogLevel.DEBUG);
                    }
                }
                catch (Exception apiEx)
                {
                    Logger.Log($"[!] Erreur API GetUserInfoByUserId: {apiEx.Message}", Logger.LogLevel.DEBUG);
                }

                Logger.Log($"[i] Utilisateur '{userLoginName}' n'est pas administrateur Vault", Logger.LogLevel.DEBUG);
                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification admin: {ex.Message}", Logger.LogLevel.DEBUG);
                return false;
            }
        }

        private void LoadPropertyDefinitions()
        {
            try
            {
                Logger.Log("[>] Chargement des Property Definitions...", Logger.LogLevel.DEBUG);
                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("--- LISTE COMPLETE DES PROPRIETES FILES ---", Logger.LogLevel.INFO);
                
                var filePropDefs = _connection!.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                
                if (filePropDefs != null)
                {
                    foreach (var propDef in filePropDefs)
                    {
                        // Stocker par SysName ET par DispName pour flexibilite
                        _propertyDefIds[propDef.SysName] = propDef.Id;
                        
                        // IMPORTANT: Aussi stocker par DispName pour proprietes avec GUID comme SysName
                        if (!string.IsNullOrEmpty(propDef.DispName) && propDef.DispName != propDef.SysName)
                        {
                            _propertyDefIds[propDef.DispName] = propDef.Id;
                        }
                        
                        // Log detaille pour diagnostic
                        string displayInfo = $"ID:{propDef.Id,4} | SysName: '{propDef.SysName}' | DispName: '{propDef.DispName}'";
                        
                        // Mettre en evidence les proprietes XNRGY
                        if (propDef.SysName == "Project" || propDef.DispName == "Project" ||
                            propDef.SysName == "Reference" || propDef.DispName == "Reference" ||
                            propDef.SysName == "Module" || propDef.DispName == "Module")
                        {
                            Logger.Log($"  >>> {displayInfo} <<<", Logger.LogLevel.INFO);
                        }
                        else
                        {
                            Logger.Log($"      {displayInfo}", Logger.LogLevel.TRACE);
                        }
                    }
                }
                
                Logger.Log("-------------------------------------------", Logger.LogLevel.INFO);
                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log($"[+] {_propertyDefIds.Count} Property Definitions chargees", Logger.LogLevel.DEBUG);
                
                // Verifier si les proprietes XNRGY existent (chercher par SysName OU DispName)
                bool hasProject = _propertyDefIds.ContainsKey("Project");
                bool hasReference = _propertyDefIds.ContainsKey("Reference");
                bool hasModule = _propertyDefIds.ContainsKey("Module");
                
                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("[>] VERIFICATION DES PROPRIETES XNRGY:", Logger.LogLevel.INFO);
                Logger.Log($"   Project   : {(hasProject ? "[+] TROUVE (ID: " + _propertyDefIds["Project"] + ")" : "[-] NON TROUVE")}", Logger.LogLevel.INFO);
                Logger.Log($"   Reference : {(hasReference ? "[+] TROUVE (ID: " + _propertyDefIds["Reference"] + ")" : "[-] NON TROUVE")}", Logger.LogLevel.INFO);
                Logger.Log($"   Module    : {(hasModule ? "[+] TROUVE (ID: " + _propertyDefIds["Module"] + ")" : "[-] NON TROUVE")}", Logger.LogLevel.INFO);
                Logger.Log("", Logger.LogLevel.INFO);
                
                if (!hasProject || !hasReference || !hasModule)
                {
                    Logger.Log("[!!!] ATTENTION: Proprietes XNRGY manquantes [!!!]", Logger.LogLevel.WARNING);
                    Logger.Log("", Logger.LogLevel.WARNING);
                    Logger.Log("[!] VERIFIEZ dans Vault Admin Console:", Logger.LogLevel.WARNING);
                    Logger.Log("   1. Ouvrir ADMS Console", Logger.LogLevel.WARNING);
                    Logger.Log("   2. Behaviors > Properties > File Properties", Logger.LogLevel.WARNING);
                    Logger.Log("   3. Verifier que les proprietes existent avec:", Logger.LogLevel.WARNING);
                    Logger.Log("      - Display Name: 'Project', 'Reference', 'Module'", Logger.LogLevel.WARNING);
                    Logger.Log("      - Apply to: [x] Files", Logger.LogLevel.WARNING);
                    Logger.Log("   4. Regarder le log ci-dessus pour voir les noms exacts", Logger.LogLevel.WARNING);
                    Logger.Log("", Logger.LogLevel.WARNING);
                    Logger.Log("[!] L'upload continuera SANS les proprietes manquantes", Logger.LogLevel.WARNING);
                    Logger.Log("", Logger.LogLevel.WARNING);
                }
                else
                {
                    Logger.Log("[+] Toutes les proprietes XNRGY sont detectees!", Logger.LogLevel.INFO);
                    Logger.Log("", Logger.LogLevel.INFO);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("LoadPropertyDefinitions", ex, Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Definit les permissions d'ecriture pour l'utilisateur connecte sur un fichier.
        /// NOTE: Cette fonctionnalite necessite des methodes avancees de l'API SecurityService
        /// qui ne sont pas disponibles dans toutes les versions du SDK.
        /// Pour definir les permissions, utilisez l'interface Vault ou configurez les permissions
        /// au niveau du dossier parent.
        /// </summary>
        /// <param name="fileMasterId">L'ID Master du fichier</param>
        private void SetFileWritePermissions(long fileMasterId)
        {
            // TODO: Implementer la definition des permissions via l'API Vault SDK
            // Les methodes necessaires (SetEntACLs, GetACEsByACLId, etc.) ne sont pas disponibles
            // dans la version actuelle du SDK ou necessitent des permissions administrateur.
            // 
            // Solution alternative: Configurer les permissions au niveau du dossier parent dans Vault
            // pour que les fichiers uploades heritent des bonnes permissions.
            Logger.Log($"   [!] Definition des permissions d'ecriture non implementee (limitation API)", Logger.LogLevel.WARNING);
            Logger.Log($"   [i] Solution: Configurez les permissions au niveau du dossier parent dans Vault", Logger.LogLevel.INFO);
        }

        /// <summary>
        /// Restaure la securite heritee sur un dossier pour permettre la creation d'elements enfants.
        /// Cette methode resout le probleme Error 1000 cause par les permissions "Overridden".
        /// </summary>
        /// <param name="folderId">L'ID du dossier sur lequel restaurer la securite heritee</param>
        private void RestoreInheritedSecurity(long folderId)
        {
            try
            {
                Logger.Log($"      [>] Restauration securite heritee pour dossier ID {folderId}...", Logger.LogLevel.DEBUG);
                
                // Obtenir les ACL actuelles du dossier
                var result = _connection!.WebServiceManager.SecurityService.GetEntACLsByEntityIds(new[] { folderId });
                
                if (result != null && result.ACLArray != null && result.ACLArray.Length > 0)
                {
                    var currentACL = result.ACLArray[0];
                    
                    // Appliquer le comportement "Combined" (heritage) au lieu de "Override"
                    // Combined = 1 signifie que la securite est heritee du parent
                    _connection.WebServiceManager.SecurityService.SetSystemACLs(
                        new[] { folderId },
                        currentACL.Id,
                        ACW.SysAclBeh.Combined  // CLE: Utilisera l'heritage au lieu de Override
                    );
                    
                    Logger.Log($"      [+] Securite heritee restauree (ACL ID: {currentACL.Id}, Beh: Combined)", Logger.LogLevel.DEBUG);
                }
                // Note: Pas de warning si ACL vide - le dossier herite automatiquement
            }
            catch (Exception ex)
            {
                // Log mais continue - le dossier a peut-etre deja la securite heritee
                Logger.Log($"      [!] RestoreInheritedSecurity a echoue (peut-etre deja herite): {ex.Message}", Logger.LogLevel.TRACE);
                Logger.LogException("RestoreInheritedSecurity", ex, Logger.LogLevel.TRACE);
            }
        }

        private VDF.Vault.Currency.Entities.Folder? EnsureVaultPathExists(string vaultFolderPath, 
            string? projectNumber = null, string? reference = null, string? module = null)
        {
            try
            {
                vaultFolderPath = vaultFolderPath.Trim();
                
                var pathParts = vaultFolderPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                
                if (pathParts.Length == 0 || pathParts[0] != "$")
                {
                    Logger.Log($"[!] Chemin invalide (doit commencer par $/): {vaultFolderPath}", Logger.LogLevel.WARNING);
                    return null;
                }

                VDF.Vault.Currency.Entities.Folder? currentFolder = _connection!.FolderManager.RootFolder;
                string currentPath = "$";

                for (int i = 1; i < pathParts.Length; i++)
                {
                    string folderName = pathParts[i];
                    string nextPath = currentPath + "/" + folderName;

                    VDF.Vault.Currency.Entities.Folder? nextFolder = null;
                    
                    try
                    {
                        var folders = _connection.WebServiceManager.DocumentService.FindFoldersByPaths(new[] { nextPath });
                        
                        if (folders != null && folders.Length > 0 && folders[0] != null && folders[0].Id > 0)
                        {
                            nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, folders[0]);
                            Logger.Log($"   [+] {nextPath}", Logger.LogLevel.TRACE);
                        }
                    }
                    catch
                    {
                    }

                    if (nextFolder == null)
                    {
                        Logger.Log($"   + Creation: {nextPath}", Logger.LogLevel.DEBUG);
                        
                        // Restaurer securite heritee sur parent AVANT creation
                        RestoreInheritedSecurity(currentFolder!.Id);
                        
                        try
                        {
                            var newFolder = _connection.WebServiceManager.DocumentService.AddFolder(
                                folderName,
                                currentFolder!.Id,
                                false
                            );
                            
                            if (newFolder != null && newFolder.Id > 0)
                            {
                                nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, newFolder);
                                Logger.Log($"   [+] Cree: {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                            }
                            else
                            {
                                Logger.Log($"   [-] Echec creation: {nextPath}", Logger.LogLevel.ERROR);
                                return null;
                            }
                        }
                        catch (ACW.VaultServiceErrorException ex) when (ex.ErrorCode == 1000)
                        {
                            // Retry une fois apres avoir fixe la securite
                            Logger.Log($"   [!] Error 1000 - Nouvelle tentative...", Logger.LogLevel.WARNING);
                            
                            try
                            {
                                var newFolder = _connection.WebServiceManager.DocumentService.AddFolder(
                                    folderName,
                                    currentFolder!.Id,
                                    false
                                );
                                
                                if (newFolder != null && newFolder.Id > 0)
                                {
                                    nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, newFolder);
                                    Logger.Log($"   [+] Cree (2eme tentative): {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                                }
                                else
                                {
                                    Logger.Log($"   [-] Echec creation (2eme tentative): {nextPath}", Logger.LogLevel.ERROR);
                                    return null;
                                }
                            }
                            catch (Exception retryEx)
                            {
                                Logger.LogException($"AddFolder retry for {nextPath}", retryEx, Logger.LogLevel.ERROR);
                                return null;
                            }
                        }
                        catch (ACW.VaultServiceErrorException ex) when (ex.ErrorCode == 1011)
                        {
                            // Erreur 1011: Le dossier existe deja - recuperer simplement
                            Logger.Log($"   [!] Dossier existe deja (erreur 1011), recuperation...", Logger.LogLevel.DEBUG);
                            try
                            {
                                var folders = _connection.WebServiceManager.DocumentService.FindFoldersByPaths(new[] { nextPath });
                                if (folders != null && folders.Length > 0 && folders[0] != null && folders[0].Id > 0)
                                {
                                    nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, folders[0]);
                                    Logger.Log($"   [+] Dossier recupere: {nextPath} (ID: {folders[0].Id})", Logger.LogLevel.DEBUG);
                                }
                                else
                                {
                                    // Fallback: chercher dans le dossier parent
                                    var parentFolders = _connection.WebServiceManager.DocumentService.GetFoldersByParentId(currentFolder!.Id, false);
                                    var existingFolder = parentFolders?.FirstOrDefault(f => f.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));
                                    if (existingFolder != null && existingFolder.Id > 0)
                                    {
                                        nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, existingFolder);
                                        Logger.Log($"   [+] Dossier recupere via parent: {nextPath} (ID: {existingFolder.Id})", Logger.LogLevel.DEBUG);
                                    }
                                    else
                                    {
                                        // Le dossier n'existe pas vraiment - peut-etre une fausse alerte ou race condition
                                        // Reessayer de creer le dossier
                                        Logger.Log($"   [!] Dossier non trouve malgre erreur 1011, nouvelle tentative de creation...", Logger.LogLevel.WARNING);
                                        try
                                        {
                                            var newFolder = _connection.WebServiceManager.DocumentService.AddFolder(
                                                folderName,
                                                currentFolder!.Id,
                                                false
                                            );
                                            
                                            if (newFolder != null && newFolder.Id > 0)
                                            {
                                                nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, newFolder);
                                                Logger.Log($"   [+] Dossier cree (2eme tentative): {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                                            }
                                            else
                                            {
                                                Logger.Log($"   [-] Echec creation (2eme tentative): {nextPath}", Logger.LogLevel.ERROR);
                                                return null;
                                            }
                                        }
                                        catch (ACW.VaultServiceErrorException retryEx) when (retryEx.ErrorCode == 1011)
                                        {
                                            // Encore l'erreur 1011 - le dossier existe vraiment maintenant
                                            Logger.Log($"   [!] Erreur 1011 persistante, recuperation finale...", Logger.LogLevel.DEBUG);
                                            var finalFolders = _connection.WebServiceManager.DocumentService.FindFoldersByPaths(new[] { nextPath });
                                            if (finalFolders != null && finalFolders.Length > 0 && finalFolders[0] != null && finalFolders[0].Id > 0)
                                            {
                                                nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, finalFolders[0]);
                                                Logger.Log($"   [+] Dossier recupere (final): {nextPath} (ID: {finalFolders[0].Id})", Logger.LogLevel.DEBUG);
                                            }
                                            else
                                            {
                                                Logger.Log($"   [-] Impossible de recuperer le dossier existant: {nextPath}", Logger.LogLevel.ERROR);
                                                return null;
                                            }
                                        }
                                        catch (Exception retryEx)
                                        {
                                            Logger.Log($"   [-] Erreur lors de la 2eme tentative: {retryEx.Message}", Logger.LogLevel.ERROR);
                                            return null;
                                        }
                                    }
                                }
                            }
                            catch (Exception recoverEx)
                            {
                                // Si la recuperation echoue, reessayer de creer
                                Logger.Log($"   [!] Erreur recuperation dossier: {recoverEx.Message}, nouvelle tentative de creation...", Logger.LogLevel.WARNING);
                                try
                                {
                                    var newFolder = _connection.WebServiceManager.DocumentService.AddFolder(
                                        folderName,
                                        currentFolder!.Id,
                                        false
                                    );
                                    
                                    if (newFolder != null && newFolder.Id > 0)
                                    {
                                        nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, newFolder);
                                        Logger.Log($"   [+] Dossier cree (apres erreur recuperation): {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                                    }
                                    else
                                    {
                                        Logger.Log($"   [-] Echec creation (apres erreur recuperation): {nextPath}", Logger.LogLevel.ERROR);
                                        return null;
                                    }
                                }
                                catch (Exception finalEx)
                                {
                                    Logger.Log($"   [-] Erreur finale: {finalEx.Message}", Logger.LogLevel.ERROR);
                                    return null;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogException($"AddFolder pour {nextPath}", ex, Logger.LogLevel.ERROR);
                            return null;
                        }
                    }

                    if (nextFolder != null)
                    {
                        currentFolder = nextFolder;
                    }
                    else
                    {
                        Logger.Log($"   [-] Impossible de continuer apres {nextPath}", Logger.LogLevel.ERROR);
                        return null;
                    }
                    currentPath = nextPath;
                }

                if (currentFolder == null)
                {
                    Logger.Log($"   [-] currentFolder est null a la fin", Logger.LogLevel.ERROR);
                    return null;
                }

                return currentFolder;
            }
            catch (Exception ex)
            {
                Logger.LogException($"EnsureVaultPathExists({vaultFolderPath})", ex, Logger.LogLevel.ERROR);
                return null;
            }
        }

       public bool CleanVaultFolder(string vaultFolderPath)
{
    try
    {
        Logger.Log($"[>] Nettoyage du dossier: {vaultFolderPath}", Logger.LogLevel.INFO);
        
        var folder = _connection!.WebServiceManager.DocumentService.GetFolderByPath(vaultFolderPath);
        
        if (folder == null)
        {
            Logger.Log($"   [!] Dossier non trouve - rien a nettoyer", Logger.LogLevel.WARNING);
            return true;
        }
        
        // Recuperer TOUS les fichiers (recursif)
        var files = _connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, true);
        
        if (files == null || files.Length == 0)
        {
            Logger.Log($"   [+] Dossier deja vide", Logger.LogLevel.INFO);
            return true;
        }
        
        Logger.Log($"   [>] {files.Length} fichier(s) a supprimer", Logger.LogLevel.INFO);
        
        int deleted = 0;
        int failed = 0;
        List<string> failedFiles = new List<string>();
        
        // Strategie: Supprimer d'abord les pieces (.ipt), puis les assemblages (.iam)
        // Cela evite les erreurs de dependances
        var parts = files.Where(f => f.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase)).ToArray();
        var assemblies = files.Where(f => f.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase)).ToArray();
        var otherFiles = files.Where(f => !f.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) && 
                                           !f.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase)).ToArray();
        
        // Ordre de suppression: autres fichiers -> pieces -> assemblages
        var orderedFiles = otherFiles.Concat(parts).Concat(assemblies);
        
        foreach (var file in orderedFiles)
        {
            try
            {
                // Forcer UndoCheckout si CheckedOut
                if (file.CheckedOut)
                {
                    try
                    {
                        Logger.Log($"   [~] Undo checkout: {file.Name}", Logger.LogLevel.TRACE);
                        _connection.WebServiceManager.DocumentService.UndoCheckoutFile(file.MasterId, out var ticket);
                    }
                    catch (Exception undoEx)
                    {
                        Logger.Log($"   [!] Undo checkout failed pour {file.Name}: {undoEx.Message}", Logger.LogLevel.TRACE);
                    }
                }
                
                // Tenter la suppression
                _connection.WebServiceManager.DocumentService.DeleteFileFromFolder(file.MasterId, folder.Id);
                
                deleted++;
                Logger.Log($"   [+] Supprime: {file.Name}", Logger.LogLevel.TRACE);
                
                if (deleted % 10 == 0)
                {
                    Logger.Log($"   [>] {deleted}/{files.Length} supprimes...", Logger.LogLevel.INFO);
                }
            }
            catch (ACW.VaultServiceErrorException vaultEx)
            {
                string errorMsg = $"{file.Name} (Error {vaultEx.ErrorCode})";
                
                if (vaultEx.ErrorCode == 303)
                {
                    // Error 303 = File has restrictions (lifecycle, dependencies, etc.)
                    Logger.Log($"   [!] Impossible de supprimer '{file.Name}': Error 303 - Le fichier a des restrictions (lifecycle, dependances, ou permissions)", Logger.LogLevel.WARNING);
                    Logger.Log($"      [i] Solution: Verifier dans Vault Client:", Logger.LogLevel.WARNING);
                    Logger.Log($"         - Lifecycle state permet-il la suppression?", Logger.LogLevel.WARNING);
                    Logger.Log($"         - Y a-t-il des references actives?", Logger.LogLevel.WARNING);
                    Logger.Log($"         - Permissions de suppression OK?", Logger.LogLevel.WARNING);
                }
                else if (vaultEx.ErrorCode == 1050)
                {
                    // Error 1050 = File has active references/dependencies
                    Logger.Log($"   [!] Impossible de supprimer '{file.Name}': Error 1050 - Fichier a des references actives", Logger.LogLevel.WARNING);
                    Logger.Log($"      [i] Ce fichier est reference par d'autres fichiers. Supprimer d'abord les fichiers qui l'utilisent.", Logger.LogLevel.WARNING);
                }
                else
                {
                    Logger.Log($"   [-] Impossible de supprimer '{file.Name}': Error {vaultEx.ErrorCode} - {vaultEx.Message}", Logger.LogLevel.WARNING);
                }
                
                failedFiles.Add(errorMsg);
                failed++;
            }
            catch (Exception delEx)
            {
                Logger.Log($"   [-] Impossible de supprimer '{file.Name}': {delEx.Message}", Logger.LogLevel.WARNING);
                failedFiles.Add($"{file.Name} ({delEx.Message})");
                failed++;
            }
        }
        
        Logger.Log($"", Logger.LogLevel.INFO);
        Logger.Log($"[+] Nettoyage termine: {deleted} supprimes, {failed} echoues", Logger.LogLevel.INFO);
        
        if (failed > 0)
        {
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"[!!!] FICHIERS NON SUPPRIMES [!!!]", Logger.LogLevel.WARNING);
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"Les fichiers suivants n'ont pas pu etre supprimes:", Logger.LogLevel.WARNING);
            foreach (var failedFile in failedFiles)
            {
                Logger.Log($"   - {failedFile}", Logger.LogLevel.WARNING);
            }
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"[i] RAISONS COURANTES:", Logger.LogLevel.WARNING);
            Logger.Log($"   1. Error 303 - Restrictions de lifecycle (etat Released, etc.)", Logger.LogLevel.WARNING);
            Logger.Log($"   2. Error 1050 - Fichier reference par d'autres fichiers", Logger.LogLevel.WARNING);
            Logger.Log($"   3. Permissions insuffisantes", Logger.LogLevel.WARNING);
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"[i] SOLUTIONS:", Logger.LogLevel.WARNING);
            Logger.Log($"   - Verifier le lifecycle state dans Vault Client", Logger.LogLevel.WARNING);
            Logger.Log($"   - Supprimer manuellement via Vault Client (avec plus de controle)", Logger.LogLevel.WARNING);
            Logger.Log($"   - Changer l'etat du fichier si necessaire", Logger.LogLevel.WARNING);
            Logger.Log($"   - Detacher les references avant suppression", Logger.LogLevel.WARNING);
            Logger.Log($"", Logger.LogLevel.WARNING);
        }
        Logger.Log($"", Logger.LogLevel.INFO);
        
        return failed == 0;
    }
    catch (Exception ex)
    {
        Logger.LogException($"CleanVaultFolder({vaultFolderPath})", ex, Logger.LogLevel.ERROR);
        return false;
    }
}

        /// <summary>
        /// Applique les proprietes sur un fichier
        /// -------------------------------------------------------------------------------
        /// SELON LA DOCUMENTATION SDK (topic36.html):
        /// 
        /// POUR LES UDP NON MAPPEES (proprietes Vault sans mapping iProperty):
        ///   - UpdateFileProperties fonctionne DIRECTEMENT sans checkout
        ///   - C'est le cas simple
        /// 
        /// POUR LES UDP MAPPEES (Content -> UDP, ex: iProperty -> Vault UDP):
        ///   - UpdateFileProperties NE FONCTIONNE PAS car ecrase au checkin
        ///   - Il faut modifier le fichier source avec l'API CAD
        /// 
        /// -------------------------------------------------------------------------------
        /// STRATEGIE ICI:
        /// 1. Essayer UpdateFileProperties directement (pour UDP non mappees)
        /// 2. Si ea ne marche pas ou si l'UDP est mappee, utiliser AcquireFiles
        /// -------------------------------------------------------------------------------
        /// </summary>
        /// <summary>
        /// Applique les proprietes Project, Reference, Module, Revision directement via UpdateFileProperties
        /// SOLUTION VALIDEE: Utiliser le MasterId (pas le FileId/IterationId)
        /// </summary>
        private bool SetFilePropertiesDirectly(long fileMasterId, string? projectNumber, string? reference, string? module, string? revision = null, string? checkInComment = null)
        {
            try
            {
                // -------------------------------------------------------------------------------
                // ETAPE 0: Toujours enlever Read-only AVANT toute operation
                // Cela evite que le Job Processor bloque les modifications
                // -------------------------------------------------------------------------------
                try
                {
                    var filesForReadOnly = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                    if (filesForReadOnly != null && filesForReadOnly.Length > 0)
                    {
                        var file = filesForReadOnly[0];
                        var folder = _connection.WebServiceManager.DocumentService.GetFolderById(file.FolderId);
                        var workingFolderPath = _connection.WorkingFoldersManager.GetWorkingFolder("$");
                        var workingFolder = workingFolderPath.FullPath;
                        var relativePath = folder.FullName.TrimStart('$', '/').Replace("/", "\\");
                        var localFolder = System.IO.Path.Combine(workingFolder, relativePath);
                        var localFilePath = System.IO.Path.Combine(localFolder, file.Name);
                        
                        if (System.IO.File.Exists(localFilePath))
                        {
                            var localFileInfo = new System.IO.FileInfo(localFilePath);
                            if (localFileInfo.IsReadOnly)
                            {
                                Logger.Log($"   [>] Retrait de l'attribut read-only AVANT traitement...", Logger.LogLevel.DEBUG);
                                localFileInfo.IsReadOnly = false;
                            }
                        }
                    }
                }
                catch (Exception attrEx)
                {
                    Logger.Log($"   [!] Impossible de modifier l'attribut read-only: {attrEx.Message}", Logger.LogLevel.TRACE);
                    // Continuer quand meme
                }
                Logger.Log($"[>] Application des proprietes (File MasterId: {fileMasterId})...", Logger.LogLevel.DEBUG);
                
                var properties = new List<ACW.PropInstParam>();
                
                // Verifier et ajouter Project
                if (!string.IsNullOrEmpty(projectNumber) && _propertyDefIds.ContainsKey("Project"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Project"],
                        Val = projectNumber
                    });
                    Logger.Log($"   Project = {projectNumber}", Logger.LogLevel.DEBUG);
                }
                
                // Verifier et ajouter Reference
                if (!string.IsNullOrEmpty(reference) && _propertyDefIds.ContainsKey("Reference"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Reference"],
                        Val = reference
                    });
                    Logger.Log($"   Reference = {reference}", Logger.LogLevel.DEBUG);
                }
                
                // Verifier et ajouter Module
                if (!string.IsNullOrEmpty(module) && _propertyDefIds.ContainsKey("Module"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Module"],
                        Val = module
                    });
                    Logger.Log($"   Module = {module}", Logger.LogLevel.DEBUG);
                }
                
                // Verifier et ajouter Revision
                if (!string.IsNullOrEmpty(revision) && _propertyDefIds.ContainsKey("Revision"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Revision"],
                        Val = revision
                    });
                    Logger.Log($"   Revision = {revision}", Logger.LogLevel.DEBUG);
                }
                
                if (properties.Count == 0)
                {
                    Logger.Log($"   [i] Aucune propriete a appliquer", Logger.LogLevel.DEBUG);
                    return true; // Pas d'erreur si aucune propriete a appliquer
                }
                
                    var propArray = new ACW.PropInstParamArray
                    {
                        Items = properties.ToArray()
                    };
                    
                // ================================================================================
                // SOLUTION FINALE:
                // - Inventor: GET + Checkout + Enlever Read-only + UpdateFileProperties + UndoCheckout
                //   -> Option A (UpdateFileProperties direct) supprimee - ne fonctionne pas
                //   -> Les proprietes UDP persistent apres UndoCheckout (stockees dans DB Vault)
                //   -> UndoCheckout enleve la barre dans Vault Client
                // - Non-Inventor: Checkout + UpdateFileProperties + CheckIn
                // ================================================================================
                
                // 1. Recuperer les infos du fichier
                        var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                        if (latestFiles == null || latestFiles.Length == 0)
                        {
                    Logger.Log($"   [ERREUR] Fichier non trouve (MasterId: {fileMasterId})", Logger.LogLevel.WARNING);
                    return false;
                        }
                        var fileInfo = latestFiles[0];
                
                // Determiner si c'est un fichier Inventor
                bool isInventorFile = fileInfo.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) ||
                                      fileInfo.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) ||
                                      fileInfo.Name.EndsWith(".idw", StringComparison.OrdinalIgnoreCase) ||
                                      fileInfo.Name.EndsWith(".ipn", StringComparison.OrdinalIgnoreCase);
                
                Logger.Log($"   [FILE] Fichier: {fileInfo.Name} (FileId: {fileInfo.Id}, MasterId: {fileInfo.MasterId}) [Inventor={isInventorFile}]", Logger.LogLevel.DEBUG);
                
                // ================================================================================
                // STRATEGIE POUR FICHIERS INVENTOR:
                // Essayer UpdateFileProperties DIRECTEMENT (sans checkout/checkin)
                // Si erreur 1013 (fichier doit etre checke out), alors checkout -> update -> checkin
                // ================================================================================
                if (isInventorFile)
                {
                    // -------------------------------------------------------------------------------
                    // STRATEGIE VALIDEE: GET -> CHECKOUT -> UpdateFileProperties -> CHECKIN
                    // (La tentative directe echoue toujours avec erreur 1136 - Job Processor occupe)
                    // -------------------------------------------------------------------------------
                    ACW.File? checkedOutFileInv = null;
                    string? localFilePath = null;
                    try
                    {
                        Logger.Log($"   [INVENTOR] Fallback: GET + Checkout + Update UDP + CheckIn...", Logger.LogLevel.DEBUG);
                        
                        // Etape 1: Determiner le dossier local
                        var workingFolderPath = _connection.WorkingFoldersManager.GetWorkingFolder("$");
                        var workingFolder = workingFolderPath.FullPath;
                        var folderPath = _connection.WebServiceManager.DocumentService.GetFolderById(fileInfo.FolderId);
                        var relativePath = folderPath.FullName.TrimStart('$', '/').Replace("/", "\\");
                        var localFolder = System.IO.Path.Combine(workingFolder, relativePath);
                        
                        if (!System.IO.Directory.Exists(localFolder))
                            System.IO.Directory.CreateDirectory(localFolder);
                        
                        localFilePath = System.IO.Path.Combine(localFolder, fileInfo.Name);
                        
                        // -------------------------------------------------------------------------------
                        // eTAPE 1: VRAI GET - Telecharger le fichier AVANT le checkout
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] GET - Telechargement du fichier AVANT checkout...", Logger.LogLevel.DEBUG);
                        try
                        {
                            // Obtenir le FileIteration pour le telechargement
                            var wsFile = _connection.WebServiceManager.DocumentService.GetFileById(fileInfo.Id);
                            var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, wsFile);
                            
                            // Creer les parametres de telechargement
                            var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                            downloadSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                            
                            // Telecharger le fichier
                            var downloadResult = _connection.FileManager.AcquireFiles(downloadSettings);
                            
                            if (downloadResult.FileResults != null && downloadResult.FileResults.Any(r => r.LocalPath != null))
                            {
                                localFilePath = downloadResult.FileResults.First().LocalPath.FullPath;
                                Logger.Log($"   [OK] Fichier telecharge: {localFilePath}", Logger.LogLevel.INFO);
                                
                                // Verifier que le fichier existe vraiment
                                if (!System.IO.File.Exists(localFilePath))
                                {
                                    Logger.Log($"   [!] AcquireFiles a retourne un chemin mais le fichier n'existe pas: {localFilePath}", Logger.LogLevel.WARNING);
                                    throw new System.IO.FileNotFoundException($"Fichier non trouve apres telechargement: {localFilePath}");
                                }
                            }
                            else
                            {
                                Logger.Log($"   [!] AcquireFiles n'a pas retourne de chemin local", Logger.LogLevel.WARNING);
                                throw new Exception("AcquireFiles n'a pas telecharge le fichier");
                            }
                        }
                        catch (Exception acquireEx)
                        {
                            Logger.Log($"   [-] Erreur lors du telechargement avec AcquireFiles: {acquireEx.Message}", Logger.LogLevel.ERROR);
                            throw; // Arreter si le GET echoue
                        }
                        
                        // -------------------------------------------------------------------------------
                        // eTAPE 2: CHECKOUT (apres le GET, le fichier est deje local)
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] Checkout (fichier deje telecharge e: {localFilePath})...", Logger.LogLevel.DEBUG);
                        
                        // Verifier que le fichier existe avant le checkout
                        if (!System.IO.File.Exists(localFilePath))
                        {
                            throw new System.IO.FileNotFoundException($"Fichier non trouve avant checkout: {localFilePath}");
                        }
                        
                        try
                        {
                            checkedOutFileInv = _connection.WebServiceManager.DocumentService.CheckoutFile(
                                fileInfo.Id,
                            ACW.CheckoutFileOptions.Master,
                            Environment.MachineName,
                                localFolder,
                                "Modification UDP via VaultAutomationTool",
                                out var downloadTicket
                            );
                            
                            if (checkedOutFileInv == null)
                            {
                                throw new Exception("CheckoutFile a retourne null");
                            }
                            
                            Logger.Log($"   [OK] Checkout reussi (FileId: {checkedOutFileInv.Id}, MasterId: {checkedOutFileInv.MasterId})", Logger.LogLevel.INFO);
                        }
                        catch (ACW.VaultServiceErrorException checkoutEx)
                        {
                            Logger.Log($"   [-] Erreur Vault lors du checkout: Code {checkoutEx.ErrorCode} - {checkoutEx.Message}", Logger.LogLevel.ERROR);
                            if (checkoutEx.InnerException != null)
                            {
                                Logger.Log($"   [-] Inner Exception: {checkoutEx.InnerException.Message}", Logger.LogLevel.ERROR);
                            }
                            throw;
                        }
                        catch (Exception checkoutEx)
                        {
                            Logger.Log($"   [-] Erreur lors du checkout: {checkoutEx.Message}", Logger.LogLevel.ERROR);
                            Logger.Log($"   [i] StackTrace: {checkoutEx.StackTrace}", Logger.LogLevel.DEBUG);
                            throw;
                        }
                        
                        // -------------------------------------------------------------------------------
                        // ETAPE 3: UPDATE FILE PROPERTIES (UDP) - sur le fichier checke out
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] UpdateFileProperties (UDP)...", Logger.LogLevel.DEBUG);
                        _connection.WebServiceManager.DocumentService.UpdateFileProperties(
                            new[] { checkedOutFileInv.MasterId },
                            new[] { propArray }
                        );
                        Logger.Log($"   [OK] {properties.Count} UDP appliquees sur fichier checke out", Logger.LogLevel.INFO);
                        
                        // NOTE: La synchronisation UDP -> iProperties sera faite manuellement par l'utilisateur
                        // via clic droit -> "Synchronize Properties" dans Vault Explorer
                        // (ExplorerUtil et Inventor iProperties desactives pour performance)
                        
                        // -------------------------------------------------------------------------------
                        // ETAPE 4: CHECKIN pour persister les proprietes (incluant le writeback iProperties)
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] CheckIn pour persister les proprietes...", Logger.LogLevel.DEBUG);
                        
                        // Obtenir le fichier pour le CheckIn
                        var wsFileForCheckin = _connection.WebServiceManager.DocumentService.GetFileById(checkedOutFileInv.Id);
                        var vaultFileForCheckin = new VDF.Vault.Currency.Entities.FileIteration(_connection, wsFileForCheckin);
                        
                        // Commentaire detaille pour l'application des proprietes (3eme version)
                        _connection.FileManager.CheckinFile(
                            vaultFileForCheckin,
                            "MAJ proprietes (Project/Reference/Module) via Vault Automation Tool",
                            false,  // keepCheckedOut
                            wsFileForCheckin.ModDate,
                            null,
                            null,
                            false,
                            null,
                            ACW.FileClassification.None,
                            false,
                            null
                        );
                        
                        Logger.Log($"   [+] CheckIn reussi - Proprietes PERSISTEES!", Logger.LogLevel.INFO);
                        
                        // -------------------------------------------------------------------------------
                        // GET FINAL - Telecharger le fichier pour enlever le rond rouge
                        // -------------------------------------------------------------------------------
                        try
                        {
                            Logger.Log($"   [INVENTOR] GET FINAL - Telechargement pour enlever le rond rouge...", Logger.LogLevel.DEBUG);
                            
                            // Recuperer le fichier apres CheckIn
                            var latestFilesAfterCheckin = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { checkedOutFileInv.MasterId });
                            if (latestFilesAfterCheckin != null && latestFilesAfterCheckin.Length > 0)
                            {
                                var fileAfterCheckin = latestFilesAfterCheckin[0];
                                var fileIterationAfterCheckin = new VDF.Vault.Currency.Entities.FileIteration(_connection, fileAfterCheckin);
                                
                                // Creer les parametres de telechargement
                                var finalDownloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                                finalDownloadSettings.AddFileToAcquire(fileIterationAfterCheckin, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                                
                                // Telecharger le fichier
                                var finalDownloadResult = _connection.FileManager.AcquireFiles(finalDownloadSettings);
                                
                                if (finalDownloadResult.FileResults != null && finalDownloadResult.FileResults.Any(r => r.LocalPath != null))
                                {
                                    Logger.Log($"   [+] GET FINAL reussi - Fichier synchronise, rond rouge enleve!", Logger.LogLevel.INFO);
                                }
                                else
                                {
                                    Logger.Log($"   [!] GET FINAL n'a pas retourne de chemin local", Logger.LogLevel.DEBUG);
                                }
                            }
                        }
                        catch (Exception getEx)
                        {
                            Logger.Log($"   [!] Erreur GET final: {getEx.Message}", Logger.LogLevel.DEBUG);
                            // Ne pas faire echouer si le GET echoue
                        }
                        
                        return true;
                    }
                    catch (ACW.VaultServiceErrorException vex)
                    {
                        Logger.Log($"   [WARN] Erreur Vault {vex.ErrorCode}: {vex.Message}", Logger.LogLevel.WARNING);
                        return false;
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"   [WARN] Erreur: {ex.Message}", Logger.LogLevel.WARNING);
                        return false;
                    }
                }
                
                // ================================================================================
                // STRATEGIE NON-INVENTOR: Checkout + Update + CheckIn (sans GET prealable)
                // ================================================================================
                ACW.File? checkedOutFileNonInv = null;
                string? localFilePathNonInv = null;
                
                try
                {
                    // 2. Checkout via CheckoutFile
                    Logger.Log($"   [LOCK] Checkout du fichier...", Logger.LogLevel.DEBUG);
                    
                    // Determiner le chemin local via Working Folder
                    var workingFolderPath = _connection.WorkingFoldersManager.GetWorkingFolder("$");
                    var workingFolder = workingFolderPath.FullPath;
                    var folderPath = _connection.WebServiceManager.DocumentService.GetFolderById(fileInfo.FolderId);
                    var relativePath = folderPath.FullName.TrimStart('$', '/').Replace("/", "\\");
                    var localFolder = System.IO.Path.Combine(workingFolder, relativePath);
                    localFilePathNonInv = System.IO.Path.Combine(localFolder, fileInfo.Name);
                    
                    // S'assurer que le dossier existe
                    if (!System.IO.Directory.Exists(localFolder))
                    {
                        System.IO.Directory.CreateDirectory(localFolder);
                    }
                    
                    // Checkout avec download via ticket
                    checkedOutFileNonInv = _connection.WebServiceManager.DocumentService.CheckoutFile(
                        fileInfo.Id,
                                ACW.CheckoutFileOptions.Master,
                                Environment.MachineName,
                        localFolder,
                        "Modification proprietes via VaultAutomationTool",
                        out var downloadTicket
                    );
                    
                    // Le fichier devrait maintenant etre telecharge dans localFilePathNonInv
                    // Si pas, utiliser AcquireFiles
                    if (!System.IO.File.Exists(localFilePathNonInv))
                    {
                        var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, checkedOutFileNonInv);
                        var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                        downloadSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                        var downloadResult = _connection.FileManager.AcquireFiles(downloadSettings);
                        
                        if (downloadResult.FileResults.Any(r => r.LocalPath != null))
                        {
                            localFilePathNonInv = downloadResult.FileResults.First().LocalPath.FullPath;
                        }
                    }
                    
                    Logger.Log($"   [OK] Checkout reussi vers: {localFilePathNonInv}", Logger.LogLevel.DEBUG);
                    
                    // Retirer l'attribut read-only sur le fichier local apres checkout
                    try
                    {
                        if (System.IO.File.Exists(localFilePathNonInv))
                        {
                            var localFileInfo = new System.IO.FileInfo(localFilePathNonInv);
                            if (localFileInfo.IsReadOnly)
                            {
                                Logger.Log($"   [>] Retrait de l'attribut read-only sur le fichier local...", Logger.LogLevel.DEBUG);
                                localFileInfo.IsReadOnly = false;
                            }
                        }
                    }
                    catch (Exception attrEx)
                    {
                        Logger.Log($"   [!] Impossible de modifier l'attribut read-only du fichier local: {attrEx.Message}", Logger.LogLevel.TRACE);
                        // Continuer quand meme
                    }
                    
                    // 3. UpdateFileProperties avec le MasterId
                    Logger.Log($"   [SAVE] UpdateFileProperties avec MasterId ({checkedOutFileNonInv.MasterId})...", Logger.LogLevel.DEBUG);
                            _connection.WebServiceManager.DocumentService.UpdateFileProperties(
                        new[] { checkedOutFileNonInv.MasterId },
                                new[] { propArray }
                            );
                    
                    Logger.Log($"   [OK] {properties.Count} propriete(s) modifiee(s)", Logger.LogLevel.INFO);
                    
                    // 4. CheckIn pour PERSISTER les proprietes (meme strategie pour tous les fichiers)
                    Logger.Log($"   [CHECKIN] CheckIn pour persister les proprietes...", Logger.LogLevel.DEBUG);
                    
                    // Creer FileIteration pour CheckIn
                    var wsFile = _connection.WebServiceManager.DocumentService.GetFileById(checkedOutFileNonInv.Id);
                    var vaultFile = new VDF.Vault.Currency.Entities.FileIteration(_connection, wsFile);
                    
                    // Commentaire detaille pour l'application des proprietes (3eme version)
                    _connection.FileManager.CheckinFile(
                        vaultFile,
                        "MAJ proprietes (Project / Reference / Module) via Vault Automation Tool",
                        false,  // keepCheckedOut
                        wsFile.ModDate,
                            null,
                            null,
                            false,
                            null,
                            ACW.FileClassification.None,
                            false,
                            null
                        );
                        
                    Logger.Log($"   [OK] CheckIn reussi - Proprietes PERSISTEES!", Logger.LogLevel.INFO);
                    
                    // -------------------------------------------------------------------------------
                    // GET FINAL - Telecharger le fichier pour enlever le rond rouge (comme pour Inventor)
                    // -------------------------------------------------------------------------------
                    try
                    {
                        Logger.Log($"   [NON-INVENTOR] GET FINAL - Telechargement pour enlever le rond rouge...", Logger.LogLevel.DEBUG);
                        
                        // Recuperer le fichier apres CheckIn
                        var latestFilesAfterCheckin = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { checkedOutFileNonInv.MasterId });
                        if (latestFilesAfterCheckin != null && latestFilesAfterCheckin.Length > 0)
                        {
                            var fileAfterCheckin = latestFilesAfterCheckin[0];
                            var fileIterationAfterCheckin = new VDF.Vault.Currency.Entities.FileIteration(_connection, fileAfterCheckin);
                            
                            // Creer les parametres de telechargement
                            var finalDownloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                            finalDownloadSettings.AddFileToAcquire(fileIterationAfterCheckin, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                            
                            // Telecharger le fichier
                            var finalDownloadResult = _connection.FileManager.AcquireFiles(finalDownloadSettings);
                            
                            if (finalDownloadResult.FileResults != null && finalDownloadResult.FileResults.Any(r => r.LocalPath != null))
                            {
                                Logger.Log($"   [+] GET FINAL reussi - Fichier synchronise, rond rouge enleve!", Logger.LogLevel.INFO);
                            }
                            else
                            {
                                Logger.Log($"   [!] GET FINAL n'a pas retourne de chemin local", Logger.LogLevel.DEBUG);
                            }
                        }
                    }
                    catch (Exception getEx)
                    {
                        Logger.Log($"   [!] Erreur GET final: {getEx.Message}", Logger.LogLevel.DEBUG);
                        // Ne pas faire echouer si le GET echoue
                    }
                    
                    return true;
                    }
                    catch (ACW.VaultServiceErrorException vex)
                    {
                    Logger.Log($"   [WARN] Erreur Vault {vex.ErrorCode}: {vex.Message}", Logger.LogLevel.WARNING);
                    
                    // Annuler le checkout en cas d erreur
                    if (checkedOutFileNonInv != null)
                            {
                                try
                                {
                            _connection!.WebServiceManager.DocumentService.UndoCheckoutFile(checkedOutFileNonInv.MasterId, out _);
                            Logger.Log($"   [UNLOCK] Checkout annule apres erreur", Logger.LogLevel.DEBUG);
                        }
                        catch { /* Ignorer les erreurs de cleanup */ }
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException($"SetFilePropertiesDirectly({fileMasterId})", ex, Logger.LogLevel.WARNING);
                return false;
            }
        }

        /// <summary>
        /// Verifie que les proprietes sont bien enregistrees dans Vault apres leur application
        /// </summary>
        private void VerifyFileProperties(long fileMasterId, string? projectNumber, string? reference, string? module)
                        {
                            try
                            {
                Logger.Log($"   [>] Verification des proprietes enregistrees...", Logger.LogLevel.DEBUG);
                
                // Recuperer le fichier pour obtenir ses proprietes
                var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                if (latestFiles == null || latestFiles.Length == 0)
                {
                    Logger.Log($"   [!] Fichier non trouve pour verification", Logger.LogLevel.DEBUG);
                                return;
                            }
                            
                var file = latestFiles[0];
                
                // Recuperer les definitions de proprietes
                var propDefs = _connection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                
                // Trouver les IDs des proprietes Project, Reference, Module
                long? projectPropId = null;
                long? referencePropId = null;
                long? modulePropId = null;
                
                foreach (var propDef in propDefs)
                {
                    if (propDef.SysName == "Project")
                        projectPropId = propDef.Id;
                    else if (propDef.SysName == "Reference")
                        referencePropId = propDef.Id;
                    else if (propDef.SysName == "Module")
                        modulePropId = propDef.Id;
                }
                
                if (!projectPropId.HasValue && !referencePropId.HasValue && !modulePropId.HasValue)
                {
                    Logger.Log($"   [!] Proprietes Project/Reference/Module non trouvees pour verification", Logger.LogLevel.DEBUG);
                    return;
                }
                
                // Recuperer les valeurs des proprietes
                var propIdsToGet = new List<long>();
                if (projectPropId.HasValue) propIdsToGet.Add(projectPropId.Value);
                if (referencePropId.HasValue) propIdsToGet.Add(referencePropId.Value);
                if (modulePropId.HasValue) propIdsToGet.Add(modulePropId.Value);
                
                // Utiliser GetProperties avec la bonne signature (entityClassId, entityIds, propertyDefIds)
                var propInsts = _connection.WebServiceManager.PropertyService.GetProperties(
                    "FILE",
                    new[] { file.Id },
                    propIdsToGet.ToArray()
                );
                
                if (propInsts != null && propInsts.Length > 0)
                {
                    foreach (var propInst in propInsts)
                    {
                        var propDef = propDefs.FirstOrDefault(p => p.Id == propInst.PropDefId);
                        string propName = propDef?.DispName ?? "Unknown";
                        string propValue = propInst.Val?.ToString() ?? "(vide)";
                        
                        if (propInst.PropDefId == projectPropId)
                        {
                            bool match = propValue == (projectNumber ?? "");
                            Logger.Log($"   {(match ? "[+]" : "[-]")} {propName}: '{propValue}' (attendu: '{projectNumber ?? ""}')", 
                                match ? Logger.LogLevel.INFO : Logger.LogLevel.WARNING);
                        }
                        else if (propInst.PropDefId == referencePropId)
                        {
                            bool match = propValue == (reference ?? "");
                            Logger.Log($"   {(match ? "[+]" : "[-]")} {propName}: '{propValue}' (attendu: '{reference ?? ""}')", 
                                match ? Logger.LogLevel.INFO : Logger.LogLevel.WARNING);
                        }
                        else if (propInst.PropDefId == modulePropId)
                        {
                            bool match = propValue == (module ?? "");
                            Logger.Log($"   {(match ? "[+]" : "[-]")} {propName}: '{propValue}' (attendu: '{module ?? ""}')", 
                                match ? Logger.LogLevel.INFO : Logger.LogLevel.WARNING);
                        }
                    }
                }
                else
                {
                    Logger.Log($"   [!] Aucune propriete trouvee pour verification", Logger.LogLevel.DEBUG);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   [!] Erreur lors de la verification des proprietes: {ex.Message}", Logger.LogLevel.DEBUG);
                // Ne pas faire echouer l'operation si la verification echoue
            }
        }

        /// <summary>
        /// Applique les proprietes Project, Reference, Module a un fichier
        /// Delegue a SetFilePropertiesDirectly
        /// </summary>
        private void SetFileProperties(long fileMasterId, string? projectNumber, string? reference, string? module)
                            {
                                try
                                {
                // Simplement deleguer a SetFilePropertiesDirectly qui est optimise
                bool success = SetFilePropertiesDirectly(fileMasterId, projectNumber, reference, module);
                if (!success)
                {
                    Logger.Log($"   [!] SetFileProperties: echec pour MasterId {fileMasterId}", Logger.LogLevel.WARNING);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException($"SetFileProperties(MasterId={fileMasterId})", ex, Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Assigne une categorie e un fichier via DocumentServiceExtensions.UpdateFileCategories
        /// Cette methode est ESSENTIELLE pour les categories Engineering, Office, Standard
        /// car FileManager.AddFile ne permet PAS d'assigner une categorie directement
        /// </summary>
        /// <param name="fileMasterId">ID Master du fichier</param>
        /// <param name="categoryId">ID de la categorie e assigner</param>
        /// <param name="categoryName">Nom de la categorie (pour logging)</param>
        /// <returns>true si la categorie a ete assignee avec succes</returns>
        private bool AssignCategoryToFile(long fileMasterId, long categoryId, string categoryName)
        {
            if (!IsConnected || fileMasterId <= 0 || categoryId <= 0)
            {
                return false;
            }

            try
            {
                Logger.Log($"   [~] Assignation de la categorie '{categoryName}' (ID: {categoryId}) au fichier (MasterId: {fileMasterId})...", Logger.LogLevel.INFO);

                // Recuperer le fichier pour obtenir son ID d'iteration
                var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                if (latestFiles == null || latestFiles.Length == 0)
                {
                    Logger.Log($"   [!] Fichier non trouve pour l'assignation de categorie", Logger.LogLevel.WARNING);
                    return false;
                }

                var file = latestFiles[0];
                bool categoryAssigned = false;

                // -------------------------------------------------------------------------------
                // METHODE 1: Via DocumentServiceExtensions.UpdateFileCategories (RECOMMANDE)
                // Cette methode est la plus fiable selon la documentation Vault SDK
                // Signature: UpdateFileCategories(long[] fileMasterIds, long[] catIds)
                // -------------------------------------------------------------------------------
                try
                {
                    var wsManagerType = _connection.WebServiceManager.GetType();
                    var documentServiceExtensionsProperty = wsManagerType.GetProperty("DocumentServiceExtensions");

                    if (documentServiceExtensionsProperty != null)
                    {
                        var documentServiceExtensions = documentServiceExtensionsProperty.GetValue(_connection.WebServiceManager);
                        if (documentServiceExtensions != null)
                        {
                            // SIGNATURE CORRECTE: UpdateFileCategories(long[] masterIds, long[] categoryIds, string comment)
                            var updateCatMethod = documentServiceExtensions.GetType().GetMethod("UpdateFileCategories",
                                new[] { typeof(long[]), typeof(long[]), typeof(string) });

                            if (updateCatMethod != null)
                            {
                                updateCatMethod.Invoke(documentServiceExtensions,
                                    new object[] {
                                        new[] { file.MasterId },
                                        new[] { categoryId },
                                        $"Categorie assignee automatiquement: {categoryName}"
                                    });
                                Logger.Log($"   [+] Categorie '{categoryName}' assignee via DocumentServiceExtensions.UpdateFileCategories", Logger.LogLevel.INFO);
                                categoryAssigned = true;
                            }
                            else
                            {
                                Logger.Log($"   [i] Methode UpdateFileCategories non trouvee dans DocumentServiceExtensions", Logger.LogLevel.DEBUG);
                            }
                        }
                    }
                }
                catch (Exception reflectEx)
                {
                    Logger.Log($"   [!] Erreur reflection UpdateFileCategories: {reflectEx.Message}", Logger.LogLevel.DEBUG);
                    
                    // Log l'exception interne si presente (souvent plus informative)
                    if (reflectEx.InnerException != null)
                    {
                        Logger.Log($"   [!] Inner Exception: {reflectEx.InnerException.Message}", Logger.LogLevel.DEBUG);
                    }
                }

                // Si la methode principale n'a pas fonctionne, afficher un message
                if (!categoryAssigned)
                {
                    Logger.Log($"   [!] Impossible d'assigner la categorie '{categoryName}' - API UpdateFileCategories non disponible ou erreur", Logger.LogLevel.WARNING);
                    Logger.Log($"   [i] Le fichier aura la categorie par defaut. Assignez manuellement dans Vault si necessaire.", Logger.LogLevel.INFO);
                }

                return categoryAssigned;
            }
            catch (Exception ex)
            {
                Logger.LogException($"AssignCategoryToFile(fileMasterId={fileMasterId}, categoryId={categoryId})", ex, Logger.LogLevel.WARNING);
                return false;
            }
        }

        /// <summary>
        /// Charge les Categories disponibles et trouve la Category "Base"
        /// </summary>
        private void LoadCategories()
        {
            try
            {
                Logger.Log("[>] Chargement des Categories...", Logger.LogLevel.DEBUG);
                
                var categories = _connection!.WebServiceManager.CategoryService.GetCategoriesByEntityClassId("FILE", false);
                
                if (categories != null)
                {
                    foreach (var category in categories)
                    {
                        Logger.Log($"   Category: ID={category.Id}, Name='{category.Name}'", Logger.LogLevel.TRACE);
                        
                        // Chercher la Category "Base" (par Name)
                        if (category.Name.Equals("Base", StringComparison.OrdinalIgnoreCase))
                        {
                            _baseCategoryId = category.Id;
                            Logger.Log($"   [+] Category 'Base' trouvee (ID: {category.Id})", Logger.LogLevel.INFO);
                            break;
                        }
                    }
                    
                    if (_baseCategoryId == null)
                    {
                        Logger.Log($"   [!] Category 'Base' non trouvee", Logger.LogLevel.WARNING);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("LoadCategories", ex, Logger.LogLevel.WARNING);
            }
        }

        // NOTE: Les revisions sont gerees automatiquement par Vault via les transitions d'etat
        // (Work in Progress -> Released incremente la revision selon le schema ASME configure)

        /// <summary>
        /// Obtient toutes les Categories disponibles pour les fichiers
        /// </summary>
        public List<(long Id, string Name)> GetAvailableCategories()
        {
            var categories = new List<(long Id, string Name)>();
            
            try
            {
                if (!IsConnected)
                {
                    return categories;
                }
                
                var vaultCategories = _connection!.WebServiceManager.CategoryService.GetCategoriesByEntityClassId("FILE", false);
                
                if (vaultCategories != null)
                {
                    foreach (var category in vaultCategories)
                    {
                        categories.Add((category.Id, category.Name));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("GetAvailableCategories", ex, Logger.LogLevel.WARNING);
            }
            
            return categories;
        }

        /// <summary>
        /// Obtient toutes les Lifecycle Definitions disponibles
        /// </summary>
        public List<Models.LifecycleDefinitionItem> GetAvailableLifecycleDefinitions()
        {
            var lifecycleDefs = new List<Models.LifecycleDefinitionItem>();
            
            try
            {
                if (!IsConnected)
                {
                    return lifecycleDefs;
                }
                
                var allLifecycleDefs = _connection!.WebServiceManager.LifeCycleService.GetAllLifeCycleDefinitions();
                
                if (allLifecycleDefs != null)
                {
                    foreach (var lifecycleDef in allLifecycleDefs)
                    {
                        var lifecycleItem = new Models.LifecycleDefinitionItem
                        {
                            Id = lifecycleDef.Id,
                            Name = lifecycleDef.DispName
                        };
                        
                        // Ajouter les etats
                        if (lifecycleDef.StateArray != null)
                        {
                            foreach (var state in lifecycleDef.StateArray)
                            {
                                lifecycleItem.States.Add(new Models.LifecycleStateItem
                                {
                                    Id = state.Id,
                                    Name = state.DispName,
                                    LifecycleDefinitionId = lifecycleDef.Id
                                });
                            }
                        }
                        
                        lifecycleDefs.Add(lifecycleItem);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("GetAvailableLifecycleDefinitions", ex, Logger.LogLevel.WARNING);
            }
            
            return lifecycleDefs;
        }

        /// <summary>
        /// Obtient le Lifecycle Definition ID selon la categorie selectionnee
        /// Engineering -> Flexible Release Process (For Review, Work in Progress, Released, Obsolete)
        /// Office -> Simple Release Process (Work in Progress, Released, Obsolete)
        /// Design Representation -> Design Representation Process (Released, Work in Progress, Obsolete)
        /// Standard -> Basic Release Process
        /// Base -> Aucun (pas de Lifecycle)
        /// </summary>
        public long? GetLifecycleDefinitionIdByCategory(string? categoryName)
        {
            if (string.IsNullOrEmpty(categoryName))
            {
                return null;
            }

            string categoryLower = categoryName!.ToLowerInvariant().Trim();
            
            // Base n'a pas de Lifecycle - retourner null directement
            if (categoryLower == "base")
            {
                Logger.Log($"   [i] Categorie 'Base' n'a pas de Lifecycle Definition", Logger.LogLevel.INFO);
                return null;
            }
            
            // Mapping categorie -> Lifecycle Definition Name
            // IMPORTANT: Ces mappings correspondent aux Lifecycle Definitions configurees dans Vault
            var categoryToLifecycleMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "engineering", "Flexible Release Process" },
                { "office", "Simple Release Process" },
                { "standard", "Basic Release Process" },
                { "design representation", "Design Representation Process" }
            };

            if (categoryToLifecycleMapping.TryGetValue(categoryLower, out string? lifecycleName))
            {
                try
                {
                    var allLifecycleDefs = _connection!.WebServiceManager.LifeCycleService.GetAllLifeCycleDefinitions();
                    var lifecycleDef = allLifecycleDefs?.FirstOrDefault(l => l.DispName.Equals(lifecycleName, StringComparison.OrdinalIgnoreCase));
                    
                    if (lifecycleDef != null)
                    {
                        Logger.Log($"   [>] Mapping categorie '{categoryName}' -> Lifecycle '{lifecycleName}' (ID: {lifecycleDef.Id})", Logger.LogLevel.INFO);
                        return lifecycleDef.Id;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   [!] Erreur lors de la recherche du Lifecycle Definition: {ex.Message}", Logger.LogLevel.WARNING);
                }
            }
            
            return null;
        }

        /// <summary>
        /// Obtient l'ID de l'etat "Work in Progress" pour un Lifecycle Definition donne
        /// </summary>
        public long? GetWorkInProgressStateId(long lifecycleDefinitionId)
        {
            try
            {
                var allLifecycleDefs = _connection!.WebServiceManager.LifeCycleService.GetAllLifeCycleDefinitions();
                var lifecycleDef = allLifecycleDefs?.FirstOrDefault(l => l.Id == lifecycleDefinitionId);
                
                if (lifecycleDef?.StateArray != null)
                {
                    // Chercher l'etat "Work in Progress" ou l'etat par defaut
                    var wipState = lifecycleDef.StateArray.FirstOrDefault(s => 
                        s.DispName.IndexOf("Work", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        s.DispName.IndexOf("Progress", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        s.DispName.IndexOf("WIP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        s.IsDflt);
                    
                    if (wipState != null)
                    {
                        return wipState.Id;
                    }
                    
                    // Si aucun etat WIP trouve, prendre le premier etat
                    if (lifecycleDef.StateArray.Length > 0)
                    {
                        return lifecycleDef.StateArray[0].Id;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   [!] Erreur lors de la recherche de l'etat WIP: {ex.Message}", Logger.LogLevel.WARNING);
            }
            
            return null;
        }

        private ACW.FileClassification DetermineFileClassification(string extension)
        {
            switch (extension.ToLowerInvariant())
            {
                case ".ipt":
                case ".iam":
                case ".idw":
                case ".ipn":
                case ".dwg":
                case ".dxf":
                case ".dwf":
                    return ACW.FileClassification.DesignRepresentation;
                default:
                    return ACW.FileClassification.None;
            }
        }

        /// <summary>
        /// Determine le FileClassification en fonction de la categorie selectionnee
        /// Selon la documentation SDK, les valeurs disponibles sont :
        /// None, DesignRepresentation, DesignDocument, DesignVisualization, DesignPresentation,
        /// DesignSubstitute, ConfigurationFactory, ConfigurationMember, ElectricalProject
        /// </summary>
        private ACW.FileClassification DetermineFileClassificationByCategory(long? categoryId, string categoryName)
        {
            if (!categoryId.HasValue || categoryId.Value <= 0)
            {
                // Aucune categorie selectionnee -> None
                return ACW.FileClassification.None;
            }

            // Mapper le nom de la categorie au FileClassification correspondant
            if (string.IsNullOrEmpty(categoryName))
            {
                return ACW.FileClassification.None;
            }

            string categoryLower = categoryName.ToLowerInvariant().Trim();
            
            // MAPPING EXPLICITE selon documentation SDK
            // Dictionnaire de mapping categorie -> FileClassification
            // -------------------------------------------------------------------------------
            // IMPORTANT: Le FileClassification determine comment Vault traite le fichier
            // - None: Fichiers generiques (PDF, Excel, Word, images, etc.)
            // - DesignRepresentation: Fichiers CAD qui representent un design (DWF, PDF generes)
            // - DesignDocument: Documents de design (pas couramment utilise)
            // 
            // Pour Engineering, Office, Standard: FileClassification.None est correct car
            // la categorie sera assignee SEPAREMENT via UpdateFileCategories apres l'upload
            // -------------------------------------------------------------------------------
            var categoryMapping = new Dictionary<string, ACW.FileClassification>(StringComparer.OrdinalIgnoreCase)
            {
                // -------------------------------------------------------------------------------
                // Categories utilisant FileClassification.None (la plupart des cas)
                // La vraie categorie est assignee APRES l'upload via UpdateFileCategories
                // -------------------------------------------------------------------------------
                { "base", ACW.FileClassification.None },
                { "aucune", ACW.FileClassification.None },
                { "engineering", ACW.FileClassification.None },      // Engineering -> None
                { "office", ACW.FileClassification.None },           // Office -> None
                { "standard", ACW.FileClassification.None },         // Standard -> None
                { "cad", ACW.FileClassification.None },
                { "design", ACW.FileClassification.None },
                { "document", ACW.FileClassification.None },
                { "other", ACW.FileClassification.None },
                { "autre", ACW.FileClassification.None },
                
                // -------------------------------------------------------------------------------
                // Categories avec FileClassification specifique
                // Ces categories necessitent un FileClassification particulier pour Vault
                // -------------------------------------------------------------------------------
                { "design representation", ACW.FileClassification.DesignRepresentation },
                { "designrepresentation", ACW.FileClassification.DesignRepresentation },
                { "design document", ACW.FileClassification.DesignDocument },
                { "designdocument", ACW.FileClassification.DesignDocument },
                { "design visualization", ACW.FileClassification.DesignVisualization },
                { "designvisualization", ACW.FileClassification.DesignVisualization },
                { "design presentation", ACW.FileClassification.DesignPresentation },
                { "designpresentation", ACW.FileClassification.DesignPresentation },
                { "design substitute", ACW.FileClassification.DesignSubstitute },
                { "designsubstitute", ACW.FileClassification.DesignSubstitute },
                { "configuration factory", ACW.FileClassification.ConfigurationFactory },
                { "configurationfactory", ACW.FileClassification.ConfigurationFactory },
                { "configuration member", ACW.FileClassification.ConfigurationMember },
                { "configurationmember", ACW.FileClassification.ConfigurationMember },
                { "electrical project", ACW.FileClassification.ElectricalProject },
                { "electricalproject", ACW.FileClassification.ElectricalProject }
            };

            // 1. Chercher dans le mapping explicite (nom exact)
            if (categoryMapping.TryGetValue(categoryName, out ACW.FileClassification mappedValue))
            {
                Logger.Log($"   [>] Mapping explicite: Categorie '{categoryName}' -> FileClassification.{mappedValue}", Logger.LogLevel.INFO);
                return mappedValue;
            }

            // 2. Chercher dans le mapping explicite (nom en minuscules)
            if (categoryMapping.TryGetValue(categoryLower, out ACW.FileClassification mappedValueLower))
            {
                Logger.Log($"   [>] Mapping explicite (lowercase): Categorie '{categoryName}' -> FileClassification.{mappedValueLower}", Logger.LogLevel.INFO);
                return mappedValueLower;
            }

            // 3. Essayer de parser directement (nom nettoye)
            try
            {
                string cleanCategoryName = categoryName.Replace(" ", "").Replace("-", "").Replace("_", "");
                if (System.Enum.TryParse<ACW.FileClassification>(cleanCategoryName, true, out ACW.FileClassification result))
                {
                    Logger.Log($"   [>] Mapping direct (nettoye): Categorie '{categoryName}' -> FileClassification.{result}", Logger.LogLevel.INFO);
                    return result;
                }
            }
            catch { }

            // 4. Chercher correspondance partielle dans les valeurs enum disponibles
            var fileClassificationType = typeof(ACW.FileClassification);
            var enumValues = System.Enum.GetValues(fileClassificationType);
            
            foreach (ACW.FileClassification enumValue in enumValues)
            {
                string enumName = enumValue.ToString();
                string enumNameLower = enumName.ToLowerInvariant();
                
                // Correspondance exacte (insensible a la casse)
                if (categoryLower == enumNameLower)
                {
                    Logger.Log($"   [>] Mapping exact: Categorie '{categoryName}' -> FileClassification.{enumName}", Logger.LogLevel.INFO);
                    return enumValue;
                }
                
                // Correspondance partielle (categorie contient le nom de l'enum)
                if (categoryLower.Contains(enumNameLower) || enumNameLower.Contains(categoryLower))
                {
                    Logger.Log($"   [>] Mapping partiel: Categorie '{categoryName}' -> FileClassification.{enumName}", Logger.LogLevel.INFO);
                    return enumValue;
                }
            }
            
            // 5. Mapping logique par defaut
            // Engineering, CAD, Design (generique) -> None (fichiers CAD standard selon doc)
            if (categoryLower.Contains("engineering") || categoryLower == "cad" || 
                (categoryLower.Contains("design") && !categoryLower.Contains("representation") && 
                 !categoryLower.Contains("document") && !categoryLower.Contains("visualization") &&
                 !categoryLower.Contains("presentation") && !categoryLower.Contains("substitute")))
            {
                Logger.Log($"   [>] Mapping logique: Categorie '{categoryName}' -> FileClassification.None (fichier CAD standard)", Logger.LogLevel.INFO);
                return ACW.FileClassification.None;
            }
            
            // Si aucune correspondance trouvee, logger toutes les valeurs disponibles
            Logger.Log($"   [!] Aucun FileClassification trouve pour categorie '{categoryName}'", Logger.LogLevel.WARNING);
            Logger.Log($"   [i] Valeurs FileClassification disponibles: {string.Join(", ", enumValues.Cast<ACW.FileClassification>().Select(v => v.ToString()))}", Logger.LogLevel.DEBUG);
            Logger.Log($"   [!] Utilisation FileClassification.None par defaut", Logger.LogLevel.WARNING);
            
            // Par defaut, utiliser None (fichiers CAD standard)
            return ACW.FileClassification.None;
        }

        public bool UploadFile(string filePath, string vaultFolderPath, 
            string? projectNumber = null, string? reference = null, string? module = null, 
            long? categoryId = null, string? categoryName = null,
            long? lifecycleDefinitionId = null, long? lifecycleStateId = null, string? revision = null, string? checkInComment = null)
        {
            if (!IsConnected)
            {
                Logger.Log("[-] Non connecte a Vault", Logger.LogLevel.ERROR);
                return false;
            }

            if (!File.Exists(filePath))
            {
                Logger.Log($"[-] Fichier introuvable: {filePath}", Logger.LogLevel.ERROR);
                return false;
            }

            // Retirer l'attribut read-only si present pour permettre la modification des proprietes
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                if (fileInfo.IsReadOnly)
                {
                    Logger.Log($"   [>] Retrait de l'attribut read-only sur le fichier...", Logger.LogLevel.DEBUG);
                    fileInfo.IsReadOnly = false;
                }
            }
            catch (Exception attrEx)
            {
                Logger.Log($"   [!] Impossible de modifier l'attribut read-only: {attrEx.Message}", Logger.LogLevel.WARNING);
                // Continuer quand meme - certains systemes peuvent permettre l'acces meme avec read-only
            }

            VDF.Vault.Currency.Entities.Folder? targetFolder = null;
            try
            {
                string fileName = Path.GetFileName(filePath);
                string fileExtension = Path.GetExtension(fileName);
                
                Logger.Log($"[>] Upload: '{fileName}' -> {vaultFolderPath}", Logger.LogLevel.INFO);
                if (!string.IsNullOrEmpty(projectNumber)) Logger.Log($"   Project: {projectNumber}", Logger.LogLevel.INFO);
                if (!string.IsNullOrEmpty(reference)) Logger.Log($"   Reference: {reference}", Logger.LogLevel.INFO);
                if (!string.IsNullOrEmpty(module)) Logger.Log($"   Module: {module}", Logger.LogLevel.INFO);

                targetFolder = EnsureVaultPathExists(vaultFolderPath, projectNumber, reference, module);
                
                if (targetFolder == null || targetFolder.Id <= 0)
                {
                    Logger.Log($"[-] Impossible de creer/acceder au dossier: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"   Dossier cible valide (ID: {targetFolder.Id})", Logger.LogLevel.TRACE);

                bool fileExists = false;
                long existingFileId = -1;
                
                try
                {
                    var filesInFolder = _connection!.WebServiceManager.DocumentService.GetLatestFilesByFolderId(
                        targetFolder.Id,
                        false
                    );

                    if (filesInFolder != null && filesInFolder.Length > 0)
                    {
                        var existingFile = filesInFolder.FirstOrDefault(f => f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase));
                        if (existingFile != null)
                        {
                            fileExists = true;
                            existingFileId = existingFile.MasterId;
                        }
                    }
                }
                catch (Exception searchEx)
                {
                    Logger.Log($"   Recherche fichier: {searchEx.Message}", Logger.LogLevel.TRACE);
                }

                if (fileExists && existingFileId > 0)
                {
                    Logger.Log($"   [i] Le fichier '{fileName}' existe deja (ID: {existingFileId})", Logger.LogLevel.INFO);
                    
                    // NOUVELLE STRATEGIE: Ajouter a la file d'attente au lieu d'appliquer immediatement
                    if (!string.IsNullOrEmpty(projectNumber) || !string.IsNullOrEmpty(reference) || !string.IsNullOrEmpty(module))
                    {
                        QueuePropertyUpdate(existingFileId, projectNumber, reference, module, fileName);
                        Logger.Log($"   [+] Proprietes ajoutees a la file d'attente", Logger.LogLevel.INFO);
                    }
                    
                    // Assigner la categorie au fichier existant si specifiee
                    if (categoryId.HasValue && categoryId.Value > 0)
                    {
                        AssignCategoryToFile(existingFileId, categoryId.Value, categoryName ?? "");
                    }
                    
                    // Assigner le Lifecycle Definition si specifie
                    if (lifecycleDefinitionId.HasValue && lifecycleStateId.HasValue)
                    {
                        try
                        {
                            Logger.Log($"   [~] Assignation du Lifecycle Definition au fichier existant (ID: {lifecycleDefinitionId}, State ID: {lifecycleStateId})...", Logger.LogLevel.INFO);
                            
                            var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { existingFileId });
                            if (latestFiles != null && latestFiles.Length > 0)
                            {
                                var file = latestFiles[0];
                                
                                // Assigner le lifecycle state via DocumentService
                                try
                                {
                                    var documentServiceType = _connection.WebServiceManager.DocumentService.GetType();
                                    var updateMethod = documentServiceType.GetMethod("UpdateFileLifeCycleStates", 
                                        new[] { typeof(long[]), typeof(long[]), typeof(string) });
                                    
                                    if (updateMethod != null)
                                    {
                                        updateMethod.Invoke(_connection.WebServiceManager.DocumentService, 
                                            new object[] { new[] { file.Id }, new[] { lifecycleStateId.Value }, "Assignation lifecycle via upload" });
                                        Logger.Log($"   [+] Lifecycle assigne avec succes", Logger.LogLevel.INFO);
                                    }
                                }
                                catch (Exception reflectEx)
                                {
                                    Logger.Log($"   [!] Erreur reflection UpdateFileLifeCycleStates: {reflectEx.Message}", Logger.LogLevel.WARNING);
                                }
                            }
                        }
                        catch (Exception lifecycleEx)
                        {
                            Logger.Log($"   [!] Erreur lors de l'assignation du lifecycle: {lifecycleEx.Message}", Logger.LogLevel.WARNING);
                        }
                    }
                    
                    return true; // Succes apres mise a jour des proprietes
                }

                // -------------------------------------------------------------------------------
                // OPTION A (RAPIDE) - Upload SANS modification iProperties prealable
                // 
                // Strategie optimisee pour upload massif (1800+ fichiers):
                // 1. Upload fichier vers Vault
                // 2. UpdateFileProperties -> definir UDP Vault
                // 3. Job Processor sync UDP -> iProperties (apres check-in)
                // 
                // Resultat: UDP Vault correctes + Job Processor sync vers iProperties
                // Temps: ~100ms par fichier
                // -------------------------------------------------------------------------------

                // -------------------------------------------------------------------------------
                // DESACTIVE: NativeOlePropertyService n'ecrit PAS les vraies iProperties Inventor
                // Il ecrit dans le PropertySet OLE Windows standard, pas dans le format Autodesk
                // La sync sera faite par le Job Processor via SubmitSyncPropertiesJob()
                // -------------------------------------------------------------------------------
                /*
                string ext = Path.GetExtension(fileName).ToLowerInvariant();
                bool isInventorFile = ext == ".ipt" || ext == ".iam" || ext == ".idw" || ext == ".ipn";
                if (isInventorFile && (!string.IsNullOrEmpty(projectNumber) || !string.IsNullOrEmpty(reference) || !string.IsNullOrEmpty(module)))
                {
                    try
                    {
                        Logger.Log($"   [PRE-UPLOAD] Modification iProperties via NativeOLE (rapide)...", Logger.LogLevel.INFO);
                        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                        
                        using (var oleService = new NativeOlePropertyService())
                        {
                            bool iPropsSet = oleService.SetIProperties(filePath, projectNumber ?? "", reference ?? "", module ?? "");
                            stopwatch.Stop();
                            
                            if (iPropsSet)
                            {
                                Logger.Log($"   [PRE-UPLOAD] [+] iProperties definies via NativeOLE en {stopwatch.ElapsedMilliseconds}ms!", Logger.LogLevel.INFO);
                            }
                            else
                            {
                                Logger.Log($"   [PRE-UPLOAD] [!] NativeOLE: echec modification iProperties", Logger.LogLevel.WARNING);
                            }
                        }
                    }
                    catch (Exception oleEx)
                    {
                        Logger.Log($"   [PRE-UPLOAD] [!] Exception NativeOLE: {oleEx.Message}", Logger.LogLevel.WARNING);
                        // Continuer l'upload meme si la modification iProperties echoue
                    }
                }
                */

                // SOLUTION: Utiliser FileClassification selon la categorie selectionnee
                // Base -> None, Design Representation -> DesignRepresentation, etc.
                ACW.FileClassification fileClass = DetermineFileClassificationByCategory(categoryId, categoryName ?? string.Empty);
                
                long newFileId = -1;
                long entityIterationId = -1;
                
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var fileInfo = new FileInfo(filePath);

                    // Upload via FileManager.AddFile avec FileClassification selon categorie
                    // Le 3eme parametre est le commentaire pour la version 1
                    // Documentation: "Text data to be associated with version 1 of the file"
                    Logger.Log($"   [~] Upload avec FileClassification.{fileClass} (Categorie: {categoryName ?? "Aucune"})...", Logger.LogLevel.INFO);
                    var addedFile = _connection!.FileManager.AddFile(
                        targetFolder,
                        fileName,
                        checkInComment ?? string.Empty,  // Commentaire pour la version 1
                        fileInfo.LastWriteTime,
                        null,
                        null,
                        fileClass,
                        false,
                        fileStream
                    );
                    
                    if (addedFile != null && addedFile.EntityIterationId > 0)
                    {
                        try
                        {
                            entityIterationId = addedFile.EntityIterationId;
                            var file = _connection.WebServiceManager.DocumentService.GetFileById(entityIterationId);
                            if (file != null)
                            {
                                newFileId = file.MasterId;
                                Logger.Log($"   File MasterId: {file.MasterId}, Version: {file.VerNum}", Logger.LogLevel.TRACE);
                                
                                // Le commentaire a ete passe directement a AddFile (3eme parametre)
                                // Documentation: "Text data to be associated with version 1 of the file"
                                if (!string.IsNullOrWhiteSpace(checkInComment))
                                {
                                    Logger.Log($"   [+] Commentaire applique a la version 1 via AddFile: '{checkInComment}'", Logger.LogLevel.INFO);
                                }
                            }
                        }
                        catch (Exception idEx)
                        {
                            Logger.Log($"   [i] Info fichier: {idEx.Message}", Logger.LogLevel.TRACE);
                        }
                    }
                }

                Logger.Log($"[+] Uploade: '{fileName}'", Logger.LogLevel.INFO);

                // -------------------------------------------------------------------------------
                // ETAPE 1: Assigner la CATEGORIE via AssignCategoryToFile
                // Cette etape est CRUCIALE pour Engineering, Office, Standard
                // La categorie DOIT etre assignee APRES l'upload car FileManager.AddFile ne le fait pas
                // -------------------------------------------------------------------------------
                if (newFileId > 0 && categoryId.HasValue && categoryId.Value > 0)
                {
                    AssignCategoryToFile(newFileId, categoryId.Value, categoryName ?? "");
                }

                // -------------------------------------------------------------------------------
                // ETAPE 2: Assigner le Lifecycle Definition via UpdateFileLifeCycleDefinitions
                // -------------------------------------------------------------------------------
                if (newFileId > 0 && lifecycleDefinitionId.HasValue && lifecycleStateId.HasValue)
                {
                    bool lifecycleAssigned = false;
                    
                    try
                    {
                        Logger.Log($"   [~] Assignation du Lifecycle Definition (ID: {lifecycleDefinitionId}, State ID: {lifecycleStateId})...", Logger.LogLevel.INFO);
                        
                        // Recuperer le fichier pour obtenir son ID d'iteration
                        var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { newFileId });
                        if (latestFiles != null && latestFiles.Length > 0)
                        {
                            var file = latestFiles[0];
                            
                            // Chercher dans DocumentServiceExtensions (via WebServiceManager)
                            var wsManagerType = _connection.WebServiceManager.GetType();
                            var documentServiceExtensionsProperty = wsManagerType.GetProperty("DocumentServiceExtensions");
                            
                            if (documentServiceExtensionsProperty != null)
                            {
                                var documentServiceExtensions = documentServiceExtensionsProperty.GetValue(_connection.WebServiceManager);
                                if (documentServiceExtensions != null)
                                {
                                    // ---------------------------------------------------------------
                                    // METHODE 1: UpdateFileLifeCycleDefinitions (assigne Definition + State)
                                    // ---------------------------------------------------------------
                                    try
                                    {
                                        var updateDefMethod = documentServiceExtensions.GetType().GetMethod("UpdateFileLifeCycleDefinitions", 
                                            new[] { typeof(long[]), typeof(long[]), typeof(long[]), typeof(string) });
                                        
                                        if (updateDefMethod != null)
                                        {
                                            updateDefMethod.Invoke(documentServiceExtensions, 
                                                new object[] { 
                                                    new[] { file.MasterId },  // Utiliser MasterId au lieu de Id
                                                    new[] { lifecycleDefinitionId.Value }, 
                                                    new[] { lifecycleStateId.Value }, 
                                                    "Assignation lifecycle via VaultAutomationTool" 
                                                });
                                            Logger.Log($"   [+] Lifecycle Definition + State assignes via UpdateFileLifeCycleDefinitions", Logger.LogLevel.INFO);
                                            lifecycleAssigned = true;
                                        }
                                    }
                                    catch (Exception defEx)
                                    {
                                        var innerMsg = defEx.InnerException?.Message ?? defEx.Message;
                                        Logger.Log($"   [i] UpdateFileLifeCycleDefinitions echoue: {innerMsg}", Logger.LogLevel.DEBUG);
                                        
                                        // ---------------------------------------------------------------
                                        // METHODE 2: UpdateFileLifeCycleStates (change seulement le State)
                                        // Utile si le fichier a deja une Lifecycle Definition assignee
                                        // ---------------------------------------------------------------
                                        try
                                        {
                                            var updateStateMethod = documentServiceExtensions.GetType().GetMethod("UpdateFileLifeCycleStates", 
                                                new[] { typeof(long[]), typeof(long[]), typeof(string) });
                                            
                                            if (updateStateMethod != null)
                                            {
                                                updateStateMethod.Invoke(documentServiceExtensions, 
                                                    new object[] { 
                                                        new[] { file.MasterId },
                                                        new[] { lifecycleStateId.Value }, 
                                                        "Changement state via VaultAutomationTool" 
                                                    });
                                                Logger.Log($"   [+] State assigne via UpdateFileLifeCycleStates", Logger.LogLevel.INFO);
                                                lifecycleAssigned = true;
                                            }
                                        }
                                        catch (Exception stateEx)
                                        {
                                            var stateInnerMsg = stateEx.InnerException?.Message ?? stateEx.Message;
                                            Logger.Log($"   [i] UpdateFileLifeCycleStates echoue: {stateInnerMsg}", Logger.LogLevel.DEBUG);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception lifecycleEx)
                    {
                        Logger.Log($"   [!] Erreur lors de l'assignation du lifecycle: {lifecycleEx.Message}", Logger.LogLevel.WARNING);
                    }
                    
                    if (!lifecycleAssigned)
                    {
                        Logger.Log($"   [!] Lifecycle non assigne - Le fichier sera avec le state par defaut", Logger.LogLevel.WARNING);
                        Logger.Log($"   [i] Vous pouvez changer le state manuellement dans Vault Client", Logger.LogLevel.INFO);
                    }
                }

                // -------------------------------------------------------------------------------
                // ETAPE 2.5: Appliquer la revision via DocumentServiceExtensions.UpdateFileRevisionNumbers
                // -------------------------------------------------------------------------------
                if (newFileId > 0 && !string.IsNullOrEmpty(revision))
                {
                    try
                    {
                        Logger.Log($"   [~] Assignation de la revision '{revision}' au fichier (MasterId: {newFileId})...", Logger.LogLevel.INFO);
                        
                        // Obtenir le fichier pour avoir son FileId (pas MasterId)
                        var latestFiles = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { newFileId });
                        if (latestFiles != null && latestFiles.Length > 0)
                        {
                            var file = latestFiles[0];
                            
                            // Acceder a DocumentServiceExtensions via WebServiceManager
                            var wsManagerType = _connection.WebServiceManager.GetType();
                            var documentServiceExtensionsProperty = wsManagerType.GetProperty("DocumentServiceExtensions");
                            
                            if (documentServiceExtensionsProperty != null)
                            {
                                var documentServiceExtensions = documentServiceExtensionsProperty.GetValue(_connection.WebServiceManager);
                                if (documentServiceExtensions != null)
                                {
                                    // Appeler UpdateFileRevisionNumbers
                                    // Signature: UpdateFileRevisionNumbers(long[] fileIds, string[] revNumbers, string comment)
                                    var updateRevisionMethod = documentServiceExtensions.GetType().GetMethod("UpdateFileRevisionNumbers", 
                                        new[] { typeof(long[]), typeof(string[]), typeof(string) });
                                    
                                    if (updateRevisionMethod != null)
                                    {
                                        updateRevisionMethod.Invoke(documentServiceExtensions, 
                                            new object[] { 
                                                new[] { file.Id },  // FileId (pas MasterId)
                                                new[] { revision },  // Revision
                                                "Assignation revision via VaultAutomationTool" 
                                            });
                                        Logger.Log($"   [+] Revision '{revision}' assignee via UpdateFileRevisionNumbers", Logger.LogLevel.INFO);
                                        
                                        // Recuperer la nouvelle version apres mise a jour de la revision
                                        var updatedFiles = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { newFileId });
                                        if (updatedFiles != null && updatedFiles.Length > 0)
                                        {
                                            newFileId = updatedFiles[0].MasterId;
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log($"   [!] UpdateFileRevisionNumbers non trouvee dans DocumentServiceExtensions", Logger.LogLevel.WARNING);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception revEx)
                    {
                        var innerMsg = revEx.InnerException?.Message ?? revEx.Message;
                        Logger.Log($"   [!] Erreur assignation revision: {innerMsg}", Logger.LogLevel.WARNING);
                        // Ne pas faire echouer l'upload si la revision echoue
                    }
                }
                
                // ETAPE 3: Appliquer les proprietes (Project, Reference, Module, Revision)
                // -------------------------------------------------------------------------------
                if (newFileId > 0 && (!string.IsNullOrEmpty(projectNumber) || !string.IsNullOrEmpty(reference) || !string.IsNullOrEmpty(module) || !string.IsNullOrEmpty(revision)))
                {
                    SetFilePropertiesDirectly(newFileId, projectNumber, reference, module, revision, checkInComment);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogException($"UploadFile({Path.GetFileName(filePath)})", ex, Logger.LogLevel.ERROR);
                
                if (ex is ACW.VaultServiceErrorException vaultEx)
                {
                    Logger.Log($"   Code erreur Vault: {vaultEx.ErrorCode}", Logger.LogLevel.ERROR);
                    
                    // Erreur 1000 ou 1008: Fichier existe deja
                    if (vaultEx.ErrorCode == 1000 || vaultEx.ErrorCode == 1008)
                    {
                        Logger.Log($"   [i] Le fichier existe deja (erreur {vaultEx.ErrorCode})", Logger.LogLevel.INFO);
                        
                        // Essayer de recuperer le fichier existant pour appliquer les proprietes
                        if (targetFolder != null && targetFolder.Id > 0)
                        {
                            try
                            {
                                var filesInFolder = _connection!.WebServiceManager.DocumentService.GetLatestFilesByFolderId(
                                    targetFolder.Id,
                                    false
                                );
                            
                                if (filesInFolder != null && filesInFolder.Length > 0)
                                {
                                    var existingFile = filesInFolder.FirstOrDefault(f => f.Name.Equals(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase));
                                    if (existingFile != null && existingFile.MasterId > 0)
                                    {
                                        Logger.Log($"   [+] Fichier existant trouve (ID: {existingFile.MasterId}, Version: {existingFile.VerNum})", Logger.LogLevel.INFO);
                                        
                                        // Appliquer les proprietes au fichier existant (CheckOut -> Update Properties -> CheckIn)
                                        Logger.Log($"   [~] Application des proprietes au fichier existant...", Logger.LogLevel.INFO);
                                        SetFilePropertiesDirectly(existingFile.MasterId, projectNumber, reference, module, revision, checkInComment);
                                        
                                        // Assigner le Lifecycle Definition si specifie
                                        if (lifecycleDefinitionId.HasValue && lifecycleStateId.HasValue)
                                        {
                                            try
                                            {
                                                Logger.Log($"   [~] Assignation du Lifecycle Definition au fichier existant (ID: {lifecycleDefinitionId}, State ID: {lifecycleStateId})...", Logger.LogLevel.INFO);
                                                
                                                var documentServiceType = _connection.WebServiceManager.DocumentService.GetType();
                                                var updateMethod = documentServiceType.GetMethod("UpdateFileLifeCycleStates", 
                                                    new[] { typeof(long[]), typeof(long[]), typeof(string) });
                                                
                                                if (updateMethod != null)
                                                {
                                                    updateMethod.Invoke(_connection.WebServiceManager.DocumentService, 
                                                        new object[] { new[] { existingFile.Id }, new[] { lifecycleStateId.Value }, "Assignation lifecycle via upload" });
                                                    Logger.Log($"   [+] Lifecycle assigne avec succes", Logger.LogLevel.INFO);
                                                }
                                            }
                                            catch (Exception lifecycleEx)
                                            {
                                                Logger.Log($"   [!] Erreur lors de l'assignation du lifecycle: {lifecycleEx.Message}", Logger.LogLevel.WARNING);
                                            }
                                        }
                                        
                                        return true; // Succes apres mise a jour des proprietes
                                    }
                                }
                            }
                            catch (Exception recoveryEx)
                            {
                                Logger.Log($"   [!] Impossible de recuperer le fichier existant: {recoveryEx.Message}", Logger.LogLevel.WARNING);
                            }
                        }
                        else
                        {
                            Logger.Log($"   [!] Impossible de recuperer le fichier existant: dossier non disponible", Logger.LogLevel.WARNING);
                        }
                    }
                }
                
                return false;
            }
        }

        public (int success, int failed, int skipped) UploadFolder(string localFolderPath, string vaultBasePath,
            string? projectNumber = null, string? reference = null, string? module = null)
        {
            if (!IsConnected)
            {
                Logger.Log("[-] Non connecte a Vault", Logger.LogLevel.ERROR);
                return (0, 0, 0);
            }

            if (!Directory.Exists(localFolderPath))
            {
                Logger.Log($"[-] Dossier introuvable: {localFolderPath}", Logger.LogLevel.ERROR);
                return (0, 0, 0);
            }

            Logger.Log($"[>] Upload dossier complet: {localFolderPath}", Logger.LogLevel.INFO);
            Logger.Log($"   Destination Vault: {vaultBasePath}", Logger.LogLevel.INFO);
            Logger.Log($"   Project: {projectNumber}, Reference: {reference}, Module: {module}", Logger.LogLevel.INFO);

            int successCount = 0;
            int failedCount = 0;
            int skippedCount = 0;

            try
            {
                var files = Directory.GetFiles(localFolderPath, "*.*", SearchOption.AllDirectories);

                Logger.Log($"   {files.Length} fichier(s) trouve(s)", Logger.LogLevel.INFO);

                foreach (var filePath in files)
                {
                    try
                    {
                        string? directoryName = Path.GetDirectoryName(filePath);
                        if (string.IsNullOrEmpty(directoryName)) 
                        {
                            Logger.Log($"   [!] Impossible d'obtenir le repertoire pour: {filePath}", Logger.LogLevel.WARNING);
                            continue;
                        }

                        string relativePath = directoryName
                            .Substring(localFolderPath.Length)
                            .TrimStart('\\', '/')
                            .Replace('\\', '/');
                        
                        string vaultPath = string.IsNullOrEmpty(relativePath) 
                            ? vaultBasePath 
                            : vaultBasePath + "/" + relativePath;

                        if (UploadFile(filePath, vaultPath, projectNumber, reference, module))
                        {
                            successCount++;
                        }
                        else
                        {
                            skippedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogException($"Erreur fichier {Path.GetFileName(filePath)}", ex, Logger.LogLevel.ERROR);
                        failedCount++;
                    }
                }

                Logger.Log($"[OK] Upload termine: {successCount} succes, {skippedCount} ignores, {failedCount} erreurs", Logger.LogLevel.INFO);
                return (successCount, failedCount, skippedCount);
            }
            catch (Exception ex)
            {
                Logger.LogException($"UploadFolder({localFolderPath})", ex, Logger.LogLevel.ERROR);
                return (successCount, failedCount, skippedCount);
            }
        }

        public void Disconnect()
        {
            try
            {
                if (_connection != null)
                {
                    Logger.Log("Deconnexion de Vault...", Logger.LogLevel.INFO);
                    VDF.Vault.Library.ConnectionManager.LogOut(_connection);
                    Logger.Log("[OK] Deconnecte du Vault", Logger.LogLevel.INFO);
                    _connection = null;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("Disconnect", ex, Logger.LogLevel.WARNING);
            }
        }

        // -------------------------------------------------------------------------------
        // SOLUTION: Upload SANS lifecycle "Released" comme Inventor
        // -------------------------------------------------------------------------------

        /// <summary>
        /// Upload fichier SANS lifecycle "Released" (comme Inventor)
        /// </summary>
        public bool UploadFileWithoutLifecycle(string filePath, string vaultFolderPath, 
            string? projectNumber = null, string? reference = null, string? module = null)
        {
            if (!IsConnected)
            {
                Logger.Log("[-] Non connecte a Vault", Logger.LogLevel.ERROR);
                return false;
            }

            if (!File.Exists(filePath))
            {
                Logger.Log($"[-] Fichier introuvable: {filePath}", Logger.LogLevel.ERROR);
                return false;
            }

            try
            {
                string fileName = Path.GetFileName(filePath);
                string fileExtension = Path.GetExtension(fileName);
                
                Logger.Log($"[>] Upload SANS lifecycle: '{fileName}' -> {vaultFolderPath}", Logger.LogLevel.INFO);
                
                var targetFolder = EnsureVaultPathExists(vaultFolderPath, projectNumber, reference, module);
                
                if (targetFolder == null || targetFolder.Id <= 0)
                {
                    Logger.Log($"[-] Impossible de creer/acceder au dossier: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                // Verifier si fichier existe deja
                var existingFile = FindFileInFolder(targetFolder.Id, fileName);
                if (existingFile != null)
                {
                    Logger.Log($"   [i] Le fichier '{fileName}' existe deja", Logger.LogLevel.INFO);
                    
                    // Verifier le lifecycle (peut etre bloque)
                    ForceFileToWorkInProgress(existingFile.MasterId);
                    
                    // Appliquer les proprietes (peut echouer si lifecycle "Released")
                    return ApplyPropertiesWithoutLifecycle(existingFile.MasterId, projectNumber, reference, module);
                }

                ACW.FileClassification fileClass = DetermineFileClassification(fileExtension);
                long newFileId = -1;
                
                // Upload fichier
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var fileInfo = new FileInfo(filePath);

                    // Le 3eme parametre est le commentaire pour la version 1
                    // Documentation: "Text data to be associated with version 1 of the file"
                    var addedFile = _connection!.FileManager.AddFile(
                        targetFolder,
                        fileName,
                        string.Empty,  // Commentaire pour la version 1 (pas de commentaire pour cette methode)
                        fileInfo.LastWriteTime,
                        null,
                        null,
                        fileClass,
                        false,
                        fileStream
                    );
                    
                    if (addedFile != null && addedFile.EntityIterationId > 0)
                    {
                        var file = _connection.WebServiceManager.DocumentService.GetFileById(addedFile.EntityIterationId);
                        if (file != null)
                        {
                            newFileId = file.MasterId;
                            Logger.Log($"   [+] Fichier uploade (MasterId: {newFileId})", Logger.LogLevel.INFO);
                        }
                    }
                }

                if (newFileId <= 0)
                {
                    Logger.Log($"   [-] Echec upload", Logger.LogLevel.ERROR);
                    return false;
                }

                // CRITIQUE: Verifier le lifecycle (peut etre "Released" et bloquer les modifications)
                bool lifecycleOk = ForceFileToWorkInProgress(newFileId);
                
                if (!lifecycleOk)
                {
                    Logger.Log($"   [!] Le fichier peut etre bloque par le lifecycle 'Released'", Logger.LogLevel.WARNING);
                    Logger.Log($"   [i] Configurez Vault pour eviter l'assignation automatique du lifecycle", Logger.LogLevel.INFO);
                }
                
                // Appliquer les proprietes (peut echouer si lifecycle "Released")
                if (!string.IsNullOrEmpty(projectNumber) || !string.IsNullOrEmpty(reference) || !string.IsNullOrEmpty(module))
                {
                    Logger.Log($"   [~] Application des proprietes...", Logger.LogLevel.INFO);
                    bool propsApplied = ApplyPropertiesWithoutLifecycle(newFileId, projectNumber, reference, module);
                    
                    if (!propsApplied)
                    {
                        Logger.Log($"   [!] Echec application proprietes - peut etre du au lifecycle 'Released'", Logger.LogLevel.WARNING);
                    }
                    
                    return propsApplied;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogException($"UploadFileWithoutLifecycle({Path.GetFileName(filePath)})", ex, Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Force un fichier a l'etat "Work in Progress"
        /// 
        /// IMPORTANT: Les methodes de gestion du lifecycle ne sont pas disponibles dans cette version du SDK.
        /// La vraie solution est de configurer Vault pour ne pas assigner automatiquement le lifecycle "Released":
        /// 
        /// 1. Dans ADMS Console -> Behaviors -> File Lifecycle -> Default Lifecycle
        /// 2. Changer le lifecycle par defaut pour les extensions .ipt, .iam, .idw vers un lifecycle qui commence en "Work in Progress"
        /// 3. Ou desactiver l'assignation automatique de lifecycle
        /// 
        /// Cette methode verifie simplement si le fichier a un lifecycle et log l'information.
        /// </summary>
        private bool ForceFileToWorkInProgress(long fileMasterId)
        {
            try
            {
                Logger.Log($"   [~] Verification du lifecycle...", Logger.LogLevel.DEBUG);
                
                var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                if (latestFiles == null || latestFiles.Length == 0)
                {
                    return false;
                }
                
                var file = latestFiles[0];
                
                // Pas de lifecycle = OK
                if (file.FileLfCyc == null || file.FileLfCyc.LfCycDefId <= 0)
                {
                    Logger.Log($"   [+] Pas de lifecycle assigne - fichier modifiable", Logger.LogLevel.INFO);
                    return true;
                }
                
                Logger.Log($"   [!] Lifecycle detecte (Def ID: {file.FileLfCyc.LfCycDefId}, State ID: {file.FileLfCyc.LfCycStateId})", Logger.LogLevel.WARNING);
                Logger.Log($"   [i] SOLUTION: Configurer Vault pour ne pas assigner automatiquement le lifecycle 'Released'", Logger.LogLevel.INFO);
                Logger.Log($"      -> ADMS Console -> Behaviors -> File Lifecycle -> Default Lifecycle", Logger.LogLevel.INFO);
                Logger.Log($"      -> Changer le lifecycle par defaut vers 'Work in Progress' pour les fichiers Inventor", Logger.LogLevel.INFO);
                
                // Le fichier peut etre bloque par le lifecycle "Released"
                // Les methodes UpdateFileLifeCycleStates ne sont pas disponibles dans cette version du SDK
                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"   [!] Erreur verification lifecycle: {ex.Message}", Logger.LogLevel.WARNING);
                return false;
            }
        }

        /// <summary>
        /// Retire completement le lifecycle (non disponible dans cette version SDK)
        /// </summary>
        private bool RemoveLifecycleFromFile(long fileIterationId)
        {
            // Les methodes UpdateFileLifeCycleDefinitions ne sont pas disponibles dans cette version du SDK
            Logger.Log($"   [!] Retrait lifecycle non disponible dans cette version SDK", Logger.LogLevel.WARNING);
            Logger.Log($"   [i] Configurez Vault pour eviter l'assignation automatique du lifecycle", Logger.LogLevel.INFO);
            return false;
        }

        /// <summary>
        /// Applique les proprietes en gerant le lifecycle
        /// </summary>
        private bool ApplyPropertiesWithoutLifecycle(long fileMasterId, string? projectNumber, string? reference, string? module)
        {
            try
            {
                var properties = new List<ACW.PropInstParam>();
                
                // Project
                if (!string.IsNullOrEmpty(projectNumber) && _propertyDefIds.ContainsKey("Project"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Project"],
                        Val = projectNumber
                    });
                }
                
                // Reference
                if (!string.IsNullOrEmpty(reference) && _propertyDefIds.ContainsKey("Reference"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Reference"],
                        Val = reference
                    });
                }
                
                // Module
                if (!string.IsNullOrEmpty(module) && _propertyDefIds.ContainsKey("Module"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Module"],
                        Val = module
                    });
                }
                
                if (properties.Count == 0)
                {
                    return true;
                }

                var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                if (latestFiles == null || latestFiles.Length == 0)
                {
                    return false;
                }
                
                var fileInfo = latestFiles[0];

                // Annuler checkout si necessaire
                if (fileInfo.CheckedOut)
                {
                    try
                    {
                        _connection.WebServiceManager.DocumentService.UndoCheckoutFile(fileMasterId, out _);
                        System.Threading.Thread.Sleep(500);
                    }
                    catch
                    {
                        return false;
                    }
                }

                // Attendre disponibilite (10 secondes max)
                for (int i = 0; i < 10; i++)
                {
                    latestFiles = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                    if (latestFiles != null && latestFiles.Length > 0)
                    {
                        fileInfo = latestFiles[0];
                        if (!fileInfo.CheckedOut && !fileInfo.Locked)
                        {
                            break;
                        }
                    }
                    System.Threading.Thread.Sleep(1000);
                }

                // CheckOut
                bool checkedOut = false;
                try
                {
                    _connection.WebServiceManager.DocumentService.CheckoutFile(
                        fileMasterId,
                        ACW.CheckoutFileOptions.Master,
                        Environment.MachineName,
                        "c:\\temp",
                        "c:\\temp",
                        out _
                    );
                    checkedOut = true;
                }
                catch
                {
                    return false;
                }

                try
                {
                    // Recharger
                    latestFiles = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                    fileInfo = latestFiles[0];
                    
                    // Update Properties
                    var propArray = new ACW.PropInstParamArray
                    {
                        Items = properties.ToArray()
                    };
                    
                    _connection.WebServiceManager.DocumentService.UpdateFileProperties(
                        new[] { fileInfo.Id },
                        new[] { propArray }
                    );
                    
                    // CheckIn
                    var wsFile = _connection.WebServiceManager.DocumentService.GetFileById(fileInfo.Id);
                    var vaultFile = new VDF.Vault.Currency.Entities.FileIteration(_connection, wsFile);
                    
                    _connection.FileManager.CheckinFile(
                        vaultFile,
                        "Proprietes XNRGY",
                        false,
                        fileInfo.ModDate,
                        null,
                        null,
                        false,
                        null,
                        ACW.FileClassification.None,
                        false,
                        null
                    );
                    
                    checkedOut = false;
                    
                    // Re-forcer WIP apres CheckIn
                    System.Threading.Thread.Sleep(1000);
                    ForceFileToWorkInProgress(fileMasterId);
                    
                    Logger.Log($"   [+] {properties.Count} propriete(s) appliquee(s)", Logger.LogLevel.INFO);
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.LogException("ApplyProperties", ex, Logger.LogLevel.ERROR);
                    
                    // Rollback
                    if (checkedOut)
                    {
                        try
                        {
                            _connection.WebServiceManager.DocumentService.UndoCheckoutFile(fileMasterId, out _);
                        }
                        catch { }
                    }
                    
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException($"ApplyPropertiesWithoutLifecycle", ex, Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Recherche un fichier dans un dossier
        /// </summary>
        private ACW.File? FindFileInFolder(long folderId, string fileName)
        {
            try
            {
                var filesInFolder = _connection!.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folderId, false);
                return filesInFolder?.FirstOrDefault(f => f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Analyse un fichier dans Vault et retourne toutes ses informations (categorie, FileClassification, etc.)
        /// </summary>
        public void AnalyzeFileInVault(string vaultFolderPath, string fileName)
        {
            if (!IsConnected)
            {
                Logger.Log("[-] Non connecte a Vault", Logger.LogLevel.ERROR);
                return;
            }

            try
            {
                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("---------------------------------------------------------------", Logger.LogLevel.INFO);
                Logger.Log($"[>] ANALYSE DU FICHIER: {fileName}", Logger.LogLevel.INFO);
                Logger.Log($"   Chemin Vault: {vaultFolderPath}", Logger.LogLevel.INFO);
                Logger.Log("---------------------------------------------------------------", Logger.LogLevel.INFO);

                // Trouver le dossier
                var targetFolder = EnsureVaultPathExists(vaultFolderPath, null, null, null);
                if (targetFolder == null || targetFolder.Id <= 0)
                {
                    Logger.Log($"[-] Impossible de trouver le dossier: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return;
                }

                // Chercher le fichier
                var file = FindFileInFolder(targetFolder.Id, fileName);
                if (file == null)
                {
                    Logger.Log($"[-] Fichier '{fileName}' introuvable dans {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return;
                }

                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"[i] INFORMATIONS DU FICHIER:", Logger.LogLevel.INFO);
                Logger.Log($"   ID: {file.Id}", Logger.LogLevel.INFO);
                Logger.Log($"   MasterId: {file.MasterId}", Logger.LogLevel.INFO);
                Logger.Log($"   Nom: {file.Name}", Logger.LogLevel.INFO);
                Logger.Log($"   Version: {file.VerNum}", Logger.LogLevel.INFO);
                Logger.Log($"   Chemin complet: {file.Name ?? "N/A"}", Logger.LogLevel.INFO);

                // Analyser la categorie
                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"[i] CATEGORIE:", Logger.LogLevel.INFO);
                try
                {
                    bool categoryFound = false;
                    
                    // Methode 1: Essayer GetCategoryIdsByEntityMasterIds avec string
                    try
                    {
                        var method1 = _connection!.WebServiceManager.CategoryService.GetType().GetMethod("GetCategoryIdsByEntityMasterIds", new[] { typeof(long[]), typeof(string) });
                        if (method1 != null)
                        {
                            Logger.Log($"   [~] Tentative methode 1: GetCategoryIdsByEntityMasterIds(long[], string)", Logger.LogLevel.DEBUG);
                            var result1 = method1.Invoke(_connection.WebServiceManager.CategoryService, new object[] { new[] { file.MasterId }, "FILE" });
                            if (result1 is long[][] categoryIdsResult1 && categoryIdsResult1.Length > 0 && categoryIdsResult1[0] != null && categoryIdsResult1[0].Length > 0)
                            {
                                long categoryId = categoryIdsResult1[0][0];
                                var category = _connection.WebServiceManager.CategoryService.GetCategoryById(categoryId);
                                if (category != null)
                                {
                                    Logger.Log($"   [+] Category ID: {category.Id}", Logger.LogLevel.INFO);
                                    Logger.Log($"   [+] Category Name: '{category.Name}'", Logger.LogLevel.INFO);
                                    Logger.Log($"   [+] Category SystemName: '{category.SysName}'", Logger.LogLevel.INFO);
                                    categoryFound = true;
                                }
                            }
                        }
                    }
                    catch (Exception ex1)
                    {
                        Logger.Log($"   [i] Methode 1 echouee: {ex1.Message}", Logger.LogLevel.DEBUG);
                    }

                    // Methode 2: Essayer GetCategoryIdsByEntityMasterIds avec long[]
                    if (!categoryFound)
                    {
                        try
                        {
                            var method2 = _connection!.WebServiceManager.CategoryService.GetType().GetMethod("GetCategoryIdsByEntityMasterIds", new[] { typeof(long[]), typeof(long[]) });
                            if (method2 != null)
                            {
                                Logger.Log($"   [~] Tentative methode 2: GetCategoryIdsByEntityMasterIds(long[], long[])", Logger.LogLevel.DEBUG);
                                // EntityClassId pour FILE est generalement 1
                                var result2 = method2.Invoke(_connection.WebServiceManager.CategoryService, new object[] { new[] { file.MasterId }, new[] { 1L } });
                                if (result2 is long[][] categoryIdsResult2 && categoryIdsResult2.Length > 0 && categoryIdsResult2[0] != null && categoryIdsResult2[0].Length > 0)
                                {
                                    long categoryId = categoryIdsResult2[0][0];
                                    var category = _connection.WebServiceManager.CategoryService.GetCategoryById(categoryId);
                                    if (category != null)
                                    {
                                        Logger.Log($"   [+] Category ID: {category.Id}", Logger.LogLevel.INFO);
                                        Logger.Log($"   [+] Category Name: '{category.Name}'", Logger.LogLevel.INFO);
                                        Logger.Log($"   [+] Category SystemName: '{category.SysName}'", Logger.LogLevel.INFO);
                                        categoryFound = true;
                                    }
                                }
                            }
                        }
                        catch (Exception ex2)
                        {
                            Logger.Log($"   [i] Methode 2 echouee: {ex2.Message}", Logger.LogLevel.DEBUG);
                        }
                    }

                    // Methode 3: Essayer via FileIteration (si disponible)
                    if (!categoryFound)
                    {
                        try
                        {
                            Logger.Log($"   [~] Tentative methode 3: Via FileIteration", Logger.LogLevel.DEBUG);
                            var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, file);
                            if (fileIteration != null)
                            {
                                // Essayer d'acceder a la categorie via FileIteration
                                var categoryProp = fileIteration.GetType().GetProperty("Category");
                                if (categoryProp != null)
                                {
                                    var categoryObj = categoryProp.GetValue(fileIteration);
                                    if (categoryObj != null)
                                    {
                                        var catIdProp = categoryObj.GetType().GetProperty("Id");
                                        var catNameProp = categoryObj.GetType().GetProperty("Name");
                                        if (catIdProp != null && catNameProp != null)
                                        {
                                            long catId = (long)catIdProp.GetValue(categoryObj);
                                            string catName = catNameProp.GetValue(categoryObj)?.ToString() ?? "";
                                            Logger.Log($"   [+] Category ID: {catId}", Logger.LogLevel.INFO);
                                            Logger.Log($"   [+] Category Name: '{catName}'", Logger.LogLevel.INFO);
                                            categoryFound = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex3)
                        {
                            Logger.Log($"   [i] Methode 3 echouee: {ex3.Message}", Logger.LogLevel.DEBUG);
                        }
                    }

                    if (!categoryFound)
                    {
                        Logger.Log($"   [!] Aucune categorie assignee (toutes les methodes ont echoue)", Logger.LogLevel.WARNING);
                        Logger.Log($"   [i] Note: Le fichier peut avoir une categorie dans Vault Client mais non accessible via SDK", Logger.LogLevel.INFO);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   [!] Erreur lors de la recuperation de la categorie: {ex.Message}", Logger.LogLevel.WARNING);
                    Logger.Log($"   [i] Stack trace: {ex.StackTrace}", Logger.LogLevel.DEBUG);
                }

                // Obtenir le FileClassification depuis le fichier
                // Note: FileClassification n'est pas directement dans ACW.File, il faut le recuperer via FileManager
                try
                {
                    var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, file);
                    if (fileIteration != null)
                    {
                        // Le FileClassification est stocke dans l'entite FileIteration
                        var fileClass = fileIteration.FileClassification;
                        Logger.Log($"", Logger.LogLevel.INFO);
                        Logger.Log($"[i] FILECLASSIFICATION:", Logger.LogLevel.INFO);
                        Logger.Log($"   FileClassification: {fileClass}", Logger.LogLevel.INFO);
                        Logger.Log($"   FileClassification (int): {(int)fileClass}", Logger.LogLevel.INFO);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   [!] Impossible de recuperer FileClassification: {ex.Message}", Logger.LogLevel.WARNING);
                }

                // Analyser les proprietes (simplifie - afficher seulement les proprietes importantes)
                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"[i] PROPRIETES IMPORTANTES:", Logger.LogLevel.INFO);
                try
                {
                    // Utiliser GetLatestFilesByMasterIds qui retourne les proprietes directement dans l'objet File
                    var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { file.MasterId });
                    if (latestFiles != null && latestFiles.Length > 0)
                    {
                        var latestFile = latestFiles[0];
                        // Les proprietes sont accessibles via les methodes GetPropertyValue ou directement dans l'objet File
                        Logger.Log($"   Note: Utilisez Vault Client pour voir toutes les proprietes", Logger.LogLevel.INFO);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   [!] Erreur lors de la recuperation des proprietes: {ex.Message}", Logger.LogLevel.WARNING);
                }

                // Analyser le lifecycle
                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"[i] LIFECYCLE:", Logger.LogLevel.INFO);
                if (file.FileLfCyc != null)
                {
                    Logger.Log($"   Lifecycle Definition ID: {file.FileLfCyc.LfCycDefId}", Logger.LogLevel.INFO);
                    Logger.Log($"   Lifecycle State ID: {file.FileLfCyc.LfCycStateId}", Logger.LogLevel.INFO);
                    
                    try
                    {
                        // Recuperer toutes les definitions de lifecycle
                        var allLifecycleDefs = _connection!.WebServiceManager.LifeCycleService.GetAllLifeCycleDefinitions();
                        var lifecycleDef = allLifecycleDefs?.FirstOrDefault(l => l.Id == file.FileLfCyc.LfCycDefId);
                        
                        if (lifecycleDef != null)
                        {
                            Logger.Log($"   Lifecycle Definition Name: '{lifecycleDef.DispName}'", Logger.LogLevel.INFO);
                            
                            if (lifecycleDef.StateArray != null)
                            {
                                var currentState = lifecycleDef.StateArray.FirstOrDefault(s => s.Id == file.FileLfCyc.LfCycStateId);
                                if (currentState != null)
                                {
                                    Logger.Log($"   Lifecycle State Name: '{currentState.DispName}'", Logger.LogLevel.INFO);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"   [!] Erreur lors de la recuperation du lifecycle: {ex.Message}", Logger.LogLevel.WARNING);
                    }
                }
                else
                {
                    Logger.Log($"   [+] Aucun lifecycle assigne", Logger.LogLevel.INFO);
                }

                // Note: Les comportements (Behaviors) peuvent etre analyses via GetCategoryConfigurationById
                // mais necessitent des parametres specifiques selon la version du SDK

                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("---------------------------------------------------------------", Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.LogException($"AnalyzeFileInVault({fileName})", ex, Logger.LogLevel.ERROR);
            }
        }

        /// <summary>
        /// Telecharge (GET) un dossier Vault vers le working folder local
        /// Utilise pour la mise a jour du workspace (librairies, Content Center, etc.)
        /// </summary>
        /// <param name="vaultFolderPath">Chemin Vault (ex: $/Engineering/Library/Cabinet)</param>
        /// <returns>True si succes</returns>
        public async Task<bool> GetFolderAsync(string vaultFolderPath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    if (_connection == null)
                    {
                        Logger.Log($"[-] Non connecte a Vault", Logger.LogLevel.ERROR);
                        return false;
                    }

                    Logger.Log($"[>] GET: {vaultFolderPath}", Logger.LogLevel.INFO);

                    // Trouver le dossier dans Vault
                    var folder = _connection.WebServiceManager.DocumentService.GetFolderByPath(vaultFolderPath);
                    if (folder == null)
                    {
                        Logger.Log($"   [!] Dossier non trouve: {vaultFolderPath}", Logger.LogLevel.WARNING);
                        return false;
                    }

                    // Creer le VaultFolder
                    var vaultFolder = new VDF.Vault.Currency.Entities.Folder(_connection, folder);

                    // Obtenir tous les fichiers du dossier (recursif)
                    var files = _connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(
                        folder.Id, 
                        false  // false = recuperer aussi les sous-dossiers
                    );

                    if (files == null || files.Length == 0)
                    {
                        Logger.Log($"   [i] Aucun fichier dans {vaultFolderPath}", Logger.LogLevel.DEBUG);
                        return true;
                    }

                    Logger.Log($"   [~] {files.Length} fichiers a telecharger...", Logger.LogLevel.DEBUG);

                    // Telecharger les fichiers
                    int successCount = 0;
                    foreach (var file in files)
                    {
                        try
                        {
                            var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, file);
                            
                            var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                            downloadSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                            
                            var result = _connection.FileManager.AcquireFiles(downloadSettings);
                            
                            if (result.FileResults != null && result.FileResults.Any(r => r.LocalPath != null))
                            {
                                successCount++;
                            }
                        }
                        catch (Exception fileEx)
                        {
                            Logger.Log($"   [!] Erreur telechargement {file.Name}: {fileEx.Message}", Logger.LogLevel.DEBUG);
                        }
                    }

                    Logger.Log($"   [+] {successCount}/{files.Length} fichiers telecharges", Logger.LogLevel.INFO);
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Log($"[-] Erreur GET folder {vaultFolderPath}: {ex.Message}", Logger.LogLevel.ERROR);
                    return false;
                }
            });
        }

        #region Configuration File Management (VaultSettingsService Support)

        /// <summary>
        /// Recherche un fichier dans Vault par son chemin complet (ex: $/Engineering/Config/file.config)
        /// </summary>
        /// <param name="vaultFilePath">Chemin Vault complet du fichier</param>
        /// <returns>Objet File ou null si non trouve</returns>
        public ACW.File? FindFileByPath(string vaultFilePath)
        {
            if (!IsConnected || _connection == null || string.IsNullOrEmpty(vaultFilePath))
                return null;

            try
            {
                // Separer le chemin du dossier et le nom du fichier
                int lastSlash = vaultFilePath.LastIndexOf('/');
                if (lastSlash < 0) return null;

                string folderPath = vaultFilePath.Substring(0, lastSlash);
                string fileName = vaultFilePath.Substring(lastSlash + 1);

                Logger.Log($"[?] Recherche fichier: {fileName} dans {folderPath}", Logger.LogLevel.DEBUG);

                // Trouver le dossier
                var folder = _connection.WebServiceManager.DocumentService.GetFolderByPath(folderPath);
                if (folder == null)
                {
                    Logger.Log($"[i] Dossier non trouve: {folderPath}", Logger.LogLevel.DEBUG);
                    return null;
                }

                // Chercher le fichier dans le dossier
                var files = _connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, false);
                var file = files?.FirstOrDefault(f => f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase));

                if (file != null)
                {
                    Logger.Log($"[+] Fichier trouve: {file.Name} (ID: {file.Id})", Logger.LogLevel.DEBUG);
                }
                else
                {
                    Logger.Log($"[i] Fichier non trouve: {fileName}", Logger.LogLevel.DEBUG);
                }

                return file;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur FindFileByPath: {ex.Message}", Logger.LogLevel.DEBUG);
                return null;
            }
        }

        /// <summary>
        /// Telecharge un fichier depuis Vault vers un emplacement local
        /// </summary>
        /// <param name="file">Fichier Vault a telecharger</param>
        /// <param name="localPath">Chemin local de destination</param>
        /// <param name="checkout">True pour CheckOut, False pour simple Get</param>
        /// <returns>True si succes</returns>
        public bool AcquireFile(ACW.File file, string localPath, bool checkout = false)
        {
            if (!IsConnected || _connection == null || file == null)
                return false;

            try
            {
                var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, file);
                
                var settings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                var option = checkout 
                    ? VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout 
                    : VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                    
                settings.AddFileToAcquire(fileIteration, option);
                
                // Definir le dossier de destination
                string? destFolder = Path.GetDirectoryName(localPath);
                if (!string.IsNullOrEmpty(destFolder) && !Directory.Exists(destFolder))
                {
                    Directory.CreateDirectory(destFolder);
                }
                settings.LocalPath = new VDF.Currency.FolderPathAbsolute(destFolder ?? "");

                var result = _connection.FileManager.AcquireFiles(settings);
                
                if (result.FileResults != null && result.FileResults.Any(r => r.LocalPath != null))
                {
                    // Copier vers le chemin exact demande si different
                    var downloadedPath = result.FileResults.First().LocalPath?.FullPath;
                    if (!string.IsNullOrEmpty(downloadedPath) && downloadedPath != localPath && System.IO.File.Exists(downloadedPath))
                    {
                        System.IO.File.Copy(downloadedPath, localPath, true);
                    }
                    
                    Logger.Log($"[+] Fichier telecharge: {localPath}", Logger.LogLevel.DEBUG);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur AcquireFile: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Met a jour un fichier existant dans Vault (CheckOut, Update, CheckIn)
        /// Utilise le FileManager pour une gestion correcte du workflow Vault
        /// </summary>
        public bool UpdateFileInVault(ACW.File file, string localFilePath, string comment)
        {
            if (!IsConnected || _connection == null || file == null || !System.IO.File.Exists(localFilePath))
                return false;

            ACW.File? checkedOutFile = null;
            try
            {
                Logger.Log($"[>] Mise a jour fichier Vault: {file.Name}", Logger.LogLevel.DEBUG);

                // 1. CheckOut via DocumentService
                checkedOutFile = _connection.WebServiceManager.DocumentService.CheckoutFile(
                    file.Id,
                    ACW.CheckoutFileOptions.Master,
                    Environment.MachineName,
                    localFilePath,
                    comment,
                    out var downloadTicket);

                if (checkedOutFile == null)
                {
                    Logger.Log("[-] CheckOut a retourne null", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"   [+] CheckOut reussi (FileId: {checkedOutFile.Id})", Logger.LogLevel.DEBUG);

                // 2. CheckIn avec le fichier local mis a jour
                var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, checkedOutFile);
                
                _connection.FileManager.CheckinFile(
                    fileIteration,
                    comment,
                    false,  // keepCheckedOut
                    new FileInfo(localFilePath).LastWriteTimeUtc,
                    null,   // associations
                    null,   // bom
                    false,  // copyBom
                    null,   // filePathAbsolute - utilise le fichier local existant
                    ACW.FileClassification.None,
                    false,  // hidden
                    null    // dependencyOptions
                );

                Logger.Log($"[+] Fichier mis a jour dans Vault: {file.Name}", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur UpdateFileInVault: {ex.Message}", Logger.LogLevel.ERROR);
                
                // Tenter UndoCheckOut en cas d'erreur
                try
                {
                    if (checkedOutFile != null)
                    {
                        _connection?.WebServiceManager.DocumentService.UndoCheckoutFile(checkedOutFile.MasterId, out _);
                        Logger.Log("[i] UndoCheckOut effectue", Logger.LogLevel.DEBUG);
                    }
                }
                catch { /* Ignorer */ }
                
                return false;
            }
        }

        /// <summary>
        /// Ajoute un nouveau fichier dans Vault via FileManager.AddFile
        /// </summary>
        public bool AddFileToVault(string localFilePath, string vaultFolderPath, string comment)
        {
            if (!IsConnected || _connection == null || !System.IO.File.Exists(localFilePath))
                return false;

            try
            {
                string fileName = Path.GetFileName(localFilePath);
                Logger.Log($"[>] Ajout fichier dans Vault: {vaultFolderPath}/{fileName}", Logger.LogLevel.DEBUG);

                // Utiliser EnsureVaultPathExists qui gere correctement la creation recursive des dossiers
                var folder = EnsureVaultPathExists(vaultFolderPath);
                if (folder == null)
                {
                    Logger.Log($"[-] Impossible de creer/trouver le dossier: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                // Lire le fichier
                var fileInfo = new FileInfo(localFilePath);

                // Upload vers Vault via FileManager
                using (var stream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read))
                {
                    var uploadResult = _connection.FileManager.AddFile(
                        folder,
                        fileName,
                        comment,
                        fileInfo.LastWriteTimeUtc,
                        null,  // associations
                        null,  // bom
                        ACW.FileClassification.None,
                        false, // hidden
                        stream);

                    if (uploadResult != null)
                    {
                        Logger.Log($"[+] Fichier ajoute dans Vault: {vaultFolderPath}/{fileName}", Logger.LogLevel.INFO);
                        return true;
                    }
                }

                Logger.Log("[-] AddFile n'a pas retourne de resultat", Logger.LogLevel.ERROR);
                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur AddFileToVault: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Cree un dossier Vault de maniere recursive
        /// </summary>
        private ACW.Folder? CreateVaultFolderRecursive(string vaultFolderPath)
        {
            try
            {
                Logger.Log($"[>] CreateVaultFolderRecursive: {vaultFolderPath}", Logger.LogLevel.DEBUG);
                
                // Verifier si le dossier existe deja
                try
                {
                    var existing = _connection?.WebServiceManager.DocumentService.GetFolderByPath(vaultFolderPath);
                    if (existing != null)
                    {
                        Logger.Log($"[i] Dossier existe deja: {vaultFolderPath}", Logger.LogLevel.DEBUG);
                        return existing;
                    }
                }
                catch
                {
                    // Dossier n'existe pas - continuer pour le creer
                    Logger.Log($"[i] Dossier inexistant, creation necessaire: {vaultFolderPath}", Logger.LogLevel.DEBUG);
                }

                // Trouver le parent
                int lastSlash = vaultFolderPath.LastIndexOf('/');
                if (lastSlash <= 0)
                {
                    // C'est la racine $
                    try
                    {
                        return _connection?.WebServiceManager.DocumentService.GetFolderByPath("$");
                    }
                    catch
                    {
                        return null;
                    }
                }

                string parentPath = vaultFolderPath.Substring(0, lastSlash);
                string folderName = vaultFolderPath.Substring(lastSlash + 1);
                
                Logger.Log($"[i] Parent: {parentPath}, Nouveau dossier: {folderName}", Logger.LogLevel.DEBUG);

                // Recursion pour creer le parent si necessaire
                ACW.Folder? parentFolder = null;
                try
                {
                    parentFolder = _connection?.WebServiceManager.DocumentService.GetFolderByPath(parentPath);
                    Logger.Log($"[i] Parent trouve: {parentPath}", Logger.LogLevel.DEBUG);
                }
                catch
                {
                    // Parent n'existe pas - creer recursivement
                    Logger.Log($"[>] Creation recursive du parent: {parentPath}", Logger.LogLevel.DEBUG);
                    parentFolder = CreateVaultFolderRecursive(parentPath);
                }
                
                if (parentFolder == null)
                {
                    Logger.Log($"[-] Parent introuvable: {parentPath}", Logger.LogLevel.ERROR);
                    return null;
                }

                // Creer le dossier
                Logger.Log($"[>] Creation du dossier '{folderName}' dans {parentPath}...", Logger.LogLevel.DEBUG);
                var newFolder = _connection?.WebServiceManager.DocumentService.AddFolder(
                    folderName, 
                    parentFolder.Id, 
                    false); // isLibrary

                if (newFolder != null)
                {
                    Logger.Log($"[+] Dossier Vault cree: {vaultFolderPath}", Logger.LogLevel.INFO);
                }
                return newFolder;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur creation dossier Vault '{vaultFolderPath}': {ex.Message}", Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// CheckOut un fichier Vault pour mise a jour (rend le fichier local writable)
        /// </summary>
        public ACW.File? CheckOutFileForUpdate(ACW.File file, string localFilePath)
        {
            if (!IsConnected || _connection == null || file == null)
                return null;

            try
            {
                Logger.Log($"[>] CheckOut fichier: {file.Name}", Logger.LogLevel.DEBUG);

                var checkedOutFile = _connection.WebServiceManager.DocumentService.CheckoutFile(
                    file.Id,
                    ACW.CheckoutFileOptions.Master,
                    Environment.MachineName,
                    localFilePath,
                    "CheckOut pour mise a jour",
                    out var downloadTicket);

                if (checkedOutFile != null)
                {
                    Logger.Log($"[+] CheckOut reussi: {file.Name} (FileId: {checkedOutFile.Id})", Logger.LogLevel.DEBUG);
                    
                    // S'assurer que le fichier local n'est pas read-only apres checkout
                    if (System.IO.File.Exists(localFilePath))
                    {
                        var fileInfo = new FileInfo(localFilePath);
                        if (fileInfo.IsReadOnly)
                        {
                            fileInfo.IsReadOnly = false;
                            Logger.Log("[i] Flag read-only enleve apres CheckOut", Logger.LogLevel.DEBUG);
                        }
                    }
                }

                return checkedOutFile;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur CheckOut: {ex.Message}", Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// CheckIn un fichier apres modification locale - uploade le contenu modifie vers Vault
        /// </summary>
        public bool CheckInFile(ACW.File checkedOutFile, string localFilePath, string comment)
        {
            if (!IsConnected || _connection == null || checkedOutFile == null || !System.IO.File.Exists(localFilePath))
                return false;

            try
            {
                Logger.Log($"[>] CheckIn fichier: {checkedOutFile.Name}", Logger.LogLevel.DEBUG);
                Logger.Log($"   [i] Fichier local: {localFilePath}", Logger.LogLevel.DEBUG);

                var fileInfo = new FileInfo(localFilePath);
                Logger.Log($"   [i] Taille fichier: {fileInfo.Length} bytes", Logger.LogLevel.DEBUG);

                // Utiliser FileManager.CheckinFile avec le chemin du fichier local
                var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, checkedOutFile);
                
                // CheckIn avec upload du fichier local modifie
                // Le parametre filePathAbsolute doit contenir le chemin du fichier a uploader
                _connection.FileManager.CheckinFile(
                    fileIteration,
                    comment,
                    false,  // keepCheckedOut
                    fileInfo.LastWriteTimeUtc,
                    null,   // associations
                    null,   // bom
                    false,  // copyBom
                    localFilePath,  // filePathAbsolute - chemin du fichier local a uploader
                    ACW.FileClassification.None,
                    false,  // hidden
                    null    // dependencyOptions
                );

                Logger.Log($"[+] CheckIn reussi: {checkedOutFile.Name}", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur CheckIn: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Annule un CheckOut en cours
        /// </summary>
        public bool UndoCheckOut(ACW.File checkedOutFile)
        {
            if (!IsConnected || _connection == null || checkedOutFile == null)
                return false;

            try
            {
                Logger.Log($"[>] UndoCheckOut fichier: {checkedOutFile.Name}", Logger.LogLevel.DEBUG);
                
                _connection.WebServiceManager.DocumentService.UndoCheckoutFile(checkedOutFile.MasterId, out _);
                
                Logger.Log($"[+] UndoCheckOut reussi: {checkedOutFile.Name}", Logger.LogLevel.DEBUG);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur UndoCheckOut: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Supprime un fichier de Vault (pour fichiers de configuration seulement)
        /// </summary>
        public bool DeleteFileFromVault(ACW.File file)
        {
            if (!IsConnected || _connection == null || file == null)
                return false;

            try
            {
                Logger.Log($"[>] Suppression fichier Vault: {file.Name}", Logger.LogLevel.DEBUG);
                
                // Si le fichier est checke out, annuler le checkout d'abord
                if (file.CheckedOut)
                {
                    try
                    {
                        Logger.Log("[i] Fichier checke out - UndoCheckout...", Logger.LogLevel.DEBUG);
                        _connection.WebServiceManager.DocumentService.UndoCheckoutFile(file.MasterId, out _);
                    }
                    catch (Exception undoEx)
                    {
                        Logger.Log($"[!] UndoCheckout failed: {undoEx.Message}", Logger.LogLevel.WARNING);
                    }
                }
                
                // Obtenir le folder ID du fichier
                var folder = _connection.WebServiceManager.DocumentService.GetFolderById(file.FolderId);
                if (folder == null)
                {
                    Logger.Log("[-] Dossier parent introuvable", Logger.LogLevel.ERROR);
                    return false;
                }
                
                // Supprimer le fichier du dossier
                _connection.WebServiceManager.DocumentService.DeleteFileFromFolder(file.MasterId, folder.Id);
                
                Logger.Log($"[+] Fichier supprime de Vault: {file.Name}", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur suppression fichier Vault: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        #endregion
    }
}
