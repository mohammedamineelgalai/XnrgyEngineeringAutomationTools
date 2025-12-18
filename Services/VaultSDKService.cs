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
    /// Structure pour stocker les propri�t�s � appliquer en diff�r� (apr�s Job Processor)
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
        
        // Propri�t�s de connexion
        public string? VaultName { get; private set; }
        public string? UserName { get; private set; }
        public string? ServerName { get; private set; }
        
        // -------------------------------------------------------------------------------
        // ? FILE D'ATTENTE pour les propri�t�s � appliquer apr�s que le Job Processor ait fini
        // -------------------------------------------------------------------------------
        private readonly List<PendingPropertyUpdate> _pendingPropertyUpdates = new();
        private readonly object _pendingLock = new object();

        public bool IsConnected => _connection != null;
        
        /// <summary>
        /// Nombre de propri�t�s en attente d'application
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
        /// Ajoute une mise � jour de propri�t�s � la file d'attente (pour application apr�s Job Processor)
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
                Logger.Log($"   ?? Propri�t�s ajout�es � la file d'attente: {fileName} (MasterId: {fileMasterId})", Logger.LogLevel.DEBUG);
            }
        }

        /// <summary>
        /// Vide la file d'attente des propri�t�s
        /// </summary>
        public void ClearPendingPropertyUpdates()
        {
            lock (_pendingLock)
            {
                _pendingPropertyUpdates.Clear();
                Logger.Log($"   ??? File d'attente des propri�t�s vid�e", Logger.LogLevel.DEBUG);
            }
        }

        /// <summary>
        /// ? Applique toutes les propri�t�s en attente (� appeler APR�S tous les uploads)
        /// Le Job Processor aura eu le temps de traiter tous les fichiers
        /// </summary>
        /// <param name="waitBeforeStart">D�lai initial d'attente en secondes (par d�faut 0s - pas d'attente)</param>
        /// <returns>(succ�s, �checs)</returns>
        public (int successCount, int failedCount) ApplyPendingPropertyUpdates(int waitBeforeStart = 0)
        {
            List<PendingPropertyUpdate> updates;
            lock (_pendingLock)
            {
                if (_pendingPropertyUpdates.Count == 0)
                {
                    Logger.Log("   ? Aucune propri�t� en attente", Logger.LogLevel.INFO);
                    return (0, 0);
                }
                
                updates = new List<PendingPropertyUpdate>(_pendingPropertyUpdates);
                _pendingPropertyUpdates.Clear();
            }

            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);
            Logger.Log($"?? APPLICATION DES PROPRI�T�S EN BATCH ({updates.Count} fichiers)", Logger.LogLevel.INFO);
            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);
            
            // Attendre que le Job Processor ait fini avec les fichiers
            if (waitBeforeStart > 0)
            {
                Logger.Log($"   ? Attente de {waitBeforeStart}s pour le Job Processor...", Logger.LogLevel.INFO);
                System.Threading.Thread.Sleep(waitBeforeStart * 1000);
            }

            int successCount = 0;
            int failedCount = 0;

            foreach (var update in updates)
            {
                Logger.Log($"   ?? [{successCount + failedCount + 1}/{updates.Count}] {update.FileName}...", Logger.LogLevel.INFO);
                
                bool success = SetFilePropertiesDirectly(update.FileMasterId, update.ProjectNumber, update.Reference, update.Module);
                
                if (success)
                {
                    successCount++;
                    Logger.Log($"      ? Propri�t�s appliqu�es", Logger.LogLevel.INFO);
                }
                else
                {
                    failedCount++;
                    Logger.Log($"      ?? �chec (Job Processor occup� ou autre erreur)", Logger.LogLevel.WARNING);
                }
            }

            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);
            Logger.Log($"?? R�SULTAT: {successCount} r�ussi(s), {failedCount} �chec(s)", Logger.LogLevel.INFO);
            Logger.Log($"---------------------------------------------------------------", Logger.LogLevel.INFO);

            return (successCount, failedCount);
        }

        /// <summary>
        /// Soumet un Job de synchronisation des propri�t�s (Vault ? iProperties) via le Job Processor.
        /// Le Job Processor doit �tre configur� avec le handler "Autodesk.Vault.SyncProperties".
        /// </summary>
        /// <param name="fileVersionId">ID de la version du fichier (FileId, pas MasterId)</param>
        /// <param name="fileName">Nom du fichier (pour le log)</param>
        /// <returns>True si le job a �t� soumis avec succ�s</returns>
        public bool SubmitSyncPropertiesJob(long fileVersionId, string fileName)
        {
            if (_connection == null)
            {
                Logger.Log($"   ?? [JOB] Non connect� � Vault", Logger.LogLevel.WARNING);
                return false;
            }

            try
            {
                // Le type de job pour synchroniser les propri�t�s Vault ? iProperties
                // C'est le handler standard de Vault pour "Property Sync"
                const string JOB_TYPE = "Autodesk.Vault.SyncProperties";
                
                // Le job attend FileVersionId (pas FileMasterId)
                // Erreur si on utilise FileMasterId: "Missing required parameter. Requires either FileVersionIds or FileVersionId."
                var jobParams = new ACW.JobParam[]
                {
                    new ACW.JobParam { Name = "FileVersionId", Val = fileVersionId.ToString() }
                };

                // Soumettre le job avec priorit� normale (10)
                var job = _connection.WebServiceManager.JobService.AddJob(
                    JOB_TYPE,
                    $"Sync properties for {fileName}",
                    jobParams,
                    10  // Priorit� (1 = haute, 10 = normale)
                );

                if (job != null && job.Id > 0)
                {
                    Logger.Log($"   ? [JOB] Sync properties soumis (JobId: {job.Id}) pour {fileName}", Logger.LogLevel.INFO);
                    return true;
                }
                else
                {
                    Logger.Log($"   ?? [JOB] Job retourn� null ou ID invalide", Logger.LogLevel.WARNING);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   ?? [JOB] Erreur soumission: {ex.Message}", Logger.LogLevel.WARNING);
                // Ne pas faire �chouer le process principal si le job �choue
                return false;
            }
        }

        public bool Connect(string server, string vaultName, string username, string password)
        {
            try
            {
                Logger.Log($"?? Tentative de connexion � Vault...", Logger.LogLevel.INFO);
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
                    
                    Logger.Log($"? Connect� au Vault '{vaultName}' sur '{server}'", Logger.LogLevel.INFO);
                    Logger.Log($"   Utilisateur: {username}", Logger.LogLevel.INFO);
                    Logger.Log($"   Dossier racine: {_connection.FolderManager.RootFolder.FullName}", Logger.LogLevel.INFO);
                    
                    // Log des versions pour diagnostiquer les incompatibilit�s
                    try
                    {
                        // Version du SDK client (via Assembly)
                        var sdkAssembly = typeof(VDF.Vault.Currency.Connections.Connection).Assembly;
                        var sdkVersion = sdkAssembly.GetName().Version;
                        Logger.Log($"   ?? SDK Client Version: {sdkVersion}", Logger.LogLevel.INFO);
                        
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
                                    Logger.Log($"   ?? Serveur Vault: Informations de session disponibles", Logger.LogLevel.DEBUG);
                                }
                            }
                            catch (Exception verEx)
                            {
                                Logger.Log($"   ?? Impossible de r�cup�rer la version serveur: {verEx.Message}", Logger.LogLevel.DEBUG);
                            }
                        }
                        
                        // Version de l'application
                        var appAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                        var appVersion = appAssembly.GetName().Version;
                        Logger.Log($"   ?? Application Version: {appVersion}", Logger.LogLevel.DEBUG);
                    }
                    catch (Exception verEx)
                    {
                        Logger.Log($"   ?? Erreur lors de la r�cup�ration des versions: {verEx.Message}", Logger.LogLevel.DEBUG);
                    }
                    
                    LoadPropertyDefinitions();
                    LoadCategories();
                    
                    // V�rifier si writeback activ� (info seulement)
                    try
                    {
                        bool writebackEnabled = _connection?.WebServiceManager?.DocumentService
                            ?.GetEnableItemPropertyWritebackToFiles() ?? false;
                        Logger.Log($"   Writeback vers fichiers: {(writebackEnabled ? "? ACTIV�" : "?? D�SACTIV�")}", 
                            Logger.LogLevel.INFO);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"   ?? Impossible de v�rifier le writeback: {ex.Message}", Logger.LogLevel.DEBUG);
                    }
                    
                    // NOTE: ListRecentJobs() supprim� - �tait du diagnostic non n�cessaire
                    
                    return true;
                }
                else
                {
                    string errorMsg = "�chec d'authentification";
                    if (result.ErrorMessages != null && result.ErrorMessages.Count > 0)
                    {
                        errorMsg = result.ErrorMessages.First().Value;
                    }
                    Logger.Log($"? �chec connexion: {errorMsg}", Logger.LogLevel.ERROR);
                    
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

        private void LoadPropertyDefinitions()
        {
            try
            {
                Logger.Log("?? Chargement des Property Definitions...", Logger.LogLevel.DEBUG);
                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("--- LISTE COMPL�TE DES PROPRI�T�S FILES ---", Logger.LogLevel.INFO);
                
                var filePropDefs = _connection!.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                
                if (filePropDefs != null)
                {
                    foreach (var propDef in filePropDefs)
                    {
                        // Stocker par SysName ET par DispName pour flexibilit�
                        _propertyDefIds[propDef.SysName] = propDef.Id;
                        
                        // IMPORTANT: Aussi stocker par DispName pour propri�t�s avec GUID comme SysName
                        if (!string.IsNullOrEmpty(propDef.DispName) && propDef.DispName != propDef.SysName)
                        {
                            _propertyDefIds[propDef.DispName] = propDef.Id;
                        }
                        
                        // Log d�taill� pour diagnostic
                        string displayInfo = $"ID:{propDef.Id,4} | SysName: '{propDef.SysName}' | DispName: '{propDef.DispName}'";
                        
                        // Mettre en �vidence les propri�t�s XNRGY
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
                Logger.Log($"? {_propertyDefIds.Count} Property Definitions charg�es", Logger.LogLevel.DEBUG);
                
                // V�rifier si les propri�t�s XNRGY existent (chercher par SysName OU DispName)
                bool hasProject = _propertyDefIds.ContainsKey("Project");
                bool hasReference = _propertyDefIds.ContainsKey("Reference");
                bool hasModule = _propertyDefIds.ContainsKey("Module");
                
                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("?? V�RIFICATION DES PROPRI�T�S XNRGY:", Logger.LogLevel.INFO);
                Logger.Log($"   Project   : {(hasProject ? "? TROUV� (ID: " + _propertyDefIds["Project"] + ")" : "? NON TROUV�")}", Logger.LogLevel.INFO);
                Logger.Log($"   Reference : {(hasReference ? "? TROUV� (ID: " + _propertyDefIds["Reference"] + ")" : "? NON TROUV�")}", Logger.LogLevel.INFO);
                Logger.Log($"   Module    : {(hasModule ? "? TROUV� (ID: " + _propertyDefIds["Module"] + ")" : "? NON TROUV�")}", Logger.LogLevel.INFO);
                Logger.Log("", Logger.LogLevel.INFO);
                
                if (!hasProject || !hasReference || !hasModule)
                {
                    Logger.Log("?????? ATTENTION: Propri�t�s XNRGY manquantes ??????", Logger.LogLevel.WARNING);
                    Logger.Log("", Logger.LogLevel.WARNING);
                    Logger.Log("?? V�RIFIEZ dans Vault Admin Console:", Logger.LogLevel.WARNING);
                    Logger.Log("   1. Ouvrir ADMS Console", Logger.LogLevel.WARNING);
                    Logger.Log("   2. Behaviors > Properties > File Properties", Logger.LogLevel.WARNING);
                    Logger.Log("   3. V�rifier que les propri�t�s existent avec:", Logger.LogLevel.WARNING);
                    Logger.Log("      - Display Name: 'Project', 'Reference', 'Module'", Logger.LogLevel.WARNING);
                    Logger.Log("      - Apply to: ? Files", Logger.LogLevel.WARNING);
                    Logger.Log("   4. Regarder le log ci-dessus pour voir les noms exacts", Logger.LogLevel.WARNING);
                    Logger.Log("", Logger.LogLevel.WARNING);
                    Logger.Log("?? L'upload continuera SANS les propri�t�s manquantes", Logger.LogLevel.WARNING);
                    Logger.Log("", Logger.LogLevel.WARNING);
                }
                else
                {
                    Logger.Log("? Toutes les propri�t�s XNRGY sont d�tect�es!", Logger.LogLevel.INFO);
                    Logger.Log("", Logger.LogLevel.INFO);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("LoadPropertyDefinitions", ex, Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// D�finit les permissions d'�criture pour l'utilisateur connect� sur un fichier.
        /// NOTE: Cette fonctionnalit� n�cessite des m�thodes avanc�es de l'API SecurityService
        /// qui ne sont pas disponibles dans toutes les versions du SDK.
        /// Pour d�finir les permissions, utilisez l'interface Vault ou configurez les permissions
        /// au niveau du dossier parent.
        /// </summary>
        /// <param name="fileMasterId">L'ID Master du fichier</param>
        private void SetFileWritePermissions(long fileMasterId)
        {
            // TODO: Impl�menter la d�finition des permissions via l'API Vault SDK
            // Les m�thodes n�cessaires (SetEntACLs, GetACEsByACLId, etc.) ne sont pas disponibles
            // dans la version actuelle du SDK ou n�cessitent des permissions administrateur.
            // 
            // Solution alternative: Configurer les permissions au niveau du dossier parent dans Vault
            // pour que les fichiers upload�s h�ritent des bonnes permissions.
            Logger.Log($"   ?? D�finition des permissions d'�criture non impl�ment�e (limitation API)", Logger.LogLevel.WARNING);
            Logger.Log($"   ?? Solution: Configurez les permissions au niveau du dossier parent dans Vault", Logger.LogLevel.INFO);
        }

        /// <summary>
        /// Restaure la s�curit� h�rit�e sur un dossier pour permettre la cr�ation d'�l�ments enfants.
        /// Cette m�thode r�sout le probl�me Error 1000 caus� par les permissions "Overridden".
        /// </summary>
        /// <param name="folderId">L'ID du dossier sur lequel restaurer la s�curit� h�rit�e</param>
        private void RestoreInheritedSecurity(long folderId)
        {
            try
            {
                Logger.Log($"      ?? Restauration s�curit� h�rit�e pour dossier ID {folderId}...", Logger.LogLevel.DEBUG);
                
                // Obtenir les ACL actuelles du dossier
                var result = _connection!.WebServiceManager.SecurityService.GetEntACLsByEntityIds(new[] { folderId });
                
                if (result != null && result.ACLArray != null && result.ACLArray.Length > 0)
                {
                    var currentACL = result.ACLArray[0];
                    
                    // Appliquer le comportement "Combined" (h�ritage) au lieu de "Override"
                    // Combined = 1 signifie que la s�curit� est h�rit�e du parent
                    _connection.WebServiceManager.SecurityService.SetSystemACLs(
                        new[] { folderId },
                        currentACL.Id,
                        ACW.SysAclBeh.Combined  // ? CLE: Utilisera l'h�ritage au lieu de Override
                    );
                    
                    Logger.Log($"      ? S�curit� h�rit�e restaur�e (ACL ID: {currentACL.Id}, Beh: Combined)", Logger.LogLevel.DEBUG);
                }
                // Note: Pas de warning si ACL vide - le dossier h�rite automatiquement
            }
            catch (Exception ex)
            {
                // Log mais continue - le dossier a peut-�tre d�j� la s�curit� h�rit�e
                Logger.Log($"      ?? RestoreInheritedSecurity a �chou� (peut-�tre d�j� h�rit�): {ex.Message}", Logger.LogLevel.TRACE);
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
                    Logger.Log($"?? Chemin invalide (doit commencer par $/): {vaultFolderPath}", Logger.LogLevel.WARNING);
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
                            Logger.Log($"   ? {nextPath}", Logger.LogLevel.TRACE);
                        }
                    }
                    catch
                    {
                    }

                    if (nextFolder == null)
                    {
                        Logger.Log($"   + Cr�ation: {nextPath}", Logger.LogLevel.DEBUG);
                        
                        // Restaurer s�curit� h�rit�e sur parent AVANT cr�ation
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
                                Logger.Log($"   ? Cr��: {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                            }
                            else
                            {
                                Logger.Log($"   ? �chec cr�ation: {nextPath}", Logger.LogLevel.ERROR);
                                return null;
                            }
                        }
                        catch (ACW.VaultServiceErrorException ex) when (ex.ErrorCode == 1000)
                        {
                            // Retry une fois apr�s avoir fix� la s�curit�
                            Logger.Log($"   ?? Error 1000 - Nouvelle tentative...", Logger.LogLevel.WARNING);
                            
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
                                    Logger.Log($"   ? Cr�� (2�me tentative): {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                                }
                                else
                                {
                                    Logger.Log($"   ? �chec cr�ation (2�me tentative): {nextPath}", Logger.LogLevel.ERROR);
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
                            // Erreur 1011: Le dossier existe d�j� - r�cup�rer simplement
                            Logger.Log($"   ?? Dossier existe d�j� (erreur 1011), r�cup�ration...", Logger.LogLevel.DEBUG);
                            try
                            {
                                var folders = _connection.WebServiceManager.DocumentService.FindFoldersByPaths(new[] { nextPath });
                                if (folders != null && folders.Length > 0 && folders[0] != null && folders[0].Id > 0)
                                {
                                    nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, folders[0]);
                                    Logger.Log($"   ? Dossier r�cup�r�: {nextPath} (ID: {folders[0].Id})", Logger.LogLevel.DEBUG);
                                }
                                else
                                {
                                    // Fallback: chercher dans le dossier parent
                                    var parentFolders = _connection.WebServiceManager.DocumentService.GetFoldersByParentId(currentFolder!.Id, false);
                                    var existingFolder = parentFolders?.FirstOrDefault(f => f.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));
                                    if (existingFolder != null && existingFolder.Id > 0)
                                    {
                                        nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, existingFolder);
                                        Logger.Log($"   ? Dossier r�cup�r� via parent: {nextPath} (ID: {existingFolder.Id})", Logger.LogLevel.DEBUG);
                                    }
                                    else
                                    {
                                        // Le dossier n'existe pas vraiment - peut-�tre une fausse alerte ou race condition
                                        // R�essayer de cr�er le dossier
                                        Logger.Log($"   ?? Dossier non trouv� malgr� erreur 1011, nouvelle tentative de cr�ation...", Logger.LogLevel.WARNING);
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
                                                Logger.Log($"   ? Dossier cr�� (2�me tentative): {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                                            }
                                            else
                                            {
                                                Logger.Log($"   ? �chec cr�ation (2�me tentative): {nextPath}", Logger.LogLevel.ERROR);
                                                return null;
                                            }
                                        }
                                        catch (ACW.VaultServiceErrorException retryEx) when (retryEx.ErrorCode == 1011)
                                        {
                                            // Encore l'erreur 1011 - le dossier existe vraiment maintenant
                                            Logger.Log($"   ?? Erreur 1011 persistante, r�cup�ration finale...", Logger.LogLevel.DEBUG);
                                            var finalFolders = _connection.WebServiceManager.DocumentService.FindFoldersByPaths(new[] { nextPath });
                                            if (finalFolders != null && finalFolders.Length > 0 && finalFolders[0] != null && finalFolders[0].Id > 0)
                                            {
                                                nextFolder = new VDF.Vault.Currency.Entities.Folder(_connection, finalFolders[0]);
                                                Logger.Log($"   ? Dossier r�cup�r� (final): {nextPath} (ID: {finalFolders[0].Id})", Logger.LogLevel.DEBUG);
                                            }
                                            else
                                            {
                                                Logger.Log($"   ? Impossible de r�cup�rer le dossier existant: {nextPath}", Logger.LogLevel.ERROR);
                                                return null;
                                            }
                                        }
                                        catch (Exception retryEx)
                                        {
                                            Logger.Log($"   ? Erreur lors de la 2�me tentative: {retryEx.Message}", Logger.LogLevel.ERROR);
                                            return null;
                                        }
                                    }
                                }
                            }
                            catch (Exception recoverEx)
                            {
                                // Si la r�cup�ration �choue, r�essayer de cr�er
                                Logger.Log($"   ?? Erreur r�cup�ration dossier: {recoverEx.Message}, nouvelle tentative de cr�ation...", Logger.LogLevel.WARNING);
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
                                        Logger.Log($"   ? Dossier cr�� (apr�s erreur r�cup�ration): {nextPath} (ID: {newFolder.Id})", Logger.LogLevel.INFO);
                                    }
                                    else
                                    {
                                        Logger.Log($"   ? �chec cr�ation (apr�s erreur r�cup�ration): {nextPath}", Logger.LogLevel.ERROR);
                                        return null;
                                    }
                                }
                                catch (Exception finalEx)
                                {
                                    Logger.Log($"   ? Erreur finale: {finalEx.Message}", Logger.LogLevel.ERROR);
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
                        Logger.Log($"   ? Impossible de continuer apr�s {nextPath}", Logger.LogLevel.ERROR);
                        return null;
                    }
                    currentPath = nextPath;
                }

                if (currentFolder == null)
                {
                    Logger.Log($"   ? currentFolder est null � la fin", Logger.LogLevel.ERROR);
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
        Logger.Log($"??? Nettoyage du dossier: {vaultFolderPath}", Logger.LogLevel.INFO);
        
        var folder = _connection!.WebServiceManager.DocumentService.GetFolderByPath(vaultFolderPath);
        
        if (folder == null)
        {
            Logger.Log($"   ?? Dossier non trouv� - rien � nettoyer", Logger.LogLevel.WARNING);
            return true;
        }
        
        // R�cup�rer TOUS les fichiers (r�cursif)
        var files = _connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, true);
        
        if (files == null || files.Length == 0)
        {
            Logger.Log($"   ? Dossier d�j� vide", Logger.LogLevel.INFO);
            return true;
        }
        
        Logger.Log($"   ?? {files.Length} fichier(s) � supprimer", Logger.LogLevel.INFO);
        
        int deleted = 0;
        int failed = 0;
        List<string> failedFiles = new List<string>();
        
        // Strat�gie: Supprimer d'abord les pi�ces (.ipt), puis les assemblages (.iam)
        // Cela �vite les erreurs de d�pendances
        var parts = files.Where(f => f.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase)).ToArray();
        var assemblies = files.Where(f => f.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase)).ToArray();
        var otherFiles = files.Where(f => !f.Name.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) && 
                                           !f.Name.EndsWith(".iam", StringComparison.OrdinalIgnoreCase)).ToArray();
        
        // Ordre de suppression: autres fichiers ? pi�ces ? assemblages
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
                        Logger.Log($"   ?? Undo checkout: {file.Name}", Logger.LogLevel.TRACE);
                        _connection.WebServiceManager.DocumentService.UndoCheckoutFile(file.MasterId, out var ticket);
                    }
                    catch (Exception undoEx)
                    {
                        Logger.Log($"   ?? Undo checkout failed pour {file.Name}: {undoEx.Message}", Logger.LogLevel.TRACE);
                    }
                }
                
                // Tenter la suppression
                _connection.WebServiceManager.DocumentService.DeleteFileFromFolder(file.MasterId, folder.Id);
                
                deleted++;
                Logger.Log($"   ? Supprim�: {file.Name}", Logger.LogLevel.TRACE);
                
                if (deleted % 10 == 0)
                {
                    Logger.Log($"   ??? {deleted}/{files.Length} supprim�s...", Logger.LogLevel.INFO);
                }
            }
            catch (ACW.VaultServiceErrorException vaultEx)
            {
                string errorMsg = $"{file.Name} (Error {vaultEx.ErrorCode})";
                
                if (vaultEx.ErrorCode == 303)
                {
                    // Error 303 = File has restrictions (lifecycle, dependencies, etc.)
                    Logger.Log($"   ?? Impossible de supprimer '{file.Name}': Error 303 - Le fichier a des restrictions (lifecycle, d�pendances, ou permissions)", Logger.LogLevel.WARNING);
                    Logger.Log($"      ?? Solution: V�rifier dans Vault Client:", Logger.LogLevel.WARNING);
                    Logger.Log($"         - Lifecycle state permet-il la suppression?", Logger.LogLevel.WARNING);
                    Logger.Log($"         - Y a-t-il des r�f�rences actives?", Logger.LogLevel.WARNING);
                    Logger.Log($"         - Permissions de suppression OK?", Logger.LogLevel.WARNING);
                }
                else if (vaultEx.ErrorCode == 1050)
                {
                    // Error 1050 = File has active references/dependencies
                    Logger.Log($"   ?? Impossible de supprimer '{file.Name}': Error 1050 - Fichier a des r�f�rences actives", Logger.LogLevel.WARNING);
                    Logger.Log($"      ?? Ce fichier est r�f�renc� par d'autres fichiers. Supprimer d'abord les fichiers qui l'utilisent.", Logger.LogLevel.WARNING);
                }
                else
                {
                    Logger.Log($"   ? Impossible de supprimer '{file.Name}': Error {vaultEx.ErrorCode} - {vaultEx.Message}", Logger.LogLevel.WARNING);
                }
                
                failedFiles.Add(errorMsg);
                failed++;
            }
            catch (Exception delEx)
            {
                Logger.Log($"   ? Impossible de supprimer '{file.Name}': {delEx.Message}", Logger.LogLevel.WARNING);
                failedFiles.Add($"{file.Name} ({delEx.Message})");
                failed++;
            }
        }
        
        Logger.Log($"", Logger.LogLevel.INFO);
        Logger.Log($"? Nettoyage termin�: {deleted} supprim�s, {failed} �chou�s", Logger.LogLevel.INFO);
        
        if (failed > 0)
        {
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"?????? FICHIERS NON SUPPRIM�S ??????", Logger.LogLevel.WARNING);
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"Les fichiers suivants n'ont pas pu �tre supprim�s:", Logger.LogLevel.WARNING);
            foreach (var failedFile in failedFiles)
            {
                Logger.Log($"   � {failedFile}", Logger.LogLevel.WARNING);
            }
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"?? RAISONS COURANTES:", Logger.LogLevel.WARNING);
            Logger.Log($"   1. Error 303 - Restrictions de lifecycle (�tat Released, etc.)", Logger.LogLevel.WARNING);
            Logger.Log($"   2. Error 1050 - Fichier r�f�renc� par d'autres fichiers", Logger.LogLevel.WARNING);
            Logger.Log($"   3. Permissions insuffisantes", Logger.LogLevel.WARNING);
            Logger.Log($"", Logger.LogLevel.WARNING);
            Logger.Log($"?? SOLUTIONS:", Logger.LogLevel.WARNING);
            Logger.Log($"   - V�rifier le lifecycle state dans Vault Client", Logger.LogLevel.WARNING);
            Logger.Log($"   - Supprimer manuellement via Vault Client (avec plus de contr�le)", Logger.LogLevel.WARNING);
            Logger.Log($"   - Changer l'�tat du fichier si n�cessaire", Logger.LogLevel.WARNING);
            Logger.Log($"   - D�tacher les r�f�rences avant suppression", Logger.LogLevel.WARNING);
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
        /// Applique les propri�t�s sur un fichier
        /// -------------------------------------------------------------------------------
        /// ?? SELON LA DOCUMENTATION SDK (topic36.html):
        /// 
        /// POUR LES UDP NON MAPP�ES (propri�t�s Vault sans mapping iProperty):
        ///   ? UpdateFileProperties fonctionne DIRECTEMENT sans checkout
        ///   ? C'est le cas simple
        /// 
        /// POUR LES UDP MAPP�ES (Content ? UDP, ex: iProperty ? Vault UDP):
        ///   ? UpdateFileProperties NE FONCTIONNE PAS car �cras� au checkin
        ///   ? Il faut modifier le fichier source avec l'API CAD
        /// 
        /// -------------------------------------------------------------------------------
        /// STRAT�GIE ICI:
        /// 1. Essayer UpdateFileProperties directement (pour UDP non mapp�es)
        /// 2. Si �a ne marche pas ou si l'UDP est mapp�e, utiliser AcquireFiles
        /// -------------------------------------------------------------------------------
        /// </summary>
        /// <summary>
        /// Applique les propri�t�s Project, Reference, Module, Revision directement via UpdateFileProperties
        /// ? SOLUTION VALID�E: Utiliser le MasterId (pas le FileId/IterationId)
        /// </summary>
        private bool SetFilePropertiesDirectly(long fileMasterId, string? projectNumber, string? reference, string? module, string? revision = null, string? checkInComment = null)
        {
            try
            {
                // -------------------------------------------------------------------------------
                // �TAPE 0: Toujours enlever Read-only AVANT toute op�ration
                // Cela �vite que le Job Processor bloque les modifications
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
                                Logger.Log($"   ?? Retrait de l'attribut read-only AVANT traitement...", Logger.LogLevel.DEBUG);
                                localFileInfo.IsReadOnly = false;
                            }
                        }
                    }
                }
                catch (Exception attrEx)
                {
                    Logger.Log($"   ?? Impossible de modifier l'attribut read-only: {attrEx.Message}", Logger.LogLevel.TRACE);
                    // Continuer quand m�me
                }
                Logger.Log($"?? Application des propri�t�s (File MasterId: {fileMasterId})...", Logger.LogLevel.DEBUG);
                
                var properties = new List<ACW.PropInstParam>();
                
                // V�rifier et ajouter Project
                if (!string.IsNullOrEmpty(projectNumber) && _propertyDefIds.ContainsKey("Project"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Project"],
                        Val = projectNumber
                    });
                    Logger.Log($"   Project = {projectNumber}", Logger.LogLevel.DEBUG);
                }
                
                // V�rifier et ajouter Reference
                if (!string.IsNullOrEmpty(reference) && _propertyDefIds.ContainsKey("Reference"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Reference"],
                        Val = reference
                    });
                    Logger.Log($"   Reference = {reference}", Logger.LogLevel.DEBUG);
                }
                
                // V�rifier et ajouter Module
                if (!string.IsNullOrEmpty(module) && _propertyDefIds.ContainsKey("Module"))
                {
                    properties.Add(new ACW.PropInstParam
                    {
                        PropDefId = _propertyDefIds["Module"],
                        Val = module
                    });
                    Logger.Log($"   Module = {module}", Logger.LogLevel.DEBUG);
                }
                
                // V�rifier et ajouter Revision
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
                    Logger.Log($"   ?? Aucune propri�t� � appliquer", Logger.LogLevel.DEBUG);
                    return true; // Pas d'erreur si aucune propri�t� � appliquer
                }
                
                    var propArray = new ACW.PropInstParamArray
                    {
                        Items = properties.ToArray()
                    };
                    
                // ================================================================================
                // SOLUTION FINALE:
                // - Inventor: GET + Checkout + Enlever Read-only + UpdateFileProperties + UndoCheckout
                //   ? Option A (UpdateFileProperties direct) supprim�e - ne fonctionne pas
                //   ? Les propri�t�s UDP persistent apr�s UndoCheckout (stock�es dans DB Vault)
                //   ? UndoCheckout enl�ve la barre dans Vault Client
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
                // STRAT�GIE POUR FICHIERS INVENTOR:
                // Essayer UpdateFileProperties DIRECTEMENT (sans checkout/checkin)
                // Si erreur 1013 (fichier doit �tre check� out), alors checkout ? update ? checkin
                // ================================================================================
                if (isInventorFile)
                {
                    // -------------------------------------------------------------------------------
                    // STRAT�GIE VALID�E: GET ? CHECKOUT ? UpdateFileProperties ? CHECKIN
                    // (La tentative directe �choue toujours avec erreur 1136 - Job Processor occup�)
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
                        // �TAPE 1: VRAI GET - T�l�charger le fichier AVANT le checkout
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] GET - T�l�chargement du fichier AVANT checkout...", Logger.LogLevel.DEBUG);
                        try
                        {
                            // Obtenir le FileIteration pour le t�l�chargement
                            var wsFile = _connection.WebServiceManager.DocumentService.GetFileById(fileInfo.Id);
                            var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, wsFile);
                            
                            // Cr�er les param�tres de t�l�chargement
                            var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                            downloadSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                            
                            // T�l�charger le fichier
                            var downloadResult = _connection.FileManager.AcquireFiles(downloadSettings);
                            
                            if (downloadResult.FileResults != null && downloadResult.FileResults.Any(r => r.LocalPath != null))
                            {
                                localFilePath = downloadResult.FileResults.First().LocalPath.FullPath;
                                Logger.Log($"   [OK] Fichier t�l�charg�: {localFilePath}", Logger.LogLevel.INFO);
                                
                                // V�rifier que le fichier existe vraiment
                                if (!System.IO.File.Exists(localFilePath))
                                {
                                    Logger.Log($"   ?? AcquireFiles a retourn� un chemin mais le fichier n'existe pas: {localFilePath}", Logger.LogLevel.WARNING);
                                    throw new System.IO.FileNotFoundException($"Fichier non trouv� apr�s t�l�chargement: {localFilePath}");
                                }
                            }
                            else
                            {
                                Logger.Log($"   ?? AcquireFiles n'a pas retourn� de chemin local", Logger.LogLevel.WARNING);
                                throw new Exception("AcquireFiles n'a pas t�l�charg� le fichier");
                            }
                        }
                        catch (Exception acquireEx)
                        {
                            Logger.Log($"   ? Erreur lors du t�l�chargement avec AcquireFiles: {acquireEx.Message}", Logger.LogLevel.ERROR);
                            throw; // Arr�ter si le GET �choue
                        }
                        
                        // -------------------------------------------------------------------------------
                        // �TAPE 2: CHECKOUT (apr�s le GET, le fichier est d�j� local)
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] Checkout (fichier d�j� t�l�charg� �: {localFilePath})...", Logger.LogLevel.DEBUG);
                        
                        // V�rifier que le fichier existe avant le checkout
                        if (!System.IO.File.Exists(localFilePath))
                        {
                            throw new System.IO.FileNotFoundException($"Fichier non trouv� avant checkout: {localFilePath}");
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
                                throw new Exception("CheckoutFile a retourn� null");
                            }
                            
                            Logger.Log($"   [OK] Checkout reussi (FileId: {checkedOutFileInv.Id}, MasterId: {checkedOutFileInv.MasterId})", Logger.LogLevel.INFO);
                        }
                        catch (ACW.VaultServiceErrorException checkoutEx)
                        {
                            Logger.Log($"   ? Erreur Vault lors du checkout: Code {checkoutEx.ErrorCode} - {checkoutEx.Message}", Logger.LogLevel.ERROR);
                            if (checkoutEx.InnerException != null)
                            {
                                Logger.Log($"   ? Inner Exception: {checkoutEx.InnerException.Message}", Logger.LogLevel.ERROR);
                            }
                            throw;
                        }
                        catch (Exception checkoutEx)
                        {
                            Logger.Log($"   ? Erreur lors du checkout: {checkoutEx.Message}", Logger.LogLevel.ERROR);
                            Logger.Log($"   ? StackTrace: {checkoutEx.StackTrace}", Logger.LogLevel.DEBUG);
                            throw;
                        }
                        
                        // -------------------------------------------------------------------------------
                        // �TAPE 3: UPDATE FILE PROPERTIES (UDP) - sur le fichier checke out
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] UpdateFileProperties (UDP)...", Logger.LogLevel.DEBUG);
                        _connection.WebServiceManager.DocumentService.UpdateFileProperties(
                            new[] { checkedOutFileInv.MasterId },
                            new[] { propArray }
                        );
                        Logger.Log($"   [OK] {properties.Count} UDP appliquees sur fichier checke out", Logger.LogLevel.INFO);
                        
                        // NOTE: La synchronisation UDP ? iProperties sera faite manuellement par l'utilisateur
                        // via clic droit ? "Synchronize Properties" dans Vault Explorer
                        // (ExplorerUtil et Inventor iProperties d�sactiv�s pour performance)
                        
                        // -------------------------------------------------------------------------------
                        // �TAPE 4: CHECKIN pour persister les propri�t�s (incluant le writeback iProperties)
                        // -------------------------------------------------------------------------------
                        Logger.Log($"   [INVENTOR] CheckIn pour persister les proprietes...", Logger.LogLevel.DEBUG);
                        
                        // Obtenir le fichier pour le CheckIn
                        var wsFileForCheckin = _connection.WebServiceManager.DocumentService.GetFileById(checkedOutFileInv.Id);
                        var vaultFileForCheckin = new VDF.Vault.Currency.Entities.FileIteration(_connection, wsFileForCheckin);
                        
                        // Commentaire d�taill� pour l'application des propri�t�s (3�me version)
                        _connection.FileManager.CheckinFile(
                            vaultFileForCheckin,
                            "MAJ propri�t�s (Project/Reference/Module) via Vault Automation Tool",
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
                        
                        Logger.Log($"   ? CheckIn reussi - Proprietes PERSISTEES!", Logger.LogLevel.INFO);
                        
                        // -------------------------------------------------------------------------------
                        // GET FINAL - T�l�charger le fichier pour enlever le rond rouge
                        // -------------------------------------------------------------------------------
                        try
                        {
                            Logger.Log($"   [INVENTOR] GET FINAL - T�l�chargement pour enlever le rond rouge...", Logger.LogLevel.DEBUG);
                            
                            // R�cup�rer le fichier apr�s CheckIn
                            var latestFilesAfterCheckin = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { checkedOutFileInv.MasterId });
                            if (latestFilesAfterCheckin != null && latestFilesAfterCheckin.Length > 0)
                            {
                                var fileAfterCheckin = latestFilesAfterCheckin[0];
                                var fileIterationAfterCheckin = new VDF.Vault.Currency.Entities.FileIteration(_connection, fileAfterCheckin);
                                
                                // Cr�er les param�tres de t�l�chargement
                                var finalDownloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                                finalDownloadSettings.AddFileToAcquire(fileIterationAfterCheckin, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                                
                                // T�l�charger le fichier
                                var finalDownloadResult = _connection.FileManager.AcquireFiles(finalDownloadSettings);
                                
                                if (finalDownloadResult.FileResults != null && finalDownloadResult.FileResults.Any(r => r.LocalPath != null))
                                {
                                    Logger.Log($"   ? GET FINAL reussi - Fichier synchronise, rond rouge enleve!", Logger.LogLevel.INFO);
                                }
                                else
                                {
                                    Logger.Log($"   ?? GET FINAL n'a pas retourne de chemin local", Logger.LogLevel.DEBUG);
                                }
                            }
                        }
                        catch (Exception getEx)
                        {
                            Logger.Log($"   ?? Erreur GET final: {getEx.Message}", Logger.LogLevel.DEBUG);
                            // Ne pas faire �chouer si le GET �choue
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
                    
                    // Retirer l'attribut read-only sur le fichier local apr�s checkout
                    try
                    {
                        if (System.IO.File.Exists(localFilePathNonInv))
                        {
                            var localFileInfo = new System.IO.FileInfo(localFilePathNonInv);
                            if (localFileInfo.IsReadOnly)
                            {
                                Logger.Log($"   ?? Retrait de l'attribut read-only sur le fichier local...", Logger.LogLevel.DEBUG);
                                localFileInfo.IsReadOnly = false;
                            }
                        }
                    }
                    catch (Exception attrEx)
                    {
                        Logger.Log($"   ?? Impossible de modifier l'attribut read-only du fichier local: {attrEx.Message}", Logger.LogLevel.TRACE);
                        // Continuer quand m�me
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
                    
                    // Commentaire d�taill� pour l'application des propri�t�s (3�me version)
                    _connection.FileManager.CheckinFile(
                        vaultFile,
                        "MAJ propri�t�s (Project / Reference / Module) via Vault Automation Tool",
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
                    // GET FINAL - T�l�charger le fichier pour enlever le rond rouge (comme pour Inventor)
                    // -------------------------------------------------------------------------------
                    try
                    {
                        Logger.Log($"   [NON-INVENTOR] GET FINAL - T�l�chargement pour enlever le rond rouge...", Logger.LogLevel.DEBUG);
                        
                        // R�cup�rer le fichier apr�s CheckIn
                        var latestFilesAfterCheckin = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { checkedOutFileNonInv.MasterId });
                        if (latestFilesAfterCheckin != null && latestFilesAfterCheckin.Length > 0)
                        {
                            var fileAfterCheckin = latestFilesAfterCheckin[0];
                            var fileIterationAfterCheckin = new VDF.Vault.Currency.Entities.FileIteration(_connection, fileAfterCheckin);
                            
                            // Cr�er les param�tres de t�l�chargement
                            var finalDownloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(_connection, false);
                            finalDownloadSettings.AddFileToAcquire(fileIterationAfterCheckin, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                            
                            // T�l�charger le fichier
                            var finalDownloadResult = _connection.FileManager.AcquireFiles(finalDownloadSettings);
                            
                            if (finalDownloadResult.FileResults != null && finalDownloadResult.FileResults.Any(r => r.LocalPath != null))
                            {
                                Logger.Log($"   ? GET FINAL reussi - Fichier synchronise, rond rouge enleve!", Logger.LogLevel.INFO);
                            }
                            else
                            {
                                Logger.Log($"   ?? GET FINAL n'a pas retourne de chemin local", Logger.LogLevel.DEBUG);
                            }
                        }
                    }
                    catch (Exception getEx)
                    {
                        Logger.Log($"   ?? Erreur GET final: {getEx.Message}", Logger.LogLevel.DEBUG);
                        // Ne pas faire �chouer si le GET �choue
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
        /// V�rifie que les propri�t�s sont bien enregistr�es dans Vault apr�s leur application
        /// </summary>
        private void VerifyFileProperties(long fileMasterId, string? projectNumber, string? reference, string? module)
                        {
                            try
                            {
                Logger.Log($"   ?? V�rification des propri�t�s enregistr�es...", Logger.LogLevel.DEBUG);
                
                // R�cup�rer le fichier pour obtenir ses propri�t�s
                var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                if (latestFiles == null || latestFiles.Length == 0)
                {
                    Logger.Log($"   ?? Fichier non trouv� pour v�rification", Logger.LogLevel.DEBUG);
                                return;
                            }
                            
                var file = latestFiles[0];
                
                // R�cup�rer les d�finitions de propri�t�s
                var propDefs = _connection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                
                // Trouver les IDs des propri�t�s Project, Reference, Module
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
                    Logger.Log($"   ?? Propri�t�s Project/Reference/Module non trouv�es pour v�rification", Logger.LogLevel.DEBUG);
                    return;
                }
                
                // R�cup�rer les valeurs des propri�t�s
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
                            Logger.Log($"   {(match ? "?" : "?")} {propName}: '{propValue}' (attendu: '{projectNumber ?? ""}')", 
                                match ? Logger.LogLevel.INFO : Logger.LogLevel.WARNING);
                        }
                        else if (propInst.PropDefId == referencePropId)
                        {
                            bool match = propValue == (reference ?? "");
                            Logger.Log($"   {(match ? "?" : "?")} {propName}: '{propValue}' (attendu: '{reference ?? ""}')", 
                                match ? Logger.LogLevel.INFO : Logger.LogLevel.WARNING);
                        }
                        else if (propInst.PropDefId == modulePropId)
                        {
                            bool match = propValue == (module ?? "");
                            Logger.Log($"   {(match ? "?" : "?")} {propName}: '{propValue}' (attendu: '{module ?? ""}')", 
                                match ? Logger.LogLevel.INFO : Logger.LogLevel.WARNING);
                        }
                    }
                }
                else
                {
                    Logger.Log($"   ?? Aucune propri�t� trouv�e pour v�rification", Logger.LogLevel.DEBUG);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   ?? Erreur lors de la v�rification des propri�t�s: {ex.Message}", Logger.LogLevel.DEBUG);
                // Ne pas faire �chouer l'op�ration si la v�rification �choue
            }
        }

        /// <summary>
        /// Applique les propri�t�s Project, Reference, Module � un fichier
        /// D�l�gue � SetFilePropertiesDirectly
        /// </summary>
        private void SetFileProperties(long fileMasterId, string? projectNumber, string? reference, string? module)
                            {
                                try
                                {
                // Simplement d�l�guer � SetFilePropertiesDirectly qui est optimis�
                bool success = SetFilePropertiesDirectly(fileMasterId, projectNumber, reference, module);
                if (!success)
                {
                    Logger.Log($"   ?? SetFileProperties: �chec pour MasterId {fileMasterId}", Logger.LogLevel.WARNING);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException($"SetFileProperties(MasterId={fileMasterId})", ex, Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Assigne une cat�gorie � un fichier via DocumentServiceExtensions.UpdateFileCategories
        /// Cette m�thode est ESSENTIELLE pour les cat�gories Engineering, Office, Standard
        /// car FileManager.AddFile ne permet PAS d'assigner une cat�gorie directement
        /// </summary>
        /// <param name="fileMasterId">ID Master du fichier</param>
        /// <param name="categoryId">ID de la cat�gorie � assigner</param>
        /// <param name="categoryName">Nom de la cat�gorie (pour logging)</param>
        /// <returns>true si la cat�gorie a �t� assign�e avec succ�s</returns>
        private bool AssignCategoryToFile(long fileMasterId, long categoryId, string categoryName)
        {
            if (!IsConnected || fileMasterId <= 0 || categoryId <= 0)
            {
                return false;
            }

            try
            {
                Logger.Log($"   ??? Assignation de la cat�gorie '{categoryName}' (ID: {categoryId}) au fichier (MasterId: {fileMasterId})...", Logger.LogLevel.INFO);

                // R�cup�rer le fichier pour obtenir son ID d'it�ration
                var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                if (latestFiles == null || latestFiles.Length == 0)
                {
                    Logger.Log($"   ?? Fichier non trouv� pour l'assignation de cat�gorie", Logger.LogLevel.WARNING);
                    return false;
                }

                var file = latestFiles[0];
                bool categoryAssigned = false;

                // -------------------------------------------------------------------------------
                // M�THODE 1: Via DocumentServiceExtensions.UpdateFileCategories (RECOMMAND�)
                // Cette m�thode est la plus fiable selon la documentation Vault SDK
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
                                        $"Cat�gorie assign�e automatiquement: {categoryName}"
                                    });
                                Logger.Log($"   ? Cat�gorie '{categoryName}' assign�e via DocumentServiceExtensions.UpdateFileCategories", Logger.LogLevel.INFO);
                                categoryAssigned = true;
                            }
                            else
                            {
                                Logger.Log($"   ?? M�thode UpdateFileCategories non trouv�e dans DocumentServiceExtensions", Logger.LogLevel.DEBUG);
                            }
                        }
                    }
                }
                catch (Exception reflectEx)
                {
                    Logger.Log($"   ?? Erreur reflection UpdateFileCategories: {reflectEx.Message}", Logger.LogLevel.DEBUG);
                    
                    // Log l'exception interne si pr�sente (souvent plus informative)
                    if (reflectEx.InnerException != null)
                    {
                        Logger.Log($"   ?? Inner Exception: {reflectEx.InnerException.Message}", Logger.LogLevel.DEBUG);
                    }
                }

                // Si la m�thode principale n'a pas fonctionn�, afficher un message
                if (!categoryAssigned)
                {
                    Logger.Log($"   ?? Impossible d'assigner la cat�gorie '{categoryName}' - API UpdateFileCategories non disponible ou erreur", Logger.LogLevel.WARNING);
                    Logger.Log($"   ?? Le fichier aura la cat�gorie par d�faut. Assignez manuellement dans Vault si n�cessaire.", Logger.LogLevel.INFO);
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
                Logger.Log("?? Chargement des Categories...", Logger.LogLevel.DEBUG);
                
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
                            Logger.Log($"   ? Category 'Base' trouv�e (ID: {category.Id})", Logger.LogLevel.INFO);
                            break;
                        }
                    }
                    
                    if (_baseCategoryId == null)
                    {
                        Logger.Log($"   ?? Category 'Base' non trouv�e", Logger.LogLevel.WARNING);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("LoadCategories", ex, Logger.LogLevel.WARNING);
            }
        }

        // NOTE: Les r�visions sont g�r�es automatiquement par Vault via les transitions d'�tat
        // (Work in Progress ? Released incr�mente la r�vision selon le sch�ma ASME configur�)

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
                        
                        // Ajouter les �tats
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
        /// Obtient le Lifecycle Definition ID selon la cat�gorie s�lectionn�e
        /// Engineering ? Flexible Release Process (For Review, Work in Progress, Released, Obsolete)
        /// Office ? Simple Release Process (Work in Progress, Released, Obsolete)
        /// Design Representation ? Design Representation Process (Released, Work in Progress, Obsolete)
        /// Standard ? Basic Release Process
        /// Base ? Aucun (pas de Lifecycle)
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
                Logger.Log($"   ?? Cat�gorie 'Base' n'a pas de Lifecycle Definition", Logger.LogLevel.INFO);
                return null;
            }
            
            // Mapping cat�gorie ? Lifecycle Definition Name
            // IMPORTANT: Ces mappings correspondent aux Lifecycle Definitions configur�es dans Vault
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
                        Logger.Log($"   ?? Mapping cat�gorie '{categoryName}' ? Lifecycle '{lifecycleName}' (ID: {lifecycleDef.Id})", Logger.LogLevel.INFO);
                        return lifecycleDef.Id;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   ?? Erreur lors de la recherche du Lifecycle Definition: {ex.Message}", Logger.LogLevel.WARNING);
                }
            }
            
            return null;
        }

        /// <summary>
        /// Obtient l'ID de l'�tat "Work in Progress" pour un Lifecycle Definition donn�
        /// </summary>
        public long? GetWorkInProgressStateId(long lifecycleDefinitionId)
        {
            try
            {
                var allLifecycleDefs = _connection!.WebServiceManager.LifeCycleService.GetAllLifeCycleDefinitions();
                var lifecycleDef = allLifecycleDefs?.FirstOrDefault(l => l.Id == lifecycleDefinitionId);
                
                if (lifecycleDef?.StateArray != null)
                {
                    // Chercher l'�tat "Work in Progress" ou l'�tat par d�faut
                    var wipState = lifecycleDef.StateArray.FirstOrDefault(s => 
                        s.DispName.IndexOf("Work", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        s.DispName.IndexOf("Progress", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        s.DispName.IndexOf("WIP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        s.IsDflt);
                    
                    if (wipState != null)
                    {
                        return wipState.Id;
                    }
                    
                    // Si aucun �tat WIP trouv�, prendre le premier �tat
                    if (lifecycleDef.StateArray.Length > 0)
                    {
                        return lifecycleDef.StateArray[0].Id;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"   ?? Erreur lors de la recherche de l'�tat WIP: {ex.Message}", Logger.LogLevel.WARNING);
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
        /// D�termine le FileClassification en fonction de la cat�gorie s�lectionn�e
        /// Selon la documentation SDK, les valeurs disponibles sont :
        /// None, DesignRepresentation, DesignDocument, DesignVisualization, DesignPresentation,
        /// DesignSubstitute, ConfigurationFactory, ConfigurationMember, ElectricalProject
        /// </summary>
        private ACW.FileClassification DetermineFileClassificationByCategory(long? categoryId, string categoryName)
        {
            if (!categoryId.HasValue || categoryId.Value <= 0)
            {
                // Aucune cat�gorie s�lectionn�e ? None
                return ACW.FileClassification.None;
            }

            // Mapper le nom de la cat�gorie au FileClassification correspondant
            if (string.IsNullOrEmpty(categoryName))
            {
                return ACW.FileClassification.None;
            }

            string categoryLower = categoryName.ToLowerInvariant().Trim();
            
            // ?? MAPPING EXPLICITE selon documentation SDK
            // Dictionnaire de mapping cat�gorie ? FileClassification
            // -------------------------------------------------------------------------------
            // IMPORTANT: Le FileClassification d�termine comment Vault traite le fichier
            // - None: Fichiers g�n�riques (PDF, Excel, Word, images, etc.)
            // - DesignRepresentation: Fichiers CAD qui repr�sentent un design (DWF, PDF g�n�r�s)
            // - DesignDocument: Documents de design (pas couramment utilis�)
            // 
            // Pour Engineering, Office, Standard: FileClassification.None est correct car
            // la cat�gorie sera assign�e S�PAR�MENT via UpdateFileCategories apr�s l'upload
            // -------------------------------------------------------------------------------
            var categoryMapping = new Dictionary<string, ACW.FileClassification>(StringComparer.OrdinalIgnoreCase)
            {
                // -------------------------------------------------------------------------------
                // Cat�gories utilisant FileClassification.None (la plupart des cas)
                // La vraie cat�gorie est assign�e APR�S l'upload via UpdateFileCategories
                // -------------------------------------------------------------------------------
                { "base", ACW.FileClassification.None },
                { "aucune", ACW.FileClassification.None },
                { "engineering", ACW.FileClassification.None },      // ? Engineering ? None
                { "office", ACW.FileClassification.None },           // ? Office ? None
                { "standard", ACW.FileClassification.None },         // ? Standard ? None
                { "cad", ACW.FileClassification.None },
                { "design", ACW.FileClassification.None },
                { "document", ACW.FileClassification.None },
                { "other", ACW.FileClassification.None },
                { "autre", ACW.FileClassification.None },
                
                // -------------------------------------------------------------------------------
                // Cat�gories avec FileClassification sp�cifique
                // Ces cat�gories n�cessitent un FileClassification particulier pour Vault
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
                Logger.Log($"   ?? Mapping explicite: Cat�gorie '{categoryName}' ? FileClassification.{mappedValue}", Logger.LogLevel.INFO);
                return mappedValue;
            }

            // 2. Chercher dans le mapping explicite (nom en minuscules)
            if (categoryMapping.TryGetValue(categoryLower, out ACW.FileClassification mappedValueLower))
            {
                Logger.Log($"   ?? Mapping explicite (lowercase): Cat�gorie '{categoryName}' ? FileClassification.{mappedValueLower}", Logger.LogLevel.INFO);
                return mappedValueLower;
            }

            // 3. Essayer de parser directement (nom nettoy�)
            try
            {
                string cleanCategoryName = categoryName.Replace(" ", "").Replace("-", "").Replace("_", "");
                if (System.Enum.TryParse<ACW.FileClassification>(cleanCategoryName, true, out ACW.FileClassification result))
                {
                    Logger.Log($"   ?? Mapping direct (nettoy�): Cat�gorie '{categoryName}' ? FileClassification.{result}", Logger.LogLevel.INFO);
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
                
                // Correspondance exacte (insensible � la casse)
                if (categoryLower == enumNameLower)
                {
                    Logger.Log($"   ?? Mapping exact: Cat�gorie '{categoryName}' ? FileClassification.{enumName}", Logger.LogLevel.INFO);
                    return enumValue;
                }
                
                // Correspondance partielle (cat�gorie contient le nom de l'enum)
                if (categoryLower.Contains(enumNameLower) || enumNameLower.Contains(categoryLower))
                {
                    Logger.Log($"   ?? Mapping partiel: Cat�gorie '{categoryName}' ? FileClassification.{enumName}", Logger.LogLevel.INFO);
                    return enumValue;
                }
            }
            
            // 5. Mapping logique par d�faut
            // Engineering, CAD, Design (g�n�rique) ? None (fichiers CAD standard selon doc)
            if (categoryLower.Contains("engineering") || categoryLower == "cad" || 
                (categoryLower.Contains("design") && !categoryLower.Contains("representation") && 
                 !categoryLower.Contains("document") && !categoryLower.Contains("visualization") &&
                 !categoryLower.Contains("presentation") && !categoryLower.Contains("substitute")))
            {
                Logger.Log($"   ?? Mapping logique: Cat�gorie '{categoryName}' ? FileClassification.None (fichier CAD standard)", Logger.LogLevel.INFO);
                return ACW.FileClassification.None;
            }
            
            // Si aucune correspondance trouv�e, logger toutes les valeurs disponibles
            Logger.Log($"   ?? Aucun FileClassification trouv� pour cat�gorie '{categoryName}'", Logger.LogLevel.WARNING);
            Logger.Log($"   ?? Valeurs FileClassification disponibles: {string.Join(", ", enumValues.Cast<ACW.FileClassification>().Select(v => v.ToString()))}", Logger.LogLevel.DEBUG);
            Logger.Log($"   ?? Utilisation FileClassification.None par d�faut", Logger.LogLevel.WARNING);
            
            // Par d�faut, utiliser None (fichiers CAD standard)
            return ACW.FileClassification.None;
        }

        public bool UploadFile(string filePath, string vaultFolderPath, 
            string? projectNumber = null, string? reference = null, string? module = null, 
            long? categoryId = null, string? categoryName = null,
            long? lifecycleDefinitionId = null, long? lifecycleStateId = null, string? revision = null, string? checkInComment = null)
        {
            if (!IsConnected)
            {
                Logger.Log("? Non connect� � Vault", Logger.LogLevel.ERROR);
                return false;
            }

            if (!File.Exists(filePath))
            {
                Logger.Log($"? Fichier introuvable: {filePath}", Logger.LogLevel.ERROR);
                return false;
            }

            // Retirer l'attribut read-only si pr�sent pour permettre la modification des propri�t�s
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                if (fileInfo.IsReadOnly)
                {
                    Logger.Log($"   ?? Retrait de l'attribut read-only sur le fichier...", Logger.LogLevel.DEBUG);
                    fileInfo.IsReadOnly = false;
                }
            }
            catch (Exception attrEx)
            {
                Logger.Log($"   ?? Impossible de modifier l'attribut read-only: {attrEx.Message}", Logger.LogLevel.WARNING);
                // Continuer quand m�me - certains syst�mes peuvent permettre l'acc�s m�me avec read-only
            }

            VDF.Vault.Currency.Entities.Folder? targetFolder = null;
            try
            {
                string fileName = Path.GetFileName(filePath);
                string fileExtension = Path.GetExtension(fileName);
                
                Logger.Log($"?? Upload: '{fileName}' ? {vaultFolderPath}", Logger.LogLevel.INFO);
                if (!string.IsNullOrEmpty(projectNumber)) Logger.Log($"   Project: {projectNumber}", Logger.LogLevel.INFO);
                if (!string.IsNullOrEmpty(reference)) Logger.Log($"   Reference: {reference}", Logger.LogLevel.INFO);
                if (!string.IsNullOrEmpty(module)) Logger.Log($"   Module: {module}", Logger.LogLevel.INFO);

                targetFolder = EnsureVaultPathExists(vaultFolderPath, projectNumber, reference, module);
                
                if (targetFolder == null || targetFolder.Id <= 0)
                {
                    Logger.Log($"? Impossible de cr�er/acc�der au dossier: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"   Dossier cible valid� (ID: {targetFolder.Id})", Logger.LogLevel.TRACE);

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
                    Logger.Log($"   ?? Le fichier '{fileName}' existe d�j� (ID: {existingFileId})", Logger.LogLevel.INFO);
                    
                    // ? NOUVELLE STRAT�GIE: Ajouter � la file d'attente au lieu d'appliquer imm�diatement
                    if (!string.IsNullOrEmpty(projectNumber) || !string.IsNullOrEmpty(reference) || !string.IsNullOrEmpty(module))
                    {
                        QueuePropertyUpdate(existingFileId, projectNumber, reference, module, fileName);
                        Logger.Log($"   ?? Propri�t�s ajout�es � la file d'attente", Logger.LogLevel.INFO);
                    }
                    
                    // ? Assigner la cat�gorie au fichier existant si sp�cifi�e
                    if (categoryId.HasValue && categoryId.Value > 0)
                    {
                        AssignCategoryToFile(existingFileId, categoryId.Value, categoryName ?? "");
                    }
                    
                    // Assigner le Lifecycle Definition si sp�cifi�
                    if (lifecycleDefinitionId.HasValue && lifecycleStateId.HasValue)
                    {
                        try
                        {
                            Logger.Log($"   ?? Assignation du Lifecycle Definition au fichier existant (ID: {lifecycleDefinitionId}, State ID: {lifecycleStateId})...", Logger.LogLevel.INFO);
                            
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
                                        Logger.Log($"   ? Lifecycle assign� avec succ�s", Logger.LogLevel.INFO);
                                    }
                                }
                                catch (Exception reflectEx)
                                {
                                    Logger.Log($"   ?? Erreur reflection UpdateFileLifeCycleStates: {reflectEx.Message}", Logger.LogLevel.WARNING);
                                }
                            }
                        }
                        catch (Exception lifecycleEx)
                        {
                            Logger.Log($"   ?? Erreur lors de l'assignation du lifecycle: {lifecycleEx.Message}", Logger.LogLevel.WARNING);
                        }
                    }
                    
                    return true; // Succ�s apr�s mise � jour des propri�t�s
                }

                // -------------------------------------------------------------------------------
                // OPTION A (RAPIDE) - Upload SANS modification iProperties pr�alable
                // 
                // ? Strat�gie optimis�e pour upload massif (1800+ fichiers):
                // 1. Upload fichier vers Vault
                // 2. UpdateFileProperties ? d�finir UDP Vault
                // 3. Job Processor sync UDP ? iProperties (apr�s check-in)
                // 
                // ? R�sultat: UDP Vault correctes + Job Processor sync vers iProperties
                // ? Temps: ~100ms par fichier
                // -------------------------------------------------------------------------------

                // -------------------------------------------------------------------------------
                // ? D�SACTIV�: NativeOlePropertyService n'�crit PAS les vraies iProperties Inventor
                // Il �crit dans le PropertySet OLE Windows standard, pas dans le format Autodesk
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
                                Logger.Log($"   [PRE-UPLOAD] ? iProperties d�finies via NativeOLE en {stopwatch.ElapsedMilliseconds}ms!", Logger.LogLevel.INFO);
                            }
                            else
                            {
                                Logger.Log($"   [PRE-UPLOAD] ?? NativeOLE: �chec modification iProperties", Logger.LogLevel.WARNING);
                            }
                        }
                    }
                    catch (Exception oleEx)
                    {
                        Logger.Log($"   [PRE-UPLOAD] ?? Exception NativeOLE: {oleEx.Message}", Logger.LogLevel.WARNING);
                        // Continuer l'upload m�me si la modification iProperties �choue
                    }
                }
                */

                // ?? SOLUTION: Utiliser FileClassification selon la cat�gorie s�lectionn�e
                // Base ? None, Design Representation ? DesignRepresentation, etc.
                ACW.FileClassification fileClass = DetermineFileClassificationByCategory(categoryId, categoryName ?? string.Empty);
                
                long newFileId = -1;
                long entityIterationId = -1;
                
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var fileInfo = new FileInfo(filePath);

                    // Upload via FileManager.AddFile avec FileClassification selon cat�gorie
                    // Le 3�me param�tre est le commentaire pour la version 1
                    // Documentation: "Text data to be associated with version 1 of the file"
                    Logger.Log($"   ?? Upload avec FileClassification.{fileClass} (Cat�gorie: {categoryName ?? "Aucune"})...", Logger.LogLevel.INFO);
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
                                
                                // Le commentaire a �t� pass� directement � AddFile (3�me param�tre)
                                // Documentation: "Text data to be associated with version 1 of the file"
                                if (!string.IsNullOrWhiteSpace(checkInComment))
                                {
                                    Logger.Log($"   ? Commentaire appliqu� � la version 1 via AddFile: '{checkInComment}'", Logger.LogLevel.INFO);
                                }
                            }
                        }
                        catch (Exception idEx)
                        {
                            Logger.Log($"   ?? Info fichier: {idEx.Message}", Logger.LogLevel.TRACE);
                        }
                    }
                }

                Logger.Log($"? Upload�: '{fileName}'", Logger.LogLevel.INFO);

                // -------------------------------------------------------------------------------
                // ? �TAPE 1: Assigner la CAT�GORIE via AssignCategoryToFile
                // Cette �tape est CRUCIALE pour Engineering, Office, Standard
                // La cat�gorie DOIT �tre assign�e APR�S l'upload car FileManager.AddFile ne le fait pas
                // -------------------------------------------------------------------------------
                if (newFileId > 0 && categoryId.HasValue && categoryId.Value > 0)
                {
                    AssignCategoryToFile(newFileId, categoryId.Value, categoryName ?? "");
                }

                // -------------------------------------------------------------------------------
                // ? �TAPE 2: Assigner le Lifecycle Definition via UpdateFileLifeCycleDefinitions
                // -------------------------------------------------------------------------------
                if (newFileId > 0 && lifecycleDefinitionId.HasValue && lifecycleStateId.HasValue)
                {
                    bool lifecycleAssigned = false;
                    
                    try
                    {
                        Logger.Log($"   ?? Assignation du Lifecycle Definition (ID: {lifecycleDefinitionId}, State ID: {lifecycleStateId})...", Logger.LogLevel.INFO);
                        
                        // R�cup�rer le fichier pour obtenir son ID d'it�ration
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
                                    // M�THODE 1: UpdateFileLifeCycleDefinitions (assigne Definition + State)
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
                                            Logger.Log($"   ? Lifecycle Definition + State assign�s via UpdateFileLifeCycleDefinitions", Logger.LogLevel.INFO);
                                            lifecycleAssigned = true;
                                        }
                                    }
                                    catch (Exception defEx)
                                    {
                                        var innerMsg = defEx.InnerException?.Message ?? defEx.Message;
                                        Logger.Log($"   ?? UpdateFileLifeCycleDefinitions �chou�: {innerMsg}", Logger.LogLevel.DEBUG);
                                        
                                        // ---------------------------------------------------------------
                                        // M�THODE 2: UpdateFileLifeCycleStates (change seulement le State)
                                        // Utile si le fichier a d�j� une Lifecycle Definition assign�e
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
                                                Logger.Log($"   ? State assign� via UpdateFileLifeCycleStates", Logger.LogLevel.INFO);
                                                lifecycleAssigned = true;
                                            }
                                        }
                                        catch (Exception stateEx)
                                        {
                                            var stateInnerMsg = stateEx.InnerException?.Message ?? stateEx.Message;
                                            Logger.Log($"   ?? UpdateFileLifeCycleStates �chou�: {stateInnerMsg}", Logger.LogLevel.DEBUG);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception lifecycleEx)
                    {
                        Logger.Log($"   ?? Erreur lors de l'assignation du lifecycle: {lifecycleEx.Message}", Logger.LogLevel.WARNING);
                    }
                    
                    if (!lifecycleAssigned)
                    {
                        Logger.Log($"   ?? Lifecycle non assign� - Le fichier sera avec le state par d�faut", Logger.LogLevel.WARNING);
                        Logger.Log($"   ?? Vous pouvez changer le state manuellement dans Vault Client", Logger.LogLevel.INFO);
                    }
                }

                // -------------------------------------------------------------------------------
                // ? �TAPE 2.5: Appliquer la r�vision via DocumentServiceExtensions.UpdateFileRevisionNumbers
                // -------------------------------------------------------------------------------
                if (newFileId > 0 && !string.IsNullOrEmpty(revision))
                {
                    try
                    {
                        Logger.Log($"   ?? Assignation de la r�vision '{revision}' au fichier (MasterId: {newFileId})...", Logger.LogLevel.INFO);
                        
                        // Obtenir le fichier pour avoir son FileId (pas MasterId)
                        var latestFiles = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { newFileId });
                        if (latestFiles != null && latestFiles.Length > 0)
                        {
                            var file = latestFiles[0];
                            
                            // Acc�der � DocumentServiceExtensions via WebServiceManager
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
                                                new[] { revision },  // R�vision
                                                "Assignation r�vision via VaultAutomationTool" 
                                            });
                                        Logger.Log($"   ? R�vision '{revision}' assign�e via UpdateFileRevisionNumbers", Logger.LogLevel.INFO);
                                        
                                        // R�cup�rer la nouvelle version apr�s mise � jour de la r�vision
                                        var updatedFiles = _connection.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { newFileId });
                                        if (updatedFiles != null && updatedFiles.Length > 0)
                                        {
                                            newFileId = updatedFiles[0].MasterId;
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log($"   ?? UpdateFileRevisionNumbers non trouv�e dans DocumentServiceExtensions", Logger.LogLevel.WARNING);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception revEx)
                    {
                        var innerMsg = revEx.InnerException?.Message ?? revEx.Message;
                        Logger.Log($"   ?? Erreur assignation r�vision: {innerMsg}", Logger.LogLevel.WARNING);
                        // Ne pas faire �chouer l'upload si la r�vision �choue
                    }
                }
                
                // ? �TAPE 3: Appliquer les propri�t�s (Project, Reference, Module, Revision)
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
                    
                    // Erreur 1000 ou 1008: Fichier existe d�j�
                    if (vaultEx.ErrorCode == 1000 || vaultEx.ErrorCode == 1008)
                    {
                        Logger.Log($"   ?? Le fichier existe d�j� (erreur {vaultEx.ErrorCode})", Logger.LogLevel.INFO);
                        
                        // Essayer de r�cup�rer le fichier existant pour appliquer les propri�t�s
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
                                        Logger.Log($"   ? Fichier existant trouv� (ID: {existingFile.MasterId}, Version: {existingFile.VerNum})", Logger.LogLevel.INFO);
                                        
                                        // Appliquer les propri�t�s au fichier existant (CheckOut ? Update Properties ? CheckIn)
                                        Logger.Log($"   ?? Application des propri�t�s au fichier existant...", Logger.LogLevel.INFO);
                                        SetFilePropertiesDirectly(existingFile.MasterId, projectNumber, reference, module, revision, checkInComment);
                                        
                                        // Assigner le Lifecycle Definition si sp�cifi�
                                        if (lifecycleDefinitionId.HasValue && lifecycleStateId.HasValue)
                                        {
                                            try
                                            {
                                                Logger.Log($"   ?? Assignation du Lifecycle Definition au fichier existant (ID: {lifecycleDefinitionId}, State ID: {lifecycleStateId})...", Logger.LogLevel.INFO);
                                                
                                                var documentServiceType = _connection.WebServiceManager.DocumentService.GetType();
                                                var updateMethod = documentServiceType.GetMethod("UpdateFileLifeCycleStates", 
                                                    new[] { typeof(long[]), typeof(long[]), typeof(string) });
                                                
                                                if (updateMethod != null)
                                                {
                                                    updateMethod.Invoke(_connection.WebServiceManager.DocumentService, 
                                                        new object[] { new[] { existingFile.Id }, new[] { lifecycleStateId.Value }, "Assignation lifecycle via upload" });
                                                    Logger.Log($"   ? Lifecycle assign� avec succ�s", Logger.LogLevel.INFO);
                                                }
                                            }
                                            catch (Exception lifecycleEx)
                                            {
                                                Logger.Log($"   ?? Erreur lors de l'assignation du lifecycle: {lifecycleEx.Message}", Logger.LogLevel.WARNING);
                                            }
                                        }
                                        
                                        return true; // Succ�s apr�s mise � jour des propri�t�s
                                    }
                                }
                            }
                            catch (Exception recoveryEx)
                            {
                                Logger.Log($"   ?? Impossible de r�cup�rer le fichier existant: {recoveryEx.Message}", Logger.LogLevel.WARNING);
                            }
                        }
                        else
                        {
                            Logger.Log($"   ?? Impossible de r�cup�rer le fichier existant: dossier non disponible", Logger.LogLevel.WARNING);
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
                Logger.Log("? Non connect� � Vault", Logger.LogLevel.ERROR);
                return (0, 0, 0);
            }

            if (!Directory.Exists(localFolderPath))
            {
                Logger.Log($"? Dossier introuvable: {localFolderPath}", Logger.LogLevel.ERROR);
                return (0, 0, 0);
            }

            Logger.Log($"?? Upload dossier complet: {localFolderPath}", Logger.LogLevel.INFO);
            Logger.Log($"   Destination Vault: {vaultBasePath}", Logger.LogLevel.INFO);
            Logger.Log($"   Project: {projectNumber}, Reference: {reference}, Module: {module}", Logger.LogLevel.INFO);

            int successCount = 0;
            int failedCount = 0;
            int skippedCount = 0;

            try
            {
                var files = Directory.GetFiles(localFolderPath, "*.*", SearchOption.AllDirectories);

                Logger.Log($"   {files.Length} fichier(s) trouv�(s)", Logger.LogLevel.INFO);

                foreach (var filePath in files)
                {
                    try
                    {
                        string? directoryName = Path.GetDirectoryName(filePath);
                        if (string.IsNullOrEmpty(directoryName)) 
                        {
                            Logger.Log($"   ?? Impossible d'obtenir le r�pertoire pour: {filePath}", Logger.LogLevel.WARNING);
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

                Logger.Log($"? Upload termin�: {successCount} succ�s, {skippedCount} ignor�s, {failedCount} erreurs", Logger.LogLevel.INFO);
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
                    Logger.Log("D�connexion de Vault...", Logger.LogLevel.INFO);
                    VDF.Vault.Library.ConnectionManager.LogOut(_connection);
                    Logger.Log("?? D�connect� du Vault", Logger.LogLevel.INFO);
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
        /// ?? Upload fichier SANS lifecycle "Released" (comme Inventor)
        /// </summary>
        public bool UploadFileWithoutLifecycle(string filePath, string vaultFolderPath, 
            string? projectNumber = null, string? reference = null, string? module = null)
        {
            if (!IsConnected)
            {
                Logger.Log("? Non connect� � Vault", Logger.LogLevel.ERROR);
                return false;
            }

            if (!File.Exists(filePath))
            {
                Logger.Log($"? Fichier introuvable: {filePath}", Logger.LogLevel.ERROR);
                return false;
            }

            try
            {
                string fileName = Path.GetFileName(filePath);
                string fileExtension = Path.GetExtension(fileName);
                
                Logger.Log($"?? Upload SANS lifecycle: '{fileName}' ? {vaultFolderPath}", Logger.LogLevel.INFO);
                
                var targetFolder = EnsureVaultPathExists(vaultFolderPath, projectNumber, reference, module);
                
                if (targetFolder == null || targetFolder.Id <= 0)
                {
                    Logger.Log($"? Impossible de cr�er/acc�der au dossier: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                // V�rifier si fichier existe d�j�
                var existingFile = FindFileInFolder(targetFolder.Id, fileName);
                if (existingFile != null)
                {
                    Logger.Log($"   ?? Le fichier '{fileName}' existe d�j�", Logger.LogLevel.INFO);
                    
                    // V�rifier le lifecycle (peut �tre bloqu�)
                    ForceFileToWorkInProgress(existingFile.MasterId);
                    
                    // Appliquer les propri�t�s (peut �chouer si lifecycle "Released")
                    return ApplyPropertiesWithoutLifecycle(existingFile.MasterId, projectNumber, reference, module);
                }

                ACW.FileClassification fileClass = DetermineFileClassification(fileExtension);
                long newFileId = -1;
                
                // Upload fichier
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var fileInfo = new FileInfo(filePath);

                    // Le 3�me param�tre est le commentaire pour la version 1
                    // Documentation: "Text data to be associated with version 1 of the file"
                    var addedFile = _connection!.FileManager.AddFile(
                        targetFolder,
                        fileName,
                        string.Empty,  // Commentaire pour la version 1 (pas de commentaire pour cette m�thode)
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
                            Logger.Log($"   ? Fichier upload� (MasterId: {newFileId})", Logger.LogLevel.INFO);
                        }
                    }
                }

                if (newFileId <= 0)
                {
                    Logger.Log($"   ? �chec upload", Logger.LogLevel.ERROR);
                    return false;
                }

                // ?? CRITIQUE: V�rifier le lifecycle (peut �tre "Released" et bloquer les modifications)
                bool lifecycleOk = ForceFileToWorkInProgress(newFileId);
                
                if (!lifecycleOk)
                {
                    Logger.Log($"   ?? Le fichier peut �tre bloqu� par le lifecycle 'Released'", Logger.LogLevel.WARNING);
                    Logger.Log($"   ?? Configurez Vault pour �viter l'assignation automatique du lifecycle", Logger.LogLevel.INFO);
                }
                
                // Appliquer les propri�t�s (peut �chouer si lifecycle "Released")
                if (!string.IsNullOrEmpty(projectNumber) || !string.IsNullOrEmpty(reference) || !string.IsNullOrEmpty(module))
                {
                    Logger.Log($"   ?? Application des propri�t�s...", Logger.LogLevel.INFO);
                    bool propsApplied = ApplyPropertiesWithoutLifecycle(newFileId, projectNumber, reference, module);
                    
                    if (!propsApplied)
                    {
                        Logger.Log($"   ?? �chec application propri�t�s - peut �tre d� au lifecycle 'Released'", Logger.LogLevel.WARNING);
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
        /// ?? Force un fichier � l'�tat "Work in Progress"
        /// 
        /// IMPORTANT: Les m�thodes de gestion du lifecycle ne sont pas disponibles dans cette version du SDK.
        /// La vraie solution est de configurer Vault pour ne pas assigner automatiquement le lifecycle "Released":
        /// 
        /// 1. Dans ADMS Console ? Behaviors ? File Lifecycle ? Default Lifecycle
        /// 2. Changer le lifecycle par d�faut pour les extensions .ipt, .iam, .idw vers un lifecycle qui commence en "Work in Progress"
        /// 3. Ou d�sactiver l'assignation automatique de lifecycle
        /// 
        /// Cette m�thode v�rifie simplement si le fichier a un lifecycle et log l'information.
        /// </summary>
        private bool ForceFileToWorkInProgress(long fileMasterId)
        {
            try
            {
                Logger.Log($"   ?? V�rification du lifecycle...", Logger.LogLevel.DEBUG);
                
                var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { fileMasterId });
                if (latestFiles == null || latestFiles.Length == 0)
                {
                    return false;
                }
                
                var file = latestFiles[0];
                
                // Pas de lifecycle = OK
                if (file.FileLfCyc == null || file.FileLfCyc.LfCycDefId <= 0)
                {
                    Logger.Log($"   ? Pas de lifecycle assign� - fichier modifiable", Logger.LogLevel.INFO);
                    return true;
                }
                
                Logger.Log($"   ?? Lifecycle d�tect� (Def ID: {file.FileLfCyc.LfCycDefId}, State ID: {file.FileLfCyc.LfCycStateId})", Logger.LogLevel.WARNING);
                Logger.Log($"   ?? SOLUTION: Configurer Vault pour ne pas assigner automatiquement le lifecycle 'Released'", Logger.LogLevel.INFO);
                Logger.Log($"      ? ADMS Console ? Behaviors ? File Lifecycle ? Default Lifecycle", Logger.LogLevel.INFO);
                Logger.Log($"      ? Changer le lifecycle par d�faut vers 'Work in Progress' pour les fichiers Inventor", Logger.LogLevel.INFO);
                
                // Le fichier peut �tre bloqu� par le lifecycle "Released"
                // Les m�thodes UpdateFileLifeCycleStates ne sont pas disponibles dans cette version du SDK
                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"   ?? Erreur v�rification lifecycle: {ex.Message}", Logger.LogLevel.WARNING);
                return false;
            }
        }

        /// <summary>
        /// ?? Retire compl�tement le lifecycle (non disponible dans cette version SDK)
        /// </summary>
        private bool RemoveLifecycleFromFile(long fileIterationId)
        {
            // Les m�thodes UpdateFileLifeCycleDefinitions ne sont pas disponibles dans cette version du SDK
            Logger.Log($"   ?? Retrait lifecycle non disponible dans cette version SDK", Logger.LogLevel.WARNING);
            Logger.Log($"   ?? Configurez Vault pour �viter l'assignation automatique du lifecycle", Logger.LogLevel.INFO);
            return false;
        }

        /// <summary>
        /// ?? Applique les propri�t�s en g�rant le lifecycle
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

                // Annuler checkout si n�cessaire
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

                // Attendre disponibilit� (10 secondes max)
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
                        "Propri�t�s XNRGY",
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
                    
                    // Re-forcer WIP apr�s CheckIn
                    System.Threading.Thread.Sleep(1000);
                    ForceFileToWorkInProgress(fileMasterId);
                    
                    Logger.Log($"   ? {properties.Count} propri�t�(s) appliqu�e(s)", Logger.LogLevel.INFO);
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
        /// Analyse un fichier dans Vault et retourne toutes ses informations (cat�gorie, FileClassification, etc.)
        /// </summary>
        public void AnalyzeFileInVault(string vaultFolderPath, string fileName)
        {
            if (!IsConnected)
            {
                Logger.Log("? Non connect� � Vault", Logger.LogLevel.ERROR);
                return;
            }

            try
            {
                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("---------------------------------------------------------------", Logger.LogLevel.INFO);
                Logger.Log($"?? ANALYSE DU FICHIER: {fileName}", Logger.LogLevel.INFO);
                Logger.Log($"   Chemin Vault: {vaultFolderPath}", Logger.LogLevel.INFO);
                Logger.Log("---------------------------------------------------------------", Logger.LogLevel.INFO);

                // Trouver le dossier
                var targetFolder = EnsureVaultPathExists(vaultFolderPath, null, null, null);
                if (targetFolder == null || targetFolder.Id <= 0)
                {
                    Logger.Log($"? Impossible de trouver le dossier: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return;
                }

                // Chercher le fichier
                var file = FindFileInFolder(targetFolder.Id, fileName);
                if (file == null)
                {
                    Logger.Log($"? Fichier '{fileName}' introuvable dans {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return;
                }

                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"?? INFORMATIONS DU FICHIER:", Logger.LogLevel.INFO);
                Logger.Log($"   ID: {file.Id}", Logger.LogLevel.INFO);
                Logger.Log($"   MasterId: {file.MasterId}", Logger.LogLevel.INFO);
                Logger.Log($"   Nom: {file.Name}", Logger.LogLevel.INFO);
                Logger.Log($"   Version: {file.VerNum}", Logger.LogLevel.INFO);
                Logger.Log($"   Chemin complet: {file.Name ?? "N/A"}", Logger.LogLevel.INFO);

                // Analyser la cat�gorie
                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"?? CAT�GORIE:", Logger.LogLevel.INFO);
                try
                {
                    bool categoryFound = false;
                    
                    // M�thode 1: Essayer GetCategoryIdsByEntityMasterIds avec string
                    try
                    {
                        var method1 = _connection!.WebServiceManager.CategoryService.GetType().GetMethod("GetCategoryIdsByEntityMasterIds", new[] { typeof(long[]), typeof(string) });
                        if (method1 != null)
                        {
                            Logger.Log($"   ?? Tentative m�thode 1: GetCategoryIdsByEntityMasterIds(long[], string)", Logger.LogLevel.DEBUG);
                            var result1 = method1.Invoke(_connection.WebServiceManager.CategoryService, new object[] { new[] { file.MasterId }, "FILE" });
                            if (result1 is long[][] categoryIdsResult1 && categoryIdsResult1.Length > 0 && categoryIdsResult1[0] != null && categoryIdsResult1[0].Length > 0)
                            {
                                long categoryId = categoryIdsResult1[0][0];
                                var category = _connection.WebServiceManager.CategoryService.GetCategoryById(categoryId);
                                if (category != null)
                                {
                                    Logger.Log($"   ? Category ID: {category.Id}", Logger.LogLevel.INFO);
                                    Logger.Log($"   ? Category Name: '{category.Name}'", Logger.LogLevel.INFO);
                                    Logger.Log($"   ? Category SystemName: '{category.SysName}'", Logger.LogLevel.INFO);
                                    categoryFound = true;
                                }
                            }
                        }
                    }
                    catch (Exception ex1)
                    {
                        Logger.Log($"   ?? M�thode 1 �chou�e: {ex1.Message}", Logger.LogLevel.DEBUG);
                    }

                    // M�thode 2: Essayer GetCategoryIdsByEntityMasterIds avec long[]
                    if (!categoryFound)
                    {
                        try
                        {
                            var method2 = _connection!.WebServiceManager.CategoryService.GetType().GetMethod("GetCategoryIdsByEntityMasterIds", new[] { typeof(long[]), typeof(long[]) });
                            if (method2 != null)
                            {
                                Logger.Log($"   ?? Tentative m�thode 2: GetCategoryIdsByEntityMasterIds(long[], long[])", Logger.LogLevel.DEBUG);
                                // EntityClassId pour FILE est g�n�ralement 1
                                var result2 = method2.Invoke(_connection.WebServiceManager.CategoryService, new object[] { new[] { file.MasterId }, new[] { 1L } });
                                if (result2 is long[][] categoryIdsResult2 && categoryIdsResult2.Length > 0 && categoryIdsResult2[0] != null && categoryIdsResult2[0].Length > 0)
                                {
                                    long categoryId = categoryIdsResult2[0][0];
                                    var category = _connection.WebServiceManager.CategoryService.GetCategoryById(categoryId);
                                    if (category != null)
                                    {
                                        Logger.Log($"   ? Category ID: {category.Id}", Logger.LogLevel.INFO);
                                        Logger.Log($"   ? Category Name: '{category.Name}'", Logger.LogLevel.INFO);
                                        Logger.Log($"   ? Category SystemName: '{category.SysName}'", Logger.LogLevel.INFO);
                                        categoryFound = true;
                                    }
                                }
                            }
                        }
                        catch (Exception ex2)
                        {
                            Logger.Log($"   ?? M�thode 2 �chou�e: {ex2.Message}", Logger.LogLevel.DEBUG);
                        }
                    }

                    // M�thode 3: Essayer via FileIteration (si disponible)
                    if (!categoryFound)
                    {
                        try
                        {
                            Logger.Log($"   ?? Tentative m�thode 3: Via FileIteration", Logger.LogLevel.DEBUG);
                            var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, file);
                            if (fileIteration != null)
                            {
                                // Essayer d'acc�der � la cat�gorie via FileIteration
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
                                            Logger.Log($"   ? Category ID: {catId}", Logger.LogLevel.INFO);
                                            Logger.Log($"   ? Category Name: '{catName}'", Logger.LogLevel.INFO);
                                            categoryFound = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex3)
                        {
                            Logger.Log($"   ?? M�thode 3 �chou�e: {ex3.Message}", Logger.LogLevel.DEBUG);
                        }
                    }

                    if (!categoryFound)
                    {
                        Logger.Log($"   ?? Aucune cat�gorie assign�e (toutes les m�thodes ont �chou�)", Logger.LogLevel.WARNING);
                        Logger.Log($"   ?? Note: Le fichier peut avoir une cat�gorie dans Vault Client mais non accessible via SDK", Logger.LogLevel.INFO);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   ?? Erreur lors de la r�cup�ration de la cat�gorie: {ex.Message}", Logger.LogLevel.WARNING);
                    Logger.Log($"   ?? Stack trace: {ex.StackTrace}", Logger.LogLevel.DEBUG);
                }

                // Obtenir le FileClassification depuis le fichier
                // Note: FileClassification n'est pas directement dans ACW.File, il faut le r�cup�rer via FileManager
                try
                {
                    var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(_connection, file);
                    if (fileIteration != null)
                    {
                        // Le FileClassification est stock� dans l'entit� FileIteration
                        var fileClass = fileIteration.FileClassification;
                        Logger.Log($"", Logger.LogLevel.INFO);
                        Logger.Log($"??? FILECLASSIFICATION:", Logger.LogLevel.INFO);
                        Logger.Log($"   FileClassification: {fileClass}", Logger.LogLevel.INFO);
                        Logger.Log($"   FileClassification (int): {(int)fileClass}", Logger.LogLevel.INFO);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   ?? Impossible de r�cup�rer FileClassification: {ex.Message}", Logger.LogLevel.WARNING);
                }

                // Analyser les propri�t�s (simplifi� - afficher seulement les propri�t�s importantes)
                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"?? PROPRI�T�S IMPORTANTES:", Logger.LogLevel.INFO);
                try
                {
                    // Utiliser GetLatestFilesByMasterIds qui retourne les propri�t�s directement dans l'objet File
                    var latestFiles = _connection!.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(new[] { file.MasterId });
                    if (latestFiles != null && latestFiles.Length > 0)
                    {
                        var latestFile = latestFiles[0];
                        // Les propri�t�s sont accessibles via les m�thodes GetPropertyValue ou directement dans l'objet File
                        Logger.Log($"   Note: Utilisez Vault Client pour voir toutes les propri�t�s", Logger.LogLevel.INFO);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"   ?? Erreur lors de la r�cup�ration des propri�t�s: {ex.Message}", Logger.LogLevel.WARNING);
                }

                // Analyser le lifecycle
                Logger.Log($"", Logger.LogLevel.INFO);
                Logger.Log($"?? LIFECYCLE:", Logger.LogLevel.INFO);
                if (file.FileLfCyc != null)
                {
                    Logger.Log($"   Lifecycle Definition ID: {file.FileLfCyc.LfCycDefId}", Logger.LogLevel.INFO);
                    Logger.Log($"   Lifecycle State ID: {file.FileLfCyc.LfCycStateId}", Logger.LogLevel.INFO);
                    
                    try
                    {
                        // R�cup�rer toutes les d�finitions de lifecycle
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
                        Logger.Log($"   ?? Erreur lors de la r�cup�ration du lifecycle: {ex.Message}", Logger.LogLevel.WARNING);
                    }
                }
                else
                {
                    Logger.Log($"   ? Aucun lifecycle assign�", Logger.LogLevel.INFO);
                }

                // Note: Les comportements (Behaviors) peuvent �tre analys�s via GetCategoryConfigurationById
                // mais n�cessitent des param�tres sp�cifiques selon la version du SDK

                Logger.Log("", Logger.LogLevel.INFO);
                Logger.Log("---------------------------------------------------------------", Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.LogException($"AnalyzeFileInVault({fileName})", ex, Logger.LogLevel.ERROR);
            }
        }

        /// <summary>
        /// T�l�charge (GET) un dossier Vault vers le working folder local
        /// Utilis� pour la mise � jour du workspace (librairies, Content Center, etc.)
        /// </summary>
        /// <param name="vaultFolderPath">Chemin Vault (ex: $/Engineering/Library/Cabinet)</param>
        /// <returns>True si succ�s</returns>
        public async Task<bool> GetFolderAsync(string vaultFolderPath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    if (_connection == null)
                    {
                        Logger.Log($"? Non connect� � Vault", Logger.LogLevel.ERROR);
                        return false;
                    }

                    Logger.Log($"?? GET: {vaultFolderPath}", Logger.LogLevel.INFO);

                    // Trouver le dossier dans Vault
                    var folder = _connection.WebServiceManager.DocumentService.GetFolderByPath(vaultFolderPath);
                    if (folder == null)
                    {
                        Logger.Log($"   ?? Dossier non trouv�: {vaultFolderPath}", Logger.LogLevel.WARNING);
                        return false;
                    }

                    // Cr�er le VaultFolder
                    var vaultFolder = new VDF.Vault.Currency.Entities.Folder(_connection, folder);

                    // Obtenir tous les fichiers du dossier (r�cursif)
                    var files = _connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(
                        folder.Id, 
                        false  // false = r�cup�rer aussi les sous-dossiers
                    );

                    if (files == null || files.Length == 0)
                    {
                        Logger.Log($"   ?? Aucun fichier dans {vaultFolderPath}", Logger.LogLevel.DEBUG);
                        return true;
                    }

                    Logger.Log($"   ?? {files.Length} fichiers � t�l�charger...", Logger.LogLevel.DEBUG);

                    // T�l�charger les fichiers
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
                            Logger.Log($"   ?? Erreur t�l�chargement {file.Name}: {fileEx.Message}", Logger.LogLevel.DEBUG);
                        }
                    }

                    Logger.Log($"   ? {successCount}/{files.Length} fichiers t�l�charg�s", Logger.LogLevel.INFO);
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Log($"? Erreur GET folder {vaultFolderPath}: {ex.Message}", Logger.LogLevel.ERROR);
                    return false;
                }
            });
        }
    }
}
