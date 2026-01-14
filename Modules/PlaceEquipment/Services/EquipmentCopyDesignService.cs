#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Inventor;
using XnrgyEngineeringAutomationTools.Models;
using XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models;
using XnrgyEngineeringAutomationTools.Services;
using IOPath = System.IO.Path;
using IOFile = System.IO.File;
using IODirectory = System.IO.Directory;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Services
{
    /// <summary>
    /// Service de Copy Design specifique aux equipements.
    /// - Utilise le IPJ specifie explicitement (pas de recherche automatique)
    /// - Copie les fichiers vers la destination sans renommer l'assemblage principal
    /// - Preserve les liens vers la Library
    /// </summary>
    public class EquipmentCopyDesignService : IDisposable
    {
        private Application? _inventorApp;
        private bool _wasAlreadyRunning;
        private bool _disposed;
        private readonly Action<string, string> _logCallback;
        private readonly Action<int, string>? _progressCallback;
        
        // Sauvegarde du projet IPJ original pour restauration
        private string? _originalProjectPath;

        // Extensions de fichiers temporaires Vault a exclure
        private static readonly string[] VaultTempExtensions = { ".v", ".v1", ".v2", ".v3", ".v4", ".v5", ".vbak", ".bak" };
        private static readonly string[] ExcludedFolders = { "_V", "OldVersions", "oldversions" };

        public EquipmentCopyDesignService(Action<string, string>? logCallback = null, Action<int, string>? progressCallback = null)
        {
            _logCallback = logCallback ?? ((msg, level) => { });
            _progressCallback = progressCallback;
        }

        #region Initialisation Inventor

        public bool Initialize()
        {
            try
            {
                Log("[>] Connexion a Inventor...", "DEBUG");

                try
                {
                    _inventorApp = (Application)Marshal.GetActiveObject("Inventor.Application");
                    _wasAlreadyRunning = true;
                    Log("[+] Connecte a instance Inventor existante", "SUCCESS");
                    return true;
                }
                catch (COMException)
                {
                    Log("[-] Aucune instance Inventor trouvee", "ERROR");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Impossible d'initialiser Inventor: {ex.Message}", "ERROR");
                return false;
            }
        }

        #endregion

        #region Gestion Projet IPJ

        /// <summary>
        /// Switch vers le projet IPJ specifie
        /// </summary>
        /// <param name="ipjPath">Chemin complet du fichier .ipj</param>
        /// <returns>True si le switch a reussi</returns>
        public bool SwitchProject(string ipjPath)
        {
            if (_inventorApp == null) return false;

            try
            {
                Log($"[>] Switch vers projet: {IOPath.GetFileName(ipjPath)}", "INFO");

                DesignProjectManager designProjectManager = _inventorApp.DesignProjectManager;

                // Sauvegarder le projet actif actuel
                try
                {
                    DesignProject activeProject = designProjectManager.ActiveDesignProject;
                    if (activeProject != null)
                    {
                        _originalProjectPath = activeProject.FullFileName;
                        Log($"[i] Projet actuel sauvegarde: {IOPath.GetFileName(_originalProjectPath)}", "DEBUG");
                    }
                }
                catch
                {
                    _originalProjectPath = null;
                }

                // Verifier que le fichier IPJ existe
                if (!IOFile.Exists(ipjPath))
                {
                    Log($"[-] Fichier IPJ introuvable: {ipjPath}", "ERROR");
                    return false;
                }

                // Fermer tous les documents avant le switch
                CloseAllDocuments();

                // Chercher ou charger le projet
                DesignProjects projectsCollection = designProjectManager.DesignProjects;
                DesignProject? targetProject = null;

                for (int i = 1; i <= projectsCollection.Count; i++)
                {
                    DesignProject proj = projectsCollection[i];
                    if (proj.FullFileName.Equals(ipjPath, StringComparison.OrdinalIgnoreCase))
                    {
                        targetProject = proj;
                        Log($"[+] Projet trouve dans la collection", "DEBUG");
                        break;
                    }
                }

                // Si pas trouve, le charger
                if (targetProject == null)
                {
                    Log($"[i] Chargement du projet: {IOPath.GetFileName(ipjPath)}", "DEBUG");
                    targetProject = projectsCollection.AddExisting(ipjPath);
                }

                // Activer le projet
                if (targetProject != null)
                {
                    targetProject.Activate();
                    Thread.Sleep(1000);
                    Log($"[+] Projet active: {IOPath.GetFileName(ipjPath)}", "SUCCESS");
                    return true;
                }
                else
                {
                    Log("[-] Impossible de charger le projet", "ERROR");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur switch projet: {ex.Message}", "ERROR");
                return false;
            }
        }

        /// <summary>
        /// Restaure le projet IPJ original
        /// </summary>
        public bool RestoreOriginalProject()
        {
            if (_inventorApp == null || string.IsNullOrEmpty(_originalProjectPath)) return false;

            try
            {
                Log($"[>] Restauration projet original: {IOPath.GetFileName(_originalProjectPath)}", "INFO");
                return SwitchProject(_originalProjectPath);
            }
            catch (Exception ex)
            {
                Log($"[!] Erreur restauration projet: {ex.Message}", "WARN");
                return false;
            }
        }

        #endregion

        #region Copy Design Equipement

        /// <summary>
        /// Execute le Copy Design pour un equipement
        /// IMPORTANT: Le switch IPJ doit etre fait AVANT d'appeler cette methode
        /// </summary>
        /// <param name="equipmentIpjPath">Chemin vers le fichier IPJ de l'equipement</param>
        /// <param name="sourceFolder">Dossier source contenant les fichiers</param>
        /// <param name="destinationFolder">Dossier de destination</param>
        /// <param name="topAssemblyFileName">Nom du fichier assemblage principal</param>
        /// <param name="filesToCopy">Liste optionnelle des fichiers a copier (si null, copie tout)</param>
        /// <param name="fileRenameMap">Dictionnaire de renommage: OriginalFileName -> NewFileName (avec suffixe _01, _02, etc.)</param>
        /// <returns>Resultat de la copie</returns>
        public async Task<EquipmentCopyResult> ExecuteEquipmentCopyDesignAsync(
            string equipmentIpjPath,
            string sourceFolder,
            string destinationFolder,
            string topAssemblyFileName,
            List<FileItem>? filesToCopy = null,
            Dictionary<string, string>? fileRenameMap = null)
        {
            // Calculer le chemin complet de l'assemblage source
            string sourceAssemblyPath = IOPath.Combine(sourceFolder, topAssemblyFileName);
            
            var result = new EquipmentCopyResult
            {
                StartTime = DateTime.Now,
                SourcePath = sourceAssemblyPath,
                DestinationPath = destinationFolder
            };

            try
            {
                if (_inventorApp == null)
                {
                    throw new Exception("Inventor non initialise");
                }

                // ══════════════════════════════════════════════════════════════════
                // CRITIQUE: Activer le mode silencieux pour éviter les dialogues
                // pendant toute l'opération de Copy Design (SaveAs, Save, etc.)
                // ══════════════════════════════════════════════════════════════════
                bool origSilentOperation = _inventorApp.SilentOperation;
                bool origUserInteractionDisabled = _inventorApp.UserInterfaceManager.UserInteractionDisabled;
                
                try
                {
                    _inventorApp.SilentOperation = true;
                    _inventorApp.UserInterfaceManager.UserInteractionDisabled = true;

                    string assemblyFileName = IOPath.GetFileName(sourceAssemblyPath);

                Log($"[>] COPY DESIGN EQUIPEMENT: {assemblyFileName}", "INFO");
                Log($"   IPJ: {IOPath.GetFileName(equipmentIpjPath)}", "DEBUG");
                Log($"   Source: {sourceFolder}", "DEBUG");
                Log($"   Destination: {destinationFolder}", "DEBUG");

                // ETAPE 0: Switch vers l'IPJ de l'equipement
                ReportProgress(2, $"Switch vers IPJ: {IOPath.GetFileName(equipmentIpjPath)}...");
                if (!SwitchProject(equipmentIpjPath))
                {
                    throw new Exception($"Impossible de switcher vers l'IPJ: {equipmentIpjPath}");
                }

                ReportProgress(5, "Creation structure dossiers...");

                // Creer la structure de dossiers
                await Task.Run(() => CreateFolderStructure(sourceFolder, destinationFolder));

                ReportProgress(10, "Ouverture fichier source...");

                // Determiner le type de fichier (assemblage ou piece)
                string sourceExt = IOPath.GetExtension(sourceAssemblyPath).ToLowerInvariant();
                bool isPartFile = sourceExt == ".ipt";
                
                // Ouvrir le fichier source (assemblage ou piece)
                Document? sourceDoc = null;
                await Task.Run(() =>
                {
                    sourceDoc = _inventorApp.Documents.Open(sourceAssemblyPath, false);
                });

                if (sourceDoc == null)
                {
                    throw new Exception($"Impossible d'ouvrir: {assemblyFileName}");
                }

                string fileTypeText = isPartFile ? "Piece" : "Assemblage";
                Log($"[+] {fileTypeText} ouvert: {assemblyFileName}", "SUCCESS");

                ReportProgress(20, "Collecte des fichiers references...");

                // Collecter tous les fichiers references
                var allReferencedDocs = new Dictionary<string, Document>(StringComparer.OrdinalIgnoreCase);
                allReferencedDocs[sourceDoc.FullFileName] = sourceDoc;
                
                // Pour les assemblages, collecter recursivement les references
                // Pour les pieces, il peut y avoir des references (iMates, etc.) mais moins nombreuses
                CollectAllReferencedDocuments(sourceDoc, allReferencedDocs);

                // Filtrer: seulement les fichiers du dossier source (pas Library)
                var filesToCopyLocal = allReferencedDocs
                    .Where(kvp => kvp.Key.StartsWith(sourceFolder, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                Log($"[i] Fichiers a copier: {filesToCopyLocal.Count}", "INFO");
                Log($"[i] Fichiers Library (liens preserves): {allReferencedDocs.Count - filesToCopyLocal.Count}", "INFO");

                ReportProgress(30, "Copie des fichiers...");

                // Trier: IPT d'abord, puis IAM, Top Assembly/Part en dernier
                var sortedFiles = filesToCopyLocal
                    .OrderBy(kvp =>
                    {
                        var doc = kvp.Value;
                        if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject) 
                        {
                            // Le fichier source (si c'est un IPT) doit etre copie en dernier
                            if (kvp.Key.Equals(sourceAssemblyPath, StringComparison.OrdinalIgnoreCase))
                                return 100;
                            return 0;
                        }
                        if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                        {
                            if (kvp.Key.Equals(sourceAssemblyPath, StringComparison.OrdinalIgnoreCase))
                                return 100;
                            return 1;
                        }
                        return 2;
                    })
                    .ToList();

                int fileIndex = 0;
                int totalFiles = sortedFiles.Count;

                foreach (var kvp in sortedFiles)
                {
                    string originalPath = kvp.Key;
                    Document doc = kvp.Value;
                    fileIndex++;

                    // Calculer le chemin relatif et destination
                    string relativePath = GetRelativePath(originalPath, sourceFolder);
                    string originalFileName = IOPath.GetFileName(originalPath);
                    
                    // Appliquer le renommage si un dictionnaire est fourni
                    string newFileName = originalFileName;
                    if (fileRenameMap != null && fileRenameMap.TryGetValue(originalFileName, out string? renamedFileName) && !string.IsNullOrEmpty(renamedFileName))
                    {
                        newFileName = renamedFileName;
                    }
                    
                    string newPath;

                    string? relativeDir = IOPath.GetDirectoryName(relativePath);
                    if (string.IsNullOrEmpty(relativeDir))
                    {
                        newPath = IOPath.Combine(destinationFolder, newFileName);
                    }
                    else
                    {
                        newPath = IOPath.Combine(destinationFolder, relativeDir, newFileName);
                    }

                    // Creer le dossier si necessaire
                    string? newDir = IOPath.GetDirectoryName(newPath);
                    if (!string.IsNullOrEmpty(newDir) && !IODirectory.Exists(newDir))
                    {
                        IODirectory.CreateDirectory(newDir);
                    }

                    try
                    {
                        // SaveAs pour copier avec mise a jour des references
                        doc.SaveAs(newPath, false);
                        
                        int progress = 30 + (int)((fileIndex / (float)totalFiles) * 50);
                        ReportProgress(progress, $"[{fileIndex}/{totalFiles}] {newFileName}");
                        Log($"  [+] [{fileIndex}/{totalFiles}] {originalFileName} -> {newFileName}", "INFO");

                        result.CopiedFiles.Add(new FileCopyInfo
                        {
                            OriginalPath = originalPath,
                            NewPath = newPath,
                            OriginalFileName = originalFileName,
                            NewFileName = newFileName
                        });
                    }
                    catch (Exception ex)
                    {
                        Log($"  [-] Erreur copie {originalFileName}: {ex.Message}", "ERROR");
                    }
                }

                ReportProgress(75, "Copie fichiers Inventor orphelins...");

                // ══════════════════════════════════════════════════════════════════
                // PHASE CRITIQUE: Copier les fichiers Inventor "orphelins" AVANT les dessins
                // Ces fichiers sont dans le dossier mais pas references directement
                // Ils sont appeles dynamiquement par iLogic selon les scenarios
                // IMPORTANT: Copier AVANT les dessins pour que les references soient resolues
                // ══════════════════════════════════════════════════════════════════
                int orphansCopied = await CopyOrphanInventorFilesAsync(sourceFolder, destinationFolder, result.CopiedFiles, fileRenameMap);
                Log($"[+] {orphansCopied} fichiers Inventor orphelins copies (appeles par iLogic)", "SUCCESS");

                ReportProgress(78, "Copie fichiers non-Inventor...");

                // Copier les fichiers non-Inventor (IPJ, images, etc.)
                await CopyNonInventorFilesAsync(sourceFolder, destinationFolder, fileRenameMap);

                ReportProgress(80, "Traitement des dessins (.idw)...");

                // ══════════════════════════════════════════════════════════════════
                // PHASE SUPPLEMENTAIRE: Traiter les fichiers .idw (dessins)
                // Les dessins ne sont PAS inclus dans AllReferencedDocuments
                // car ils referencent l'assemblage, pas l'inverse
                // IMPORTANT: Traiter APRES la copie des orphelins pour que les fichiers existent
                // ══════════════════════════════════════════════════════════════════
                
                // Creer un dictionnaire des chemins source -> destination pour les references
                var pathMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cf in result.CopiedFiles)
                {
                    if (!string.IsNullOrEmpty(cf.OriginalPath) && !string.IsNullOrEmpty(cf.NewPath))
                    {
                        pathMapping[cf.OriginalPath] = cf.NewPath;
                    }
                }

                // Collecter tous les .idw du dossier source
                var idwFiles = IODirectory.GetFiles(sourceFolder, "*.idw", System.IO.SearchOption.AllDirectories)
                    .Where(f => !ExcludedFolders.Any(ef => f.Contains($"\\{ef}\\") || f.Contains($"\\{ef}")))
                    .ToList();

                Log($"[i] Traitement de {idwFiles.Count} fichiers de dessins...", "INFO");
                int idwIndex = 0;

                foreach (var idwPath in idwFiles)
                {
                    try
                    {
                        idwIndex++;
                        string idwFileName = IOPath.GetFileName(idwPath);
                        
                        // Appliquer le renommage si un dictionnaire est fourni
                        string newIdwFileName = idwFileName;
                        if (fileRenameMap != null && fileRenameMap.TryGetValue(idwFileName, out string? renamedIdwFileName) && !string.IsNullOrEmpty(renamedIdwFileName))
                        {
                            newIdwFileName = renamedIdwFileName;
                        }
                        
                        // Calculer le nouveau chemin
                        string relativePath = GetRelativePath(idwPath, sourceFolder);
                        string? relativeDir = IOPath.GetDirectoryName(relativePath);
                        string newIdwPath;
                        
                        if (string.IsNullOrEmpty(relativeDir))
                        {
                            newIdwPath = IOPath.Combine(destinationFolder, newIdwFileName);
                        }
                        else
                        {
                            newIdwPath = IOPath.Combine(destinationFolder, relativeDir, newIdwFileName);
                        }
                        
                        // S'assurer que le dossier existe
                        string? newDir = IOPath.GetDirectoryName(newIdwPath);
                        if (!string.IsNullOrEmpty(newDir) && !IODirectory.Exists(newDir))
                        {
                            IODirectory.CreateDirectory(newDir);
                        }

                        // Ouvrir le dessin
                        var drawDoc = (DrawingDocument)_inventorApp!.Documents.Open(idwPath, false);
                        
                        // Mettre a jour les references du dessin AVANT de faire le SaveAs
                        // IMPORTANT: Passer fileRenameMap pour que les liens pointent vers les fichiers renommes (suffixe _01, _02, etc.)
                        UpdateDrawingReferences(drawDoc, sourceFolder, destinationFolder, pathMapping, fileRenameMap);
                        
                        // SaveAs vers la nouvelle destination
                        ((Document)drawDoc).SaveAs(newIdwPath, false);
                        
                        result.CopiedFiles.Add(new FileCopyInfo
                        {
                            OriginalPath = idwPath,
                            NewPath = newIdwPath,
                            OriginalFileName = idwFileName,
                            NewFileName = newIdwFileName
                        });
                        
                        Log($"  [+] [{idwIndex}/{idwFiles.Count}] {idwFileName} -> {newIdwFileName}", "SUCCESS");
                        
                        // Fermer le dessin
                        drawDoc.Close(false);
                        
                        int progress = 80 + (int)(idwIndex * 5.0 / Math.Max(idwFiles.Count, 1));
                        ReportProgress(progress, $"Dessin: {newIdwFileName}");
                    }
                    catch (Exception ex)
                    {
                        Log($"  [-] Erreur {IOPath.GetFileName(idwPath)}: {ex.Message}", "WARN");
                    }
                }

                ReportProgress(85, "Mise a jour des references assemblages...");

                // ══════════════════════════════════════════════════════════════════
                // PHASE SUPPLEMENTAIRE: Mettre a jour les references dans tous les .iam copies
                // Certains liens peuvent encore pointer vers le dossier source
                // ══════════════════════════════════════════════════════════════════
                var copiedAssemblies = result.CopiedFiles
                    .Where(f => f.NewPath.ToLowerInvariant().EndsWith(".iam"))
                    .ToList();

                Log($"[i] Verification des references dans {copiedAssemblies.Count} assemblages...", "INFO");

                foreach (var asmFile in copiedAssemblies)
                {
                    try
                    {
                        // IMPORTANT: Passer fileRenameMap pour que les liens pointent vers les fichiers renommes (suffixe _01, _02, etc.)
                        UpdateAssemblyReferences(asmFile.NewPath, sourceFolder, destinationFolder, pathMapping, fileRenameMap);
                    }
                    catch (Exception ex)
                    {
                        Log($"  [!] Erreur refs {IOPath.GetFileName(asmFile.NewPath)}: {ex.Message}", "WARN");
                    }
                }

                ReportProgress(90, "Fermeture des documents...");

                // Fermer tous les documents
                CloseAllDocuments();

                result.Success = result.CopiedFiles.Count > 0;
                result.FilesCopied = result.CopiedFiles.Count;
                result.EndTime = DateTime.Now;

                ReportProgress(100, $"[+] Copy Design termine: {result.FilesCopied} fichiers");
                Log($"[+] COPY DESIGN TERMINE: {result.FilesCopied} fichiers copies", "SUCCESS");

                return result;
                }
                finally
                {
                    // CRITIQUE: Restaurer les modes silencieux
                    _inventorApp.SilentOperation = origSilentOperation;
                    _inventorApp.UserInterfaceManager.UserInteractionDisabled = origUserInteractionDisabled;
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Log($"[-] ERREUR Copy Design: {ex.Message}", "ERROR");
                ReportProgress(0, $"[-] Erreur: {ex.Message}");
                return result;
            }
        }

        #endregion

        #region Helpers

        private void CollectAllReferencedDocuments(Document doc, Dictionary<string, Document> collected)
        {
            try
            {
                DocumentsEnumerator referencedDocs = doc.ReferencedDocuments;
                foreach (Document refDoc in referencedDocs)
                {
                    string refPath = refDoc.FullFileName;
                    if (!collected.ContainsKey(refPath))
                    {
                        collected[refPath] = refDoc;
                        CollectAllReferencedDocuments(refDoc, collected);
                    }
                }
            }
            catch { }
        }

        private void CreateFolderStructure(string sourceFolder, string destFolder)
        {
            if (!IODirectory.Exists(destFolder))
            {
                IODirectory.CreateDirectory(destFolder);
            }

            foreach (var dir in IODirectory.GetDirectories(sourceFolder, "*", System.IO.SearchOption.AllDirectories))
            {
                string dirName = IOPath.GetFileName(dir);
                if (ExcludedFolders.Contains(dirName, StringComparer.OrdinalIgnoreCase))
                    continue;

                string relativePath = GetRelativePath(dir, sourceFolder);
                string destDir = IOPath.Combine(destFolder, relativePath);
                if (!IODirectory.Exists(destDir))
                {
                    IODirectory.CreateDirectory(destDir);
                    Log($"  [i] Dossier cree: {relativePath}", "DEBUG");
                }
            }
        }

        private async Task CopyNonInventorFilesAsync(string sourceFolder, string destFolder, Dictionary<string, string>? fileRenameMap)
        {
            await Task.Run(() =>
            {
                string[] inventorExtensions = { ".ipt", ".iam", ".idw", ".ipn", ".dwg" };
                
                foreach (var file in IODirectory.GetFiles(sourceFolder, "*.*", System.IO.SearchOption.AllDirectories))
                {
                    string ext = IOPath.GetExtension(file).ToLowerInvariant();
                    string fileName = IOPath.GetFileName(file);
                    string dirName = IOPath.GetFileName(IOPath.GetDirectoryName(file) ?? "");

                    // Exclure fichiers Inventor (deja copies)
                    if (inventorExtensions.Contains(ext)) continue;

                    // Exclure fichiers temporaires Vault
                    if (VaultTempExtensions.Any(e => ext.StartsWith(e))) continue;
                    if (fileName.StartsWith("~$") || fileName.StartsWith("._")) continue;

                    // Exclure dossiers _V
                    if (ExcludedFolders.Contains(dirName, StringComparer.OrdinalIgnoreCase)) continue;

                    // Appliquer le renommage si un dictionnaire est fourni
                    string newFileName = fileName;
                    if (fileRenameMap != null && fileRenameMap.TryGetValue(fileName, out string? renamedFileName) && !string.IsNullOrEmpty(renamedFileName))
                    {
                        newFileName = renamedFileName;
                    }

                    string relativePath = GetRelativePath(file, sourceFolder);
                    string? relativeDir = IOPath.GetDirectoryName(relativePath);
                    string destPath;
                    
                    if (string.IsNullOrEmpty(relativeDir))
                    {
                        destPath = IOPath.Combine(destFolder, newFileName);
                    }
                    else
                    {
                        destPath = IOPath.Combine(destFolder, relativeDir, newFileName);
                    }
                    
                    string? destDir = IOPath.GetDirectoryName(destPath);

                    if (!string.IsNullOrEmpty(destDir) && !IODirectory.Exists(destDir))
                    {
                        IODirectory.CreateDirectory(destDir);
                    }

                    try
                    {
                        IOFile.Copy(file, destPath, true);
                        Log($"  [+] Copie: {fileName} -> {newFileName}", "DEBUG");
                    }
                    catch (Exception ex)
                    {
                        Log($"  [!] Erreur copie {fileName}: {ex.Message}", "WARN");
                    }
                }
            });
        }

        /// <summary>
        /// Copie les fichiers Inventor "orphelins" - fichiers .ipt/.iam/.ipn qui sont dans le dossier source
        /// mais qui n'ont pas ete copies par le Copy Design car non references directement.
        /// Ces fichiers sont typiquement appeles dynamiquement par les regles iLogic.
        /// IMPORTANT: Copie physique simple sans ouvrir dans Inventor (evite les erreurs de references)
        /// </summary>
        private async Task<int> CopyOrphanInventorFilesAsync(string sourceFolder, string destFolder, List<FileCopyInfo> alreadyCopied, Dictionary<string, string>? fileRenameMap)
        {
            int copiedCount = 0;
            
            await Task.Run(() =>
            {
                string[] inventorExtensions = { ".ipt", ".iam", ".ipn" };
                
                // Creer un HashSet des fichiers deja copies (pour verification rapide)
                var copiedPaths = new HashSet<string>(
                    alreadyCopied.Select(f => f.OriginalPath ?? ""),
                    StringComparer.OrdinalIgnoreCase
                );
                
                // Parcourir tous les fichiers Inventor du dossier source
                foreach (var file in IODirectory.GetFiles(sourceFolder, "*.*", System.IO.SearchOption.AllDirectories))
                {
                    string ext = IOPath.GetExtension(file).ToLowerInvariant();
                    string fileName = IOPath.GetFileName(file);
                    string dirName = IOPath.GetFileName(IOPath.GetDirectoryName(file) ?? "");

                    // Seulement les fichiers Inventor (pas .idw qui sont traites separement)
                    if (!inventorExtensions.Contains(ext)) continue;

                    // Exclure fichiers temporaires Vault
                    if (VaultTempExtensions.Any(e => ext.StartsWith(e))) continue;
                    if (fileName.StartsWith("~$") || fileName.StartsWith("._")) continue;

                    // Exclure dossiers _V et OldVersions
                    if (ExcludedFolders.Contains(dirName, StringComparer.OrdinalIgnoreCase)) continue;
                    if (file.Contains("\\_V\\") || file.Contains("\\OldVersions\\")) continue;

                    // Verifier si ce fichier a deja ete copie par le Copy Design
                    if (copiedPaths.Contains(file)) continue;

                    // Appliquer le renommage si un dictionnaire est fourni
                    string newFileName = fileName;
                    if (fileRenameMap != null && fileRenameMap.TryGetValue(fileName, out string? renamedFileName) && !string.IsNullOrEmpty(renamedFileName))
                    {
                        newFileName = renamedFileName;
                    }

                    // Calculer le chemin de destination
                    string relativePath = GetRelativePath(file, sourceFolder);
                    string? relativeDir = IOPath.GetDirectoryName(relativePath);
                    string destPath;
                    
                    if (string.IsNullOrEmpty(relativeDir))
                    {
                        destPath = IOPath.Combine(destFolder, newFileName);
                    }
                    else
                    {
                        destPath = IOPath.Combine(destFolder, relativeDir, newFileName);
                    }
                    
                    string? destDir = IOPath.GetDirectoryName(destPath);

                    // Verifier si le fichier existe deja a destination
                    if (IOFile.Exists(destPath)) continue;

                    // Creer le dossier si necessaire
                    if (!string.IsNullOrEmpty(destDir) && !IODirectory.Exists(destDir))
                    {
                        IODirectory.CreateDirectory(destDir);
                    }

                    try
                    {
                        // Copie physique simple (pas de SaveAs, pas d'ouverture Inventor)
                        IOFile.Copy(file, destPath, true);
                        copiedCount++;
                        
                        // IMPORTANT: Ajouter a la liste pour que le pathMapping soit complet
                        lock (alreadyCopied)
                        {
                            alreadyCopied.Add(new FileCopyInfo
                            {
                                OriginalPath = file,
                                NewPath = destPath,
                                OriginalFileName = fileName,
                                NewFileName = newFileName
                            });
                        }
                        
                        // Log tous les 50 fichiers pour ne pas spammer
                        if (copiedCount <= 10 || copiedCount % 50 == 0)
                        {
                            Log($"  [+] Orphelin: {fileName} -> {newFileName}", "DEBUG");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"  [!] Erreur copie orphelin {fileName}: {ex.Message}", "WARN");
                    }
                }
                
                if (copiedCount > 10)
                {
                    Log($"  ... et {copiedCount - 10} autres fichiers orphelins", "DEBUG");
                }
            });
            
            return copiedCount;
        }

        /// <summary>
        /// Met a jour les references d'un dessin pour pointer vers les fichiers copies
        /// Utilise ReferencedFileDescriptor.PutLogicalFileNameUsingFull() pour changer les liens
        /// IMPORTANT: Prend en compte le renommage des fichiers (suffixe _01, _02, etc.)
        /// </summary>
        private void UpdateDrawingReferences(
            DrawingDocument drawDoc, 
            string sourceRoot, 
            string destRoot, 
            Dictionary<string, string> pathMapping,
            Dictionary<string, string>? fileRenameMap = null)
        {
            try
            {
                Document doc = (Document)drawDoc;
                
                // Parcourir tous les ReferencedFileDescriptors
                foreach (ReferencedFileDescriptor refDesc in doc.ReferencedFileDescriptors)
                {
                    try
                    {
                        string refPath = refDesc.FullFileName;
                        
                        // Verifier si on a un mapping direct (chemin complet source -> destination)
                        if (pathMapping.TryGetValue(refPath, out string? newPath))
                        {
                            if (IOFile.Exists(newPath))
                            {
                                refDesc.PutLogicalFileNameUsingFull(newPath);
                                Log($"    -> {IOPath.GetFileName(refPath)} => {IOPath.GetFileName(newPath)}", "DEBUG");
                            }
                            continue;
                        }

                        // Si c'est un fichier du dossier source, calculer le nouveau chemin
                        if (refPath.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        {
                            string fileName = IOPath.GetFileName(refPath);
                            
                            // IMPORTANT: Appliquer le renommage si disponible (suffixe _01, _02, etc.)
                            string newFileName = fileName;
                            if (fileRenameMap != null && fileRenameMap.TryGetValue(fileName, out string? renamedFile) && !string.IsNullOrEmpty(renamedFile))
                            {
                                newFileName = renamedFile;
                            }
                            
                            string relativePath = GetRelativePath(refPath, sourceRoot);
                            string? relativeDir = IOPath.GetDirectoryName(relativePath);
                            
                            string newRefPath;
                            if (string.IsNullOrEmpty(relativeDir))
                            {
                                newRefPath = IOPath.Combine(destRoot, newFileName);
                            }
                            else
                            {
                                newRefPath = IOPath.Combine(destRoot, relativeDir, newFileName);
                            }

                            if (IOFile.Exists(newRefPath))
                            {
                                refDesc.PutLogicalFileNameUsingFull(newRefPath);
                                Log($"    -> {fileName} => {newFileName}", "DEBUG");
                            }
                            else
                            {
                                Log($"    [!] Fichier non trouve: {newRefPath}", "DEBUG");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"    [!] Reference {IOPath.GetFileName(refDesc.FullFileName)}: {ex.Message}", "DEBUG");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  [!] Mise a jour references dessin: {ex.Message}", "DEBUG");
            }
        }

        /// <summary>
        /// Met a jour les references d'un assemblage pour pointer vers les fichiers copies
        /// </summary>
        private void UpdateAssemblyReferences(
            string assemblyPath, 
            string sourceRoot, 
            string destRoot, 
            Dictionary<string, string> pathMapping,
            Dictionary<string, string>? fileRenameMap = null)
        {
            if (_inventorApp == null) return;
            
            // Activer le mode silencieux pour éviter les dialogues
            bool origSilent = _inventorApp.SilentOperation;
            bool origUserDisabled = _inventorApp.UserInterfaceManager.UserInteractionDisabled;
            
            try
            {
                _inventorApp.SilentOperation = true;
                _inventorApp.UserInterfaceManager.UserInteractionDisabled = true;
                
                // Ouvrir l'assemblage
                var doc = _inventorApp.Documents.Open(assemblyPath, false);
                bool modified = false;
                
                // Parcourir tous les ReferencedFileDescriptors
                foreach (ReferencedFileDescriptor refDesc in doc.ReferencedFileDescriptors)
                {
                    try
                    {
                        string refPath = refDesc.FullFileName;
                        string? newPath = null;
                        
                        // METHODE 1: Chercher dans le mapping exact
                        if (pathMapping.TryGetValue(refPath, out string? exactPath) && !string.IsNullOrEmpty(exactPath))
                        {
                            newPath = exactPath;
                        }
                        // METHODE 2: Reference pointe vers sourceRoot
                        else if (refPath.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        {
                            string fileName = IOPath.GetFileName(refPath);
                            
                            // IMPORTANT: Appliquer le renommage si disponible (suffixe _01, _02, etc.)
                            string newFileName = fileName;
                            if (fileRenameMap != null && fileRenameMap.TryGetValue(fileName, out string? renamedFile) && !string.IsNullOrEmpty(renamedFile))
                            {
                                newFileName = renamedFile;
                            }
                            
                            string relativePath = GetRelativePath(refPath, sourceRoot);
                            string? relativeDir = IOPath.GetDirectoryName(relativePath);
                            
                            if (string.IsNullOrEmpty(relativeDir))
                            {
                                newPath = IOPath.Combine(destRoot, newFileName);
                            }
                            else
                            {
                                newPath = IOPath.Combine(destRoot, relativeDir, newFileName);
                            }
                            
                            // Si le chemin calcule n'existe pas, rechercher le fichier
                            if (!IOFile.Exists(newPath))
                            {
                                var foundFiles = IODirectory.GetFiles(destRoot, newFileName, System.IO.SearchOption.AllDirectories);
                                if (foundFiles.Length > 0)
                                {
                                    newPath = foundFiles[0];
                                }
                                else
                                {
                                    newPath = null;
                                }
                            }
                        }
                        
                        // Appliquer la modification si on a trouve un nouveau chemin
                        if (!string.IsNullOrEmpty(newPath) && IOFile.Exists(newPath))
                        {
                            refDesc.PutLogicalFileNameUsingFull(newPath);
                            modified = true;
                            Log($"    -> {IOPath.GetFileName(refPath)} => {IOPath.GetFileName(newPath)}", "DEBUG");
                        }
                    }
                    catch { }
                }
                
                // Sauvegarder si modifie
                if (modified)
                {
                    doc.Save();
                    Log($"  [+] References mises a jour: {IOPath.GetFileName(assemblyPath)}", "SUCCESS");
                }
                
                doc.Close(false);
            }
            catch (Exception ex)
            {
                Log($"  [!] Erreur update refs {IOPath.GetFileName(assemblyPath)}: {ex.Message}", "WARN");
            }
            finally
            {
                // Restaurer les modes silencieux
                _inventorApp!.SilentOperation = origSilent;
                _inventorApp.UserInterfaceManager.UserInteractionDisabled = origUserDisabled;
            }
        }

        private string GetRelativePath(string fullPath, string basePath)
        {
            if (!basePath.EndsWith(IOPath.DirectorySeparatorChar.ToString()))
                basePath += IOPath.DirectorySeparatorChar;

            if (fullPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(basePath.Length);

            return fullPath;
        }

        private void CloseAllDocuments()
        {
            if (_inventorApp == null) return;

            try
            {
                int docCount = _inventorApp.Documents.Count;
                if (docCount > 0)
                {
                    Log($"[i] Fermeture de {docCount} document(s)...", "DEBUG");
                }

                while (_inventorApp.Documents.Count > 0)
                {
                    try
                    {
                        _inventorApp.Documents[1].Close(true);
                    }
                    catch { break; }
                }
            }
            catch { }
        }

        /// <summary>
        /// Version publique de CloseAllDocuments pour usage externe
        /// </summary>
        public void CloseAllDocumentsPublic()
        {
            CloseAllDocuments();
        }

        public bool OpenDocument(string filePath)
        {
            if (_inventorApp == null) return false;

            try
            {
                Log($"[>] Ouverture: {IOPath.GetFileName(filePath)}", "DEBUG");
                _inventorApp.Documents.Open(filePath, true);
                Log($"[+] Document ouvert: {IOPath.GetFileName(filePath)}", "SUCCESS");
                return true;
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur ouverture: {ex.Message}", "ERROR");
                return false;
            }
        }

        public bool PlaceComponent(string componentPath)
        {
            if (_inventorApp == null) return false;

            // Activer le mode silencieux pour éviter les dialogues
            bool origSilent = _inventorApp.SilentOperation;
            bool origUserDisabled = _inventorApp.UserInterfaceManager.UserInteractionDisabled;
            
            try
            {
                _inventorApp.SilentOperation = true;
                _inventorApp.UserInterfaceManager.UserInteractionDisabled = true;
                
                Document? activeDoc = _inventorApp.ActiveDocument;
                if (activeDoc == null || activeDoc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    Log("[-] Aucun assemblage actif", "ERROR");
                    return false;
                }

                AssemblyDocument asmDoc = (AssemblyDocument)activeDoc;
                AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;

                Log($"[>] Insertion composant: {IOPath.GetFileName(componentPath)}", "INFO");
                
                Matrix matrix = _inventorApp.TransientGeometry.CreateMatrix();
                ComponentOccurrence occ = asmDef.Occurrences.Add(componentPath, matrix);

                Log($"[+] Composant insere: {occ.Name}", "SUCCESS");
                return true;
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur insertion: {ex.Message}", "ERROR");
                return false;
            }
            finally
            {
                // Restaurer les modes silencieux
                _inventorApp.SilentOperation = origSilent;
                _inventorApp.UserInterfaceManager.UserInteractionDisabled = origUserDisabled;
            }
        }

        public void SaveAll()
        {
            if (_inventorApp == null) return;

            // Activer le mode silencieux pour éviter les dialogues
            bool origSilent = _inventorApp.SilentOperation;
            bool origUserDisabled = _inventorApp.UserInterfaceManager.UserInteractionDisabled;
            
            try
            {
                _inventorApp.SilentOperation = true;
                _inventorApp.UserInterfaceManager.UserInteractionDisabled = true;
                
                foreach (Document doc in _inventorApp.Documents)
                {
                    try
                    {
                        if (!doc.IsModifiable) continue;
                        doc.Save();
                    }
                    catch { }
                }
                Log("[+] Tous les documents sauvegardes", "SUCCESS");
            }
            catch (Exception ex)
            {
                Log($"[!] Erreur sauvegarde: {ex.Message}", "WARN");
            }
            finally
            {
                // Restaurer les modes silencieux
                _inventorApp.SilentOperation = origSilent;
                _inventorApp.UserInterfaceManager.UserInteractionDisabled = origUserDisabled;
            }
        }

        /// <summary>
        /// Prepare la vue pour le dessinateur:
        /// - Cache tous les plans, axes, points et esquisses
        /// - Met la vue en ISO
        /// - Fait un Zoom All
        /// </summary>
        public void PrepareViewForDesigner()
        {
            if (_inventorApp == null) return;

            try
            {
                Document? activeDoc = _inventorApp.ActiveDocument;
                if (activeDoc == null) return;

                // 1. Cacher tous les elements de reference via ObjectVisibility
                try
                {
                    if (activeDoc is AssemblyDocument asmDoc)
                    {
                        asmDoc.ObjectVisibility.UserWorkPlanes = false;
                        asmDoc.ObjectVisibility.UserWorkAxes = false;
                        asmDoc.ObjectVisibility.UserWorkPoints = false;
                        asmDoc.ObjectVisibility.OriginWorkPlanes = false;
                        asmDoc.ObjectVisibility.OriginWorkAxes = false;
                        asmDoc.ObjectVisibility.OriginWorkPoints = false;
                        asmDoc.ObjectVisibility.Sketches = false;
                        asmDoc.ObjectVisibility.Sketches3D = false;
                        Log("[+] Plans et references caches", "DEBUG");
                    }
                }
                catch (Exception ex)
                {
                    Log($"[!] ObjectVisibility: {ex.Message}", "DEBUG");
                }

                // 2. Vue Isometrique
                try
                {
                    CommandManager cmdManager = _inventorApp.CommandManager;
                    ControlDefinitions controlDefs = cmdManager.ControlDefinitions;
                    ControlDefinition cmdIso = controlDefs["AppIsometricViewCmd"];
                    cmdIso.Execute();
                    Log("[+] Vue ISO appliquee", "DEBUG");
                }
                catch
                {
                    // Fallback via Camera
                    try
                    {
                        View? activeView = _inventorApp.ActiveView;
                        if (activeView != null)
                        {
                            Camera camera = activeView.Camera;
                            camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
                            camera.Apply();
                            Log("[+] Vue ISO appliquee (fallback)", "DEBUG");
                        }
                    }
                    catch { }
                }

                // 3. Zoom All
                try
                {
                    CommandManager cmdManager = _inventorApp.CommandManager;
                    ControlDefinitions controlDefs = cmdManager.ControlDefinitions;
                    ControlDefinition cmdZoom = controlDefs["AppZoomAllCmd"];
                    cmdZoom.Execute();
                    Log("[+] Zoom All", "DEBUG");
                }
                catch
                {
                    // Fallback via View.Fit()
                    try
                    {
                        View? activeView = _inventorApp.ActiveView;
                        activeView?.Fit();
                        Log("[+] Zoom All (Fit fallback)", "DEBUG");
                    }
                    catch { }
                }

                Log("[+] Vue preparee pour le dessinateur", "SUCCESS");
            }
            catch (Exception ex)
            {
                Log($"[!] Erreur preparation vue: {ex.Message}", "WARN");
            }
        }

        private void Log(string message, string level)
        {
            _logCallback?.Invoke($"[EquipCopyDesign] {message}", level);
            
            // Log aussi dans le fichier
            switch (level.ToUpper())
            {
                case "ERROR":
                    Logger.Error($"[EquipCopyDesign] {message}");
                    break;
                case "WARN":
                case "WARNING":
                    Logger.Warning($"[EquipCopyDesign] {message}");
                    break;
                case "DEBUG":
                    Logger.Debug($"[EquipCopyDesign] {message}");
                    break;
                default:
                    Logger.Info($"[EquipCopyDesign] {message}");
                    break;
            }
        }

        private void ReportProgress(int percent, string status)
        {
            _progressCallback?.Invoke(percent, status);
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (_disposed) return;

            try
            {
                if (_inventorApp != null && !_wasAlreadyRunning)
                {
                    try
                    {
                        _inventorApp.Quit();
                    }
                    catch { }
                }
                _inventorApp = null;
            }
            catch { }

            _disposed = true;
        }

        #endregion
    }

    /// <summary>
    /// Resultat du Copy Design equipement
    /// </summary>
    public class EquipmentCopyResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public string SourcePath { get; set; } = "";
        public string DestinationPath { get; set; } = "";
        public int FilesCopied { get; set; }
        public List<FileCopyInfo> CopiedFiles { get; set; } = new();
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
    }

    /// <summary>
    /// Info sur un fichier copie
    /// </summary>
    public class FileCopyInfo
    {
        public string OriginalPath { get; set; } = "";
        public string NewPath { get; set; } = "";
        public string OriginalFileName { get; set; } = "";
        public string NewFileName { get; set; } = "";
    }
}
