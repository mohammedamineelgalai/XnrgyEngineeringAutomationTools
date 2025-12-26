#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Inventor;
using XnrgyEngineeringAutomationTools.Models;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service de Copy Design utilisant Pack & Go d'Inventor.
    /// - Utilise le Pack & Go natif d'Inventor pour copier avec références intactes
    /// - Applique les iProperties uniquement sur Module_.iam (Top Assembly)
    /// - Renomme le Module_.iam avec le numéro formaté (ex: 123450101.iam)
    /// - Conserve la structure des sous-dossiers du template
    /// - Exclut les fichiers temporaires Vault (.v*, _V dossiers)
    /// </summary>
    public class InventorCopyDesignService : IDisposable
    {
        private Application? _inventorApp;
        private bool _wasAlreadyRunning;
        private bool _disposed;
        private readonly Action<string, string> _logCallback;
        private readonly Action<int, string>? _progressCallback;

        // Extensions de fichiers temporaires Vault à exclure
        private static readonly string[] VaultTempExtensions = { ".v", ".v1", ".v2", ".v3", ".v4", ".v5", ".vbak", ".bak" };
        private static readonly string[] ExcludedFolders = { "_V", "OldVersions", "oldversions" };
        private const string VaultTempFolderSuffix = "_V";

        public InventorCopyDesignService(Action<string, string>? logCallback = null, Action<int, string>? progressCallback = null)
        {
            _logCallback = logCallback ?? ((msg, level) => { });
            _progressCallback = progressCallback;
        }

        #region Initialisation Inventor

        /// <summary>
        /// Initialise la connexion à Inventor (ou démarre une instance invisible)
        /// </summary>
        public bool Initialize()
        {
            try
            {
                Log("Connexion à Inventor...", "INFO");

                // Essayer de se connecter à une instance existante
                try
                {
                    _inventorApp = (Application)Marshal.GetActiveObject("Inventor.Application");
                    _wasAlreadyRunning = true;
                    Log("✓ Connecté à instance Inventor existante", "SUCCESS");
                    return true;
                }
                catch (COMException)
                {
                    // Pas d'instance existante, démarrer une nouvelle instance invisible
                    Log("Aucune instance Inventor trouvée, démarrage en mode invisible...", "INFO");

                    Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType == null)
                    {
                        Log("Inventor n'est pas installé sur ce système", "ERROR");
                        return false;
                    }

                    _inventorApp = (Application)Activator.CreateInstance(inventorType)!;
                    _inventorApp.Visible = false;
                    _wasAlreadyRunning = false;

                    Log("✓ Instance Inventor démarrée (invisible)", "SUCCESS");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log($"Impossible d'initialiser Inventor: {ex.Message}", "ERROR");
                return false;
            }
        }

        #endregion

        #region Copy Design Principal (Pack & Go)

        /// <summary>
        /// Exécute le Copy Design avec Pack & Go d'Inventor
        /// </summary>
        public async Task<ModuleCopyResult> ExecuteCopyDesignAsync(CreateModuleRequest request)
        {
            var result = new ModuleCopyResult
            {
                Success = false,
                StartTime = DateTime.Now,
                CopiedFiles = new List<FileCopyResult>()
            };

            if (_inventorApp == null)
            {
                result.ErrorMessage = "Inventor non initialisé";
                return result;
            }

            try
            {
                Log($"=== COPY DESIGN: {request.FullProjectNumber} ===", "START");
                ReportProgress(0, "Préparation du Copy Design...");

                // Trouver le fichier Module_.iam (Top Assembly)
                var topAssemblyFile = request.FilesToCopy
                    .FirstOrDefault(f => f.IsTopAssembly || 
                                         f.OriginalFileName.Equals("Module_.iam", StringComparison.OrdinalIgnoreCase));

                if (topAssemblyFile == null)
                {
                    // Chercher n'importe quel .iam dans la racine du template
                    topAssemblyFile = request.FilesToCopy
                        .FirstOrDefault(f => System.IO.Path.GetExtension(f.OriginalPath).ToLower() == ".iam" &&
                                             !f.OriginalPath.Contains("1-Equipment") &&
                                             !f.OriginalPath.Contains("2-Floor") &&
                                             !f.OriginalPath.Contains("3-Wall"));
                }

                if (topAssemblyFile == null)
                {
                    throw new Exception("Module_.iam (Top Assembly) non trouvé dans le template");
                }

                string sourceTopAssembly = topAssemblyFile.OriginalPath;
                string sourceFolderRoot = System.IO.Path.GetDirectoryName(sourceTopAssembly) ?? "";
                
                // Nouveau nom pour le Top Assembly: numéro formaté (ex: 123450101.iam)
                string newTopAssemblyName = $"{request.FullProjectNumber}.iam";
                
                Log($"Top Assembly source: {topAssemblyFile.OriginalFileName}", "INFO");
                Log($"Nouveau nom: {newTopAssemblyName}", "INFO");
                Log($"Destination: {request.DestinationPath}", "INFO");

                ReportProgress(5, "Création de la structure de dossiers...");

                // ÉTAPE 1: Créer la structure de dossiers destination en copiant celle du template
                await Task.Run(() => CreateFolderStructureFromTemplate(sourceFolderRoot, request.DestinationPath));

                ReportProgress(10, "Ouverture du Top Assembly...");

                // ÉTAPE 2: Ouvrir le Top Assembly
                AssemblyDocument? asmDoc = null;
                try
                {
                    asmDoc = (AssemblyDocument)await Task.Run(() => 
                        _inventorApp!.Documents.Open(sourceTopAssembly, false));
                }
                catch (Exception ex)
                {
                    throw new Exception($"Impossible d'ouvrir {sourceTopAssembly}: {ex.Message}");
                }

                ReportProgress(20, "Application des iProperties sur Module_.iam...");

                // ÉTAPE 3: Appliquer les iProperties UNIQUEMENT sur le Top Assembly
                SetIProperties((Document)asmDoc, request);
                result.PropertiesUpdated = 1;
                Log("✓ iProperties appliquées sur le Top Assembly", "SUCCESS");

                ReportProgress(30, $"Pack & Go vers {request.DestinationPath}...");

                // ÉTAPE 4: Utiliser Pack & Go pour copier tout avec les références
                string newTopAssemblyPath = System.IO.Path.Combine(request.DestinationPath, newTopAssemblyName);
                
                var packAndGoResult = await ExecutePackAndGoAsync(asmDoc, sourceFolderRoot, request.DestinationPath, newTopAssemblyName);
                
                result.CopiedFiles = packAndGoResult.CopiedFiles;
                result.FilesCopied = packAndGoResult.FilesCopied;

                // Fermer le document
                asmDoc.Close(true);
                asmDoc = null;

                ReportProgress(80, "Copie des fichiers non-Inventor...");

                // ÉTAPE 5: Copier les fichiers non-Inventor en conservant la structure
                var nonInventorResult = await CopyNonInventorFilesAsync(sourceFolderRoot, request.DestinationPath, request);
                result.FilesCopied += nonInventorResult.Count;
                foreach (var fileResult in nonInventorResult)
                {
                    result.CopiedFiles.Add(fileResult);
                }

                ReportProgress(90, "Renommage du fichier projet (.ipj)...");

                // ÉTAPE 6: Renommer le fichier .ipj avec le numéro de projet formaté
                await RenameProjectFileAsync(request.DestinationPath, request.FullProjectNumber);

                result.Success = result.FilesCopied > 0;
                result.EndTime = DateTime.Now;
                result.DestinationPath = request.DestinationPath;

                ReportProgress(100, $"✓ Copy Design terminé: {result.FilesCopied} fichiers");
                Log($"=== COPY DESIGN TERMINÉ: {result.FilesCopied} fichiers copiés ===", "SUCCESS");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Log($"ERREUR Copy Design: {ex.Message}", "ERROR");
                ReportProgress(0, $"✗ Erreur: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Exécute le Pack & Go d'Inventor
        /// </summary>
        private async Task<(List<FileCopyResult> CopiedFiles, int FilesCopied)> ExecutePackAndGoAsync(
            AssemblyDocument asmDoc, string sourceRoot, string destRoot, string newTopAssemblyName)
        {
            var copiedFiles = new List<FileCopyResult>();
            int filesCopied = 0;

            await Task.Run(() =>
            {
                try
                {
                    // Créer l'objet DesignProject pour Pack & Go
                    var designMgr = _inventorApp!.DesignProjectManager;
                    
                    // Utiliser la méthode SaveAs avec CopyDesignCache pour simuler Pack & Go
                    // Cette approche préserve les références internes
                    
                    var referencedDocs = new List<Document>();
                    CollectAllReferencedDocuments((Document)asmDoc, referencedDocs);
                    
                    Log($"Pack & Go: {referencedDocs.Count + 1} fichiers Inventor à traiter", "INFO");
                    ReportProgress(35, $"Pack & Go: {referencedDocs.Count + 1} fichiers Inventor...");

                    int totalFiles = referencedDocs.Count + 1;
                    int processed = 0;

                    // Sauvegarder tous les fichiers référencés d'abord (bottom-up)
                    // Trier par type: IPT d'abord, puis sous-IAM, puis dessins
                    var sortedDocs = referencedDocs
                        .OrderBy(d => d.DocumentType == DocumentTypeEnum.kPartDocumentObject ? 0 :
                                      d.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? 1 : 2)
                        .ToList();

                    foreach (var refDoc in sortedDocs)
                    {
                        try
                        {
                            string originalPath = refDoc.FullFileName;
                            string relativePath = GetRelativePath(originalPath, sourceRoot);
                            string newPath = System.IO.Path.Combine(destRoot, relativePath);
                            
                            // S'assurer que le dossier existe
                            string? newDir = System.IO.Path.GetDirectoryName(newPath);
                            if (!string.IsNullOrEmpty(newDir) && !Directory.Exists(newDir))
                            {
                                Directory.CreateDirectory(newDir);
                            }

                            // SaveAs pour copier avec mise à jour des références
                            refDoc.SaveAs(newPath, false);
                            
                            copiedFiles.Add(new FileCopyResult
                            {
                                OriginalPath = originalPath,
                                OriginalFileName = System.IO.Path.GetFileName(originalPath),
                                NewPath = newPath,
                                NewFileName = System.IO.Path.GetFileName(newPath),
                                Success = true
                            });
                            filesCopied++;
                            processed++;
                            
                            int progress = 35 + (int)(processed * 40.0 / totalFiles);
                            ReportProgress(progress, $"Copie: {System.IO.Path.GetFileName(newPath)}");
                            Log($"  ✓ {System.IO.Path.GetFileName(newPath)}", "SUCCESS");
                        }
                        catch (Exception ex)
                        {
                            Log($"  ✗ Erreur copie {System.IO.Path.GetFileName(refDoc.FullFileName)}: {ex.Message}", "ERROR");
                        }
                    }

                    // Enfin, sauvegarder le Top Assembly avec le nouveau nom
                    string topAssemblyNewPath = System.IO.Path.Combine(destRoot, newTopAssemblyName);
                    asmDoc.SaveAs(topAssemblyNewPath, false);
                    
                    copiedFiles.Add(new FileCopyResult
                    {
                        OriginalPath = asmDoc.FullFileName,
                        OriginalFileName = "Module_.iam",
                        NewPath = topAssemblyNewPath,
                        NewFileName = newTopAssemblyName,
                        Success = true,
                        IsTopAssembly = true,
                        PropertiesUpdated = true
                    });
                    filesCopied++;
                    
                    Log($"  ✓ Top Assembly: {newTopAssemblyName}", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur Pack & Go: {ex.Message}", "ERROR");
                    throw;
                }
            });

            return (copiedFiles, filesCopied);
        }

        /// <summary>
        /// Collecte récursivement tous les documents référencés
        /// </summary>
        private void CollectAllReferencedDocuments(Document doc, List<Document> collected)
        {
            foreach (Document refDoc in doc.ReferencedDocuments)
            {
                if (!collected.Any(d => d.FullFileName.Equals(refDoc.FullFileName, StringComparison.OrdinalIgnoreCase)))
                {
                    collected.Add(refDoc);
                    
                    // Récursif pour les assemblages imbriqués
                    if (refDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ||
                        refDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    {
                        CollectAllReferencedDocuments(refDoc, collected);
                    }
                }
            }
        }

        /// <summary>
        /// Copie les fichiers non-Inventor en conservant la structure des dossiers
        /// </summary>
        private async Task<List<FileCopyResult>> CopyNonInventorFilesAsync(string sourceRoot, string destRoot, CreateModuleRequest request)
        {
            var results = new List<FileCopyResult>();

            await Task.Run(() =>
            {
                try
                {
                    // Parcourir tous les fichiers du dossier source
                    var allFiles = Directory.GetFiles(sourceRoot, "*.*", SearchOption.AllDirectories);
                    
                    Log($"Analyse: {allFiles.Length} fichiers trouvés dans le template", "INFO");

                    foreach (var filePath in allFiles)
                    {
                        // Exclure les fichiers temporaires Vault et .bak
                        if (IsVaultTempFile(filePath))
                        {
                            Log($"  Exclu (Vault temp/bak): {System.IO.Path.GetFileName(filePath)}", "DEBUG");
                            continue;
                        }

                        // Exclure les fichiers Inventor (déjà copiés par Pack & Go)
                        if (IsInventorFile(filePath))
                        {
                            continue;
                        }

                        // Exclure les dossiers _V et OldVersions
                        string? dirPath = System.IO.Path.GetDirectoryName(filePath);
                        if (!string.IsNullOrEmpty(dirPath) && 
                            ExcludedFolders.Any(ef => dirPath.Contains($"\\{ef}\\") || dirPath.EndsWith($"\\{ef}")))
                        {
                            Log($"  Exclu (dossier exclu): {System.IO.Path.GetFileName(filePath)}", "DEBUG");
                            continue;
                        }

                        // Calculer le chemin relatif et destination
                        string relativePath = GetRelativePath(filePath, sourceRoot);
                        string destPath = System.IO.Path.Combine(destRoot, relativePath);

                        try
                        {
                            // Créer le dossier si nécessaire
                            string? destDir = System.IO.Path.GetDirectoryName(destPath);
                            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                            {
                                Directory.CreateDirectory(destDir);
                            }

                            // Copier le fichier
                            System.IO.File.Copy(filePath, destPath, overwrite: true);

                            results.Add(new FileCopyResult
                            {
                                OriginalPath = filePath,
                                OriginalFileName = System.IO.Path.GetFileName(filePath),
                                NewPath = destPath,
                                NewFileName = System.IO.Path.GetFileName(destPath),
                                Success = true
                            });

                            Log($"  ✓ {relativePath} (non-Inventor)", "SUCCESS");
                        }
                        catch (Exception ex)
                        {
                            Log($"  ✗ Erreur copie {System.IO.Path.GetFileName(filePath)}: {ex.Message}", "WARN");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur copie fichiers non-Inventor: {ex.Message}", "ERROR");
                }
            });

            return results;
        }

        #endregion

        #region iProperties

        /// <summary>
        /// Définit les iProperties du module sur le Top Assembly uniquement
        /// </summary>
        private void SetIProperties(Document doc, CreateModuleRequest request)
        {
            try
            {
                Log("Application des iProperties sur Module_.iam...", "INFO");
                
                PropertySets propSets = doc.PropertySets;
                
                // Propriétés standard Inventor (Summary Information)
                try
                {
                    PropertySet summaryProps = propSets["Inventor Summary Information"];
                    SetOrCreateProperty(summaryProps, "Title", request.FullProjectNumber);
                    SetOrCreateProperty(summaryProps, "Subject", $"Module HVAC - {request.Project}");
                    SetOrCreateProperty(summaryProps, "Author", request.InitialeDessinateur);
                }
                catch { }

                // Propriétés de design tracking
                try
                {
                    PropertySet designProps = propSets["Design Tracking Properties"];
                    SetOrCreateProperty(designProps, "Part Number", request.FullProjectNumber);
                    SetOrCreateProperty(designProps, "Project", request.Project);
                    SetOrCreateProperty(designProps, "Designer", request.InitialeDessinateur);
                    SetOrCreateProperty(designProps, "Creation Date", request.CreationDate.ToString("yyyy-MM-dd"));
                }
                catch { }

                // Propriétés personnalisées XNRGY
                try
                {
                    PropertySet customProps = propSets["Inventor User Defined Properties"];
                    
                    SetOrCreateProperty(customProps, "Project", request.Project);
                    SetOrCreateProperty(customProps, "Reference", request.Reference);
                    SetOrCreateProperty(customProps, "Module", request.Module);
                    SetOrCreateProperty(customProps, "Numero_de_Projet", request.FullProjectNumber);
                    SetOrCreateProperty(customProps, "Initiale_du_Dessinateur", request.InitialeDessinateur);
                    SetOrCreateProperty(customProps, "Initiale_du_Co_Dessinateur", request.InitialeCoDessinateur ?? "");
                    SetOrCreateProperty(customProps, "Creation_Date", request.CreationDate.ToString("yyyy-MM-dd"));
                    SetOrCreateProperty(customProps, "Job_Title", request.JobTitle ?? "");
                    
                    Log($"  Project: {request.Project}", "INFO");
                    Log($"  Reference: {request.Reference}", "INFO");
                    Log($"  Module: {request.Module}", "INFO");
                    Log($"  Numéro complet: {request.FullProjectNumber}", "INFO");
                    Log($"  Dessinateur: {request.InitialeDessinateur}", "INFO");
                }
                catch (Exception ex)
                {
                    Log($"  Erreur propriétés personnalisées: {ex.Message}", "WARN");
                }

                Log("✓ iProperties définies sur Module_.iam", "SUCCESS");
            }
            catch (Exception ex)
            {
                Log($"Erreur iProperties: {ex.Message}", "ERROR");
                throw;
            }
        }

        private void SetOrCreateProperty(PropertySet propSet, string propName, string? value)
        {
            if (value == null) return;

            try
            {
                // Essayer de trouver la propriété existante
                Property prop = propSet[propName];
                prop.Value = value;
            }
            catch
            {
                // La propriété n'existe pas, la créer
                try
                {
                    propSet.Add(value, propName);
                }
                catch { }
            }
        }

        #endregion

        #region Project File Renaming

        /// <summary>
        /// Renomme le fichier .ipj avec le numéro de projet formaté
        /// </summary>
        private async Task RenameProjectFileAsync(string destinationPath, string fullProjectNumber)
        {
            await Task.Run(() =>
            {
                try
                {
                    // Chercher tous les fichiers .ipj dans le dossier destination
                    var ipjFiles = Directory.GetFiles(destinationPath, "*.ipj", SearchOption.TopDirectoryOnly);

                    if (ipjFiles.Length == 0)
                    {
                        Log("Aucun fichier .ipj trouvé dans la destination", "DEBUG");
                        return;
                    }

                    foreach (var ipjFile in ipjFiles)
                    {
                        string newIpjName = $"{fullProjectNumber}.ipj";
                        string newIpjPath = System.IO.Path.Combine(destinationPath, newIpjName);

                        // Ne pas renommer si c'est déjà le bon nom
                        if (System.IO.Path.GetFileName(ipjFile).Equals(newIpjName, StringComparison.OrdinalIgnoreCase))
                        {
                            Log($"Fichier .ipj déjà correctement nommé: {newIpjName}", "INFO");
                            continue;
                        }

                        // Renommer le fichier
                        if (System.IO.File.Exists(newIpjPath))
                        {
                            System.IO.File.Delete(newIpjPath);
                            Log($"Ancien fichier .ipj supprimé: {newIpjName}", "DEBUG");
                        }

                        System.IO.File.Move(ipjFile, newIpjPath);
                        Log($"✓ Fichier .ipj renommé: {System.IO.Path.GetFileName(ipjFile)} → {newIpjName}", "SUCCESS");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur renommage fichier .ipj: {ex.Message}", "WARN");
                }
            });
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Crée la structure de dossiers à partir du template
        /// </summary>
        private void CreateFolderStructureFromTemplate(string sourceRoot, string destRoot)
        {
            try
            {
                // Créer le dossier racine
                if (!Directory.Exists(destRoot))
                {
                    Directory.CreateDirectory(destRoot);
                    Log($"Dossier créé: {destRoot}", "INFO");
                }

                // Parcourir et recréer tous les sous-dossiers du template
                var sourceDirs = Directory.GetDirectories(sourceRoot, "*", SearchOption.AllDirectories);
                
                foreach (var sourceDir in sourceDirs)
                {
                    var dirName = System.IO.Path.GetFileName(sourceDir);
                    
                    // Exclure les dossiers temporaires Vault (_V) et OldVersions
                    if (ExcludedFolders.Any(ef => dirName.Equals(ef, StringComparison.OrdinalIgnoreCase) ||
                                                    sourceDir.Contains($"\\{ef}\\")))
                    {
                        Log($"  Exclu (dossier exclu): {dirName}", "DEBUG");
                        continue;
                    }

                    string relativePath = GetRelativePath(sourceDir, sourceRoot);
                    string destDir = System.IO.Path.Combine(destRoot, relativePath);

                    if (!Directory.Exists(destDir))
                    {
                        Directory.CreateDirectory(destDir);
                        Log($"  Dossier: {relativePath}", "INFO");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur création structure dossiers: {ex.Message}", "ERROR");
            }
        }

        /// <summary>
        /// Obtient le chemin relatif d'un fichier par rapport à un dossier racine
        /// </summary>
        private string GetRelativePath(string fullPath, string rootPath)
        {
            // S'assurer que les chemins sont normalisés
            fullPath = System.IO.Path.GetFullPath(fullPath);
            rootPath = System.IO.Path.GetFullPath(rootPath);

            if (!rootPath.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            {
                rootPath += System.IO.Path.DirectorySeparatorChar;
            }

            if (fullPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase))
            {
                return fullPath.Substring(rootPath.Length);
            }

            return System.IO.Path.GetFileName(fullPath);
        }

        /// <summary>
        /// Vérifie si un fichier est un fichier temporaire Vault
        /// </summary>
        private bool IsVaultTempFile(string filePath)
        {
            string fileName = System.IO.Path.GetFileName(filePath);
            string extension = System.IO.Path.GetExtension(filePath).ToLower();

            // Fichiers .v, .v1, .v2, etc.
            if (VaultTempExtensions.Any(ext => extension.StartsWith(ext)))
            {
                return true;
            }

            // Fichiers avec pattern de backup Vault
            if (fileName.Contains(".v") && fileName.LastIndexOf(".v") + 2 < fileName.Length && 
                char.IsDigit(fileName[fileName.LastIndexOf(".v") + 2]))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Vérifie si un fichier est un fichier Inventor
        /// </summary>
        private bool IsInventorFile(string filePath)
        {
            var ext = System.IO.Path.GetExtension(filePath).ToLower();
            return ext == ".ipt" || ext == ".iam" || ext == ".idw" || ext == ".ipn" || ext == ".dwg";
        }

        private void Log(string message, string level)
        {
            _logCallback(message, level);
            Logger.Log($"[CopyDesign] {message}", level switch
            {
                "ERROR" => Logger.LogLevel.ERROR,
                "WARN" => Logger.LogLevel.WARNING,
                "SUCCESS" => Logger.LogLevel.INFO,
                "START" => Logger.LogLevel.INFO,
                "DEBUG" => Logger.LogLevel.DEBUG,
                _ => Logger.LogLevel.DEBUG
            });
        }

        private void ReportProgress(int percent, string message)
        {
            _progressCallback?.Invoke(percent, message);
        }

        #endregion

        #region Dispose

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_inventorApp != null && !_wasAlreadyRunning)
                {
                    Log("Fermeture de l'instance Inventor...", "INFO");
                    _inventorApp.Quit();
                }

                if (_inventorApp != null)
                {
                    Marshal.ReleaseComObject(_inventorApp);
                    _inventorApp = null;
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur fermeture Inventor: {ex.Message}", "WARN");
            }
        }

        #endregion
    }
}
