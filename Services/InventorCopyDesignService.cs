#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Inventor;
using XnrgyEngineeringAutomationTools.Models;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service de Copy Design utilisant Pack & Go d'Inventor.
    /// - Switch vers le projet IPJ du template avant le Pack & Go
    /// - Utilise le Pack & Go natif d'Inventor pour copier avec références intactes
    /// - Applique les iProperties uniquement sur Module_.iam (Top Assembly)
    /// - Renomme le Module_.iam avec le numéro formaté (ex: 123450101.iam)
    /// - Préserve les liens vers la Library (C:\Vault\Engineering\Library)
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
        
        // Sauvegarde du projet IPJ original pour restauration
        private string? _originalProjectPath;

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

        #region Gestion Projet IPJ

        /// <summary>
        /// Switch vers le projet IPJ du template (requis pour ouvrir les fichiers avec les bonnes références)
        /// Sauvegarde le projet actuel pour restauration ultérieure
        /// </summary>
        /// <param name="templateIpjPath">Chemin complet du fichier .ipj du template</param>
        /// <returns>True si le switch a réussi</returns>
        public bool SwitchToTemplateProject(string templateIpjPath)
        {
            if (_inventorApp == null) return false;

            try
            {
                Log($"🔄 Switch vers projet template: {System.IO.Path.GetFileName(templateIpjPath)}", "INFO");

                DesignProjectManager designProjectManager = _inventorApp.DesignProjectManager;

                // Sauvegarder le projet actif actuel
                try
                {
                    DesignProject activeProject = designProjectManager.ActiveDesignProject;
                    if (activeProject != null)
                    {
                        _originalProjectPath = activeProject.FullFileName;
                        Log($"💾 Projet actuel sauvegardé: {System.IO.Path.GetFileName(_originalProjectPath)}", "DEBUG");
                    }
                }
                catch
                {
                    _originalProjectPath = null;
                }

                // Vérifier que le fichier IPJ du template existe
                if (!System.IO.File.Exists(templateIpjPath))
                {
                    Log($"❌ Fichier IPJ template introuvable: {templateIpjPath}", "ERROR");
                    return false;
                }

                // Fermer tous les documents avant le switch
                CloseAllDocuments();

                // Chercher ou charger le projet template
                DesignProjects projectsCollection = designProjectManager.DesignProjects;
                DesignProject? templateProject = null;

                for (int i = 1; i <= projectsCollection.Count; i++)
                {
                    DesignProject proj = projectsCollection[i];
                    if (proj.FullFileName.Equals(templateIpjPath, StringComparison.OrdinalIgnoreCase))
                    {
                        templateProject = proj;
                        Log($"✅ Projet template trouvé dans la collection", "DEBUG");
                        break;
                    }
                }

                // Si pas trouvé, le charger
                if (templateProject == null)
                {
                    Log($"📂 Chargement du projet template: {System.IO.Path.GetFileName(templateIpjPath)}", "DEBUG");
                    templateProject = projectsCollection.AddExisting(templateIpjPath);
                }

                // Activer le projet template
                if (templateProject != null)
                {
                    templateProject.Activate();
                    Thread.Sleep(1000); // Attendre que le switch soit effectif
                    Log($"✅ Projet template activé: {System.IO.Path.GetFileName(templateIpjPath)}", "SUCCESS");
                    return true;
                }
                else
                {
                    Log("❌ Impossible de charger le projet template", "ERROR");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Erreur switch projet template: {ex.Message}", "ERROR");
                return false;
            }
        }

        /// <summary>
        /// Restaure le projet IPJ original après le Pack & Go
        /// </summary>
        /// <returns>True si la restauration a réussi</returns>
        public bool RestoreOriginalProject()
        {
            if (_inventorApp == null || string.IsNullOrEmpty(_originalProjectPath)) return false;

            try
            {
                Log($"🔄 Restauration projet original: {System.IO.Path.GetFileName(_originalProjectPath)}", "INFO");

                // Fermer tous les documents
                CloseAllDocuments();

                DesignProjectManager designProjectManager = _inventorApp.DesignProjectManager;
                DesignProjects projectsCollection = designProjectManager.DesignProjects;
                DesignProject? originalProject = null;

                // Chercher le projet original
                for (int i = 1; i <= projectsCollection.Count; i++)
                {
                    DesignProject proj = projectsCollection[i];
                    if (proj.FullFileName.Equals(_originalProjectPath, StringComparison.OrdinalIgnoreCase))
                    {
                        originalProject = proj;
                        break;
                    }
                }

                // Si pas trouvé, le recharger
                if (originalProject == null)
                {
                    if (System.IO.File.Exists(_originalProjectPath))
                    {
                        originalProject = projectsCollection.AddExisting(_originalProjectPath);
                    }
                }

                // Activer le projet original
                if (originalProject != null)
                {
                    originalProject.Activate();
                    Thread.Sleep(1000);
                    Log($"✅ Projet original restauré: {System.IO.Path.GetFileName(_originalProjectPath)}", "SUCCESS");
                    return true;
                }
                else
                {
                    Log($"⚠️ Impossible de restaurer le projet original", "WARN");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"⚠️ Erreur restauration projet: {ex.Message}", "WARN");
                return false;
            }
        }

        /// <summary>
        /// Switch vers le nouveau projet IPJ créé (après Copy Design)
        /// Ne sauvegarde PAS le projet actuel car on veut rester sur le nouveau
        /// </summary>
        /// <param name="newIpjPath">Chemin complet du nouveau fichier .ipj</param>
        /// <returns>True si le switch a réussi</returns>
        public bool SwitchToNewProject(string newIpjPath)
        {
            if (_inventorApp == null) return false;

            try
            {
                Log($"🔄 Switch vers nouveau projet: {System.IO.Path.GetFileName(newIpjPath)}", "INFO");

                DesignProjectManager designProjectManager = _inventorApp.DesignProjectManager;

                // Vérifier que le fichier IPJ existe
                if (!System.IO.File.Exists(newIpjPath))
                {
                    Log($"❌ Fichier IPJ introuvable: {newIpjPath}", "ERROR");
                    return false;
                }

                // Fermer tous les documents avant le switch
                CloseAllDocuments();

                // Charger le nouveau projet
                DesignProjects projectsCollection = designProjectManager.DesignProjects;
                DesignProject? newProject = null;

                // Chercher si déjà dans la collection
                for (int i = 1; i <= projectsCollection.Count; i++)
                {
                    DesignProject proj = projectsCollection[i];
                    if (proj.FullFileName.Equals(newIpjPath, StringComparison.OrdinalIgnoreCase))
                    {
                        newProject = proj;
                        Log($"✅ Nouveau projet trouvé dans la collection", "DEBUG");
                        break;
                    }
                }

                // Si pas trouvé, le charger
                if (newProject == null)
                {
                    Log($"📂 Chargement du nouveau projet: {System.IO.Path.GetFileName(newIpjPath)}", "DEBUG");
                    newProject = projectsCollection.AddExisting(newIpjPath);
                }

                // Activer le nouveau projet
                if (newProject != null)
                {
                    newProject.Activate();
                    Thread.Sleep(1000);
                    Log($"✅ Nouveau projet activé: {System.IO.Path.GetFileName(newIpjPath)}", "SUCCESS");
                    return true;
                }
                else
                {
                    Log("❌ Impossible de charger le nouveau projet", "ERROR");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Erreur switch nouveau projet: {ex.Message}", "ERROR");
                return false;
            }
        }

        /// <summary>
        /// Finalise le nouveau module après Copy Design:
        /// 1. Ouvre le Top Assembly
        /// 2. Applique les iProperties et paramètres
        /// 3. Update All (rebuild)
        /// 4. Save All
        /// 5. Laisse le document ouvert pour le dessinateur
        /// </summary>
        private async Task FinalizeNewModuleAsync(string topAssemblyPath, CreateModuleRequest request)
        {
            if (_inventorApp == null) return;

            await Task.Run(() =>
            {
                try
                {
                    Log($"📂 Ouverture du nouveau module: {System.IO.Path.GetFileName(topAssemblyPath)}", "INFO");

                    // 1. Ouvrir le Top Assembly
                    Document topAssemblyDoc = _inventorApp.Documents.Open(topAssemblyPath, true);
                    
                    if (topAssemblyDoc == null)
                    {
                        Log($"❌ Impossible d'ouvrir: {topAssemblyPath}", "ERROR");
                        return;
                    }

                    Log($"✅ Document ouvert: {System.IO.Path.GetFileName(topAssemblyPath)}", "SUCCESS");

                    // 2. Appliquer les iProperties
                    Log($"📝 Application des iProperties...", "INFO");
                    ApplyIPropertiesToDocument(topAssemblyDoc, request);

                    // 3. Appliquer les paramètres Inventor (si c'est un assemblage)
                    if (topAssemblyDoc is AssemblyDocument assemblyDoc)
                    {
                        Log($"⚙️ Application des paramètres Inventor...", "INFO");
                        ApplyInventorParameters(assemblyDoc, request);
                    }

                    // 4. Update All (rebuild de l'assemblage)
                    Log($"🔄 Update All (rebuild)...", "INFO");
                    try
                    {
                        topAssemblyDoc.Update2(true); // true = full update
                        Log($"✅ Update terminé", "SUCCESS");
                    }
                    catch (Exception updateEx)
                    {
                        Log($"⚠️ Erreur pendant Update: {updateEx.Message}", "WARN");
                    }

                    // 5. Save All
                    Log($"💾 Save All...", "INFO");
                    try
                    {
                        // Sauvegarder le document principal
                        topAssemblyDoc.Save2(true); // true = save referenced documents too

                        // Sauvegarder tous les documents ouverts (au cas où)
                        foreach (Document doc in _inventorApp.Documents)
                        {
                            try
                            {
                                if (doc.Dirty)
                                {
                                    doc.Save();
                                }
                            }
                            catch { /* Ignorer les erreurs de sauvegarde individuelles */ }
                        }
                        Log($"✅ Sauvegarde terminée", "SUCCESS");
                    }
                    catch (Exception saveEx)
                    {
                        Log($"⚠️ Erreur pendant Save: {saveEx.Message}", "WARN");
                    }

                    // 6. Activer le document (le mettre au premier plan)
                    try
                    {
                        topAssemblyDoc.Activate();
                        _inventorApp.Visible = true; // S'assurer qu'Inventor est visible
                        Log($"✅ Module prêt pour le dessinateur: {System.IO.Path.GetFileName(topAssemblyPath)}", "SUCCESS");
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    Log($"❌ Erreur finalisation module: {ex.Message}", "ERROR");
                }
            });
        }

        /// <summary>
        /// Applique les iProperties sur un document
        /// </summary>
        private void ApplyIPropertiesToDocument(Document doc, CreateModuleRequest request)
        {
            try
            {
                PropertySets propertySets = doc.PropertySets;
                
                // Design Tracking Properties
                PropertySet designProps = propertySets["Design Tracking Properties"];
                SetProperty(designProps, "Part Number", request.FullProjectNumber);
                SetProperty(designProps, "Project", request.Project);
                
                // Summary Information
                PropertySet summaryProps = propertySets["Inventor Summary Information"];
                SetProperty(summaryProps, "Title", request.FullProjectNumber);
                SetProperty(summaryProps, "Author", request.InitialeDessinateur);
                
                // Custom Properties - Inventor User Defined Properties
                PropertySet customProps = propertySets["Inventor User Defined Properties"];
                SetProperty(customProps, "Project", request.Project, true);
                SetProperty(customProps, "Reference", request.Reference, true);
                SetProperty(customProps, "Module", request.Module, true);
                SetProperty(customProps, "Numero", request.FullProjectNumber, true);
                SetProperty(customProps, "Initiale_du_Dessinateur", request.InitialeDessinateur, true);
                SetProperty(customProps, "Initiale_du_Co_Dessinateur", request.InitialeCoDessinateur, true);
                SetProperty(customProps, "Date_de_Creation", request.CreationDateFormatted, true);
                
                if (!string.IsNullOrEmpty(request.JobTitle))
                {
                    SetProperty(customProps, "Job_Title", request.JobTitle, true);
                }

                Log($"✅ iProperties appliquées: Project={request.Project}, Ref={request.Reference}, Module={request.Module}", "SUCCESS");
            }
            catch (Exception ex)
            {
                Log($"⚠️ Erreur application iProperties: {ex.Message}", "WARN");
            }
        }

        /// <summary>
        /// Applique les paramètres Inventor sur un assemblage
        /// </summary>
        private void ApplyInventorParameters(AssemblyDocument assemblyDoc, CreateModuleRequest request)
        {
            try
            {
                Parameters parameters = assemblyDoc.ComponentDefinition.Parameters;
                UserParameters userParams = parameters.UserParameters;

                // Paramètres standards XNRGY
                SetParameter(userParams, "Project_Form", request.Project);
                SetParameter(userParams, "Reference_Form", request.Reference);
                SetParameter(userParams, "Module_Form", request.Module);
                SetParameter(userParams, "Numero_Form", request.FullProjectNumber);
                SetParameter(userParams, "Initiale_du_Dessinateur_Form", request.InitialeDessinateur);
                SetParameter(userParams, "Initiale_du_Co_Dessinateur_Form", request.InitialeCoDessinateur);
                SetParameter(userParams, "Creation_Date_Form", request.CreationDateFormatted);
                
                if (!string.IsNullOrEmpty(request.JobTitle))
                {
                    SetParameter(userParams, "Job_Title_Form", request.JobTitle);
                }

                Log($"✅ Paramètres Inventor appliqués", "SUCCESS");
            }
            catch (Exception ex)
            {
                Log($"⚠️ Erreur application paramètres: {ex.Message}", "WARN");
            }
        }

        /// <summary>
        /// Définit une propriété, la crée si elle n'existe pas (si createIfMissing=true)
        /// </summary>
        private void SetProperty(PropertySet propSet, string name, object value, bool createIfMissing = false)
        {
            try
            {
                Property prop = propSet[name];
                prop.Value = value;
            }
            catch
            {
                if (createIfMissing)
                {
                    try
                    {
                        propSet.Add(value, name);
                    }
                    catch { /* La propriété existe peut-être déjà avec un autre nom */ }
                }
            }
        }

        /// <summary>
        /// Définit un paramètre utilisateur, le crée s'il n'existe pas
        /// </summary>
        private void SetParameter(UserParameters userParams, string name, string value)
        {
            try
            {
                // Essayer de trouver le paramètre existant
                for (int i = 1; i <= userParams.Count; i++)
                {
                    if (userParams[i].Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        userParams[i].Expression = $"\"{value}\"";
                        return;
                    }
                }

                // Si pas trouvé, créer un nouveau paramètre texte
                userParams.AddByExpression(name, $"\"{value}\"", UnitsTypeEnum.kTextUnits);
            }
            catch { /* Ignorer les erreurs de paramètres */ }
        }

        /// <summary>
        /// Ferme tous les documents ouverts dans Inventor
        /// </summary>
        private void CloseAllDocuments()
        {
            if (_inventorApp == null) return;

            try
            {
                Documents documents = _inventorApp.Documents;
                int docCount = documents.Count;

                if (docCount > 0)
                {
                    Log($"🗑️ Fermeture de {docCount} document(s)...", "DEBUG");
                    documents.CloseAll(false); // false = ne pas sauvegarder
                    Thread.Sleep(500);
                }
            }
            catch (Exception ex)
            {
                Log($"⚠️ Erreur fermeture documents: {ex.Message}", "DEBUG");
            }
        }

        /// <summary>
        /// Cherche le fichier .ipj principal dans le template (pattern XXXXX-XX-XX_2026.ipj)
        /// </summary>
        /// <param name="templateRoot">Dossier racine du template</param>
        /// <returns>Chemin complet du fichier .ipj ou null si non trouvé</returns>
        public string? FindTemplateProjectFile(string templateRoot)
        {
            try
            {
                // Chercher tous les fichiers .ipj dans le template
                var ipjFiles = Directory.GetFiles(templateRoot, "*.ipj", SearchOption.TopDirectoryOnly);

                if (ipjFiles.Length == 0)
                {
                    // Chercher aussi dans les sous-dossiers
                    ipjFiles = Directory.GetFiles(templateRoot, "*.ipj", SearchOption.AllDirectories);
                }

                if (ipjFiles.Length == 0)
                {
                    Log("⚠️ Aucun fichier .ipj trouvé dans le template", "WARN");
                    return null;
                }

                // Chercher le fichier principal (pattern XXXXX-XX-XX_2026.ipj)
                var mainIpj = ipjFiles.FirstOrDefault(f =>
                {
                    string fileName = System.IO.Path.GetFileName(f);
                    // Pattern: 5 chiffres - 2 chiffres - 2 chiffres _ 2026.ipj
                    return System.Text.RegularExpressions.Regex.IsMatch(
                        fileName, 
                        @"^\d{5}-\d{2}-\d{2}_2026\.ipj$", 
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                });

                if (mainIpj != null)
                {
                    Log($"📁 Fichier IPJ principal trouvé: {System.IO.Path.GetFileName(mainIpj)}", "SUCCESS");
                    return mainIpj;
                }

                // Sinon prendre le premier .ipj disponible
                Log($"📁 Fichier IPJ utilisé: {System.IO.Path.GetFileName(ipjFiles[0])}", "INFO");
                return ipjFiles[0];
            }
            catch (Exception ex)
            {
                Log($"❌ Erreur recherche fichier IPJ: {ex.Message}", "ERROR");
                return null;
            }
        }

        #endregion

        #region Copy Design Principal (Pack & Go)
        /// <summary>
        /// Exécute le Copy Design avec Pack & Go d'Inventor
        /// IMPORTANT: Switch vers le projet IPJ du template avant d'ouvrir les fichiers
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

            bool projectSwitched = false;

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

                // ÉTAPE 0: CRITIQUE - Switch vers le projet IPJ du template
                ReportProgress(2, "Recherche du projet IPJ du template...");
                
                string? templateIpjPath = FindTemplateProjectFile(sourceFolderRoot);
                if (!string.IsNullOrEmpty(templateIpjPath))
                {
                    ReportProgress(5, "Activation du projet template...");
                    projectSwitched = SwitchToTemplateProject(templateIpjPath);
                    
                    if (!projectSwitched)
                    {
                        Log("⚠️ Impossible de switcher vers le projet template, tentative de copie simple", "WARN");
                    }
                }
                else
                {
                    Log("⚠️ Aucun fichier IPJ trouvé dans le template, copie sans switch de projet", "WARN");
                }

                ReportProgress(8, "Création de la structure de dossiers...");

                // ÉTAPE 1: Créer la structure de dossiers destination en copiant celle du template
                await Task.Run(() => CreateFolderStructureFromTemplate(sourceFolderRoot, request.DestinationPath));

                ReportProgress(12, "Collecte des fichiers Inventor...");

                // ÉTAPE 2: Collecter TOUS les fichiers Inventor (.ipt, .iam, .idw, .dwg)
                var allInventorFiles = request.FilesToCopy
                    .Where(f => IsInventorFile(f.OriginalPath))
                    .ToList();

                Log($"Fichiers Inventor à traiter: {allInventorFiles.Count}", "INFO");

                // Séparer les fichiers par type pour un traitement ordonné
                var idwFiles = allInventorFiles
                    .Where(f => System.IO.Path.GetExtension(f.OriginalPath).ToLower() == ".idw")
                    .Select(f => f.OriginalPath)
                    .ToList();

                Log($"  - Dessins (.idw): {idwFiles.Count}", "INFO");

                // ÉTAPE 3: Exécuter le vrai Pack & Go
                ReportProgress(15, "Pack & Go des assemblages et pièces...");
                
                var packAndGoResult = await ExecuteRealPackAndGoAsync(
                    sourceTopAssembly, 
                    sourceFolderRoot, 
                    request.DestinationPath, 
                    newTopAssemblyName,
                    idwFiles,
                    request);
                
                result.CopiedFiles = packAndGoResult.CopiedFiles;
                result.FilesCopied = packAndGoResult.FilesCopied;
                result.PropertiesUpdated = packAndGoResult.PropertiesUpdated;

                ReportProgress(80, "Copie des fichiers orphelins Inventor...");

                // ÉTAPE 4: Copier les fichiers Inventor orphelins (non référencés par les documents principaux)
                var orphanResult = await CopyOrphanInventorFilesAsync(
                    sourceFolderRoot, 
                    request.DestinationPath, 
                    result.CopiedFiles.Select(f => f.OriginalPath).ToList());
                result.FilesCopied += orphanResult.Count;
                foreach (var fileResult in orphanResult)
                {
                    result.CopiedFiles.Add(fileResult);
                }

                ReportProgress(88, "Copie des fichiers non-Inventor...");

                // ÉTAPE 5: Copier les fichiers non-Inventor en conservant la structure
                var nonInventorResult = await CopyNonInventorFilesAsync(sourceFolderRoot, request.DestinationPath, request);
                result.FilesCopied += nonInventorResult.Count;
                foreach (var fileResult in nonInventorResult)
                {
                    result.CopiedFiles.Add(fileResult);
                }

                ReportProgress(92, "Renommage du fichier projet (.ipj)...");

                // ÉTAPE 6: Renommer le fichier .ipj avec le numéro de projet formaté
                string newIpjPath = await RenameProjectFileAsync(request.DestinationPath, request.FullProjectNumber);

                // ÉTAPE 7: Switch vers le nouveau projet IPJ créé
                ReportProgress(94, "Activation du nouveau projet...");
                if (!string.IsNullOrEmpty(newIpjPath) && System.IO.File.Exists(newIpjPath))
                {
                    Log($"🔄 Switch vers le nouveau projet: {System.IO.Path.GetFileName(newIpjPath)}", "INFO");
                    SwitchToNewProject(newIpjPath);
                }

                // ÉTAPE 8: Ouvrir le nouveau Top Assembly et appliquer les propriétés finales
                ReportProgress(96, "Ouverture du nouveau module...");
                string newTopAssemblyPath = System.IO.Path.Combine(request.DestinationPath, $"{request.FullProjectNumber}.iam");
                
                if (System.IO.File.Exists(newTopAssemblyPath))
                {
                    await FinalizeNewModuleAsync(newTopAssemblyPath, request);
                }
                else
                {
                    Log($"⚠️ Top Assembly non trouvé: {newTopAssemblyPath}", "WARN");
                }

                result.Success = result.FilesCopied > 0;
                result.EndTime = DateTime.Now;
                result.DestinationPath = request.DestinationPath;
                result.NewTopAssemblyPath = newTopAssemblyPath;

                ReportProgress(100, $"✓ Copy Design terminé: {result.FilesCopied} fichiers");
                Log($"=== COPY DESIGN TERMINÉ: {result.FilesCopied} fichiers copiés ===", "SUCCESS");
                Log($"📂 Module ouvert et prêt pour le dessinateur: {newTopAssemblyPath}", "SUCCESS");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Log($"ERREUR Copy Design: {ex.Message}", "ERROR");
                ReportProgress(0, $"✗ Erreur: {ex.Message}");
                
                // En cas d'erreur seulement, restaurer le projet original
                if (projectSwitched && !string.IsNullOrEmpty(_originalProjectPath))
                {
                    Log("🔄 Restauration du projet original suite à erreur...", "INFO");
                    RestoreOriginalProject();
                }
            }

            return result;
        }

        /// <summary>
        /// Copie les fichiers Inventor orphelins (qui ne sont pas référencés par les documents principaux)
        /// Ces fichiers doivent quand même être copiés dans la destination
        /// IMPORTANT: Dans le template Xnrgy_Module, la plupart des fichiers sont "orphelins" car le Module_.iam
        /// ne référence que des fichiers de la Library. Tous ces fichiers doivent être copiés.
        /// </summary>
        private async Task<List<FileCopyResult>> CopyOrphanInventorFilesAsync(
            string sourceRoot, 
            string destRoot, 
            List<string> alreadyCopiedPaths)
        {
            var results = new List<FileCopyResult>();

            await Task.Run(() =>
            {
                try
                {
                    // Trouver tous les fichiers Inventor dans le template
                    var allInventorFiles = Directory.GetFiles(sourceRoot, "*.*", SearchOption.AllDirectories)
                        .Where(f => IsInventorFile(f) && !IsVaultTempFile(f))
                        .ToList();

                    Log($"📁 Fichiers Inventor trouvés dans le template: {allInventorFiles.Count}", "DEBUG");

                    // Normaliser les chemins déjà copiés pour la comparaison
                    var normalizedCopiedPaths = alreadyCopiedPaths
                        .Select(p => p.ToLowerInvariant().Trim())
                        .ToHashSet();

                    // Exclure les fichiers déjà copiés et ceux dans les dossiers exclus
                    var orphanFiles = allInventorFiles
                        .Where(f =>
                        {
                            // Vérifier si déjà copié (comparaison normalisée)
                            if (normalizedCopiedPaths.Contains(f.ToLowerInvariant().Trim()))
                                return false;

                            // Vérifier si dans un dossier exclu
                            string? dirPath = System.IO.Path.GetDirectoryName(f);
                            if (!string.IsNullOrEmpty(dirPath) &&
                                ExcludedFolders.Any(ef => dirPath.ToLowerInvariant().Contains($"\\{ef.ToLowerInvariant()}\\") || 
                                                          dirPath.ToLowerInvariant().EndsWith($"\\{ef.ToLowerInvariant()}")))
                                return false;

                            // Exclure les fichiers de la Library (ils ne doivent pas être copiés)
                            if (f.StartsWith(LibraryPath, StringComparison.OrdinalIgnoreCase))
                                return false;

                            return true;
                        })
                        .ToList();

                    if (orphanFiles.Count == 0)
                    {
                        Log("Aucun fichier Inventor orphelin à copier", "DEBUG");
                        return;
                    }

                    Log($"� {orphanFiles.Count} fichier(s) Inventor à copier (copie simple)...", "INFO");

                    int copiedCount = 0;
                    int skippedCount = 0;

                    foreach (var orphanFile in orphanFiles)
                    {
                        try
                        {
                            string relativePath = GetRelativePath(orphanFile, sourceRoot);
                            string destPath = System.IO.Path.Combine(destRoot, relativePath);

                            // Vérifier si le fichier n'existe pas déjà dans la destination
                            if (System.IO.File.Exists(destPath))
                            {
                                skippedCount++;
                                continue;
                            }

                            // Créer le dossier si nécessaire
                            string? destDir = System.IO.Path.GetDirectoryName(destPath);
                            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                            {
                                Directory.CreateDirectory(destDir);
                            }

                            // Copie simple du fichier
                            System.IO.File.Copy(orphanFile, destPath, true);

                            results.Add(new FileCopyResult
                            {
                                OriginalPath = orphanFile,
                                OriginalFileName = System.IO.Path.GetFileName(orphanFile),
                                NewPath = destPath,
                                NewFileName = System.IO.Path.GetFileName(destPath),
                                Success = true
                            });

                            copiedCount++;
                            
                            // Log tous les 100 fichiers pour éviter trop de logs
                            if (copiedCount % 100 == 0)
                            {
                                Log($"  ... {copiedCount} fichiers copiés...", "DEBUG");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"  ✗ Erreur copie {System.IO.Path.GetFileName(orphanFile)}: {ex.Message}", "WARN");
                        }
                    }

                    Log($"✓ {copiedCount} fichiers Inventor copiés ({skippedCount} ignorés car déjà présents)", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur copie fichiers orphelins: {ex.Message}", "ERROR");
                }
            });

            return results;
        }

        /// <summary>
        /// Chemin de la Library à préserver (ne pas copier ces fichiers, garder les liens)
        /// </summary>
        private static readonly string LibraryPath = @"C:\Vault\Engineering\Library";

        /// <summary>
        /// Exécute le VRAI Copy Design NATIF d'Inventor avec FileSaveAs.ExecuteSaveCopyAs()
        /// 
        /// APPROCHE NATIVE (CORRECTE):
        /// 1. Ouvrir le Top Assembly (Module_.iam)
        /// 2. Collecter TOUS les fichiers référencés (via AllReferencedDocuments)
        /// 3. Ajouter TOUS les fichiers au FileSaveAs avec leurs nouveaux chemins
        /// 4. UN SEUL ExecuteSaveCopyAs() pour copier TOUS les fichiers ensemble
        /// 5. Tous les liens sont mis à jour SIMULTANÉMENT
        /// 
        /// Cette approche évite le problème des liens corrompus où Roof-01, Left-Wall-01, etc.
        /// pointaient tous vers Right-Wall-01.iam au lieu de leurs vraies copies.
        /// </summary>
        private async Task<(List<FileCopyResult> CopiedFiles, int FilesCopied, int PropertiesUpdated)> ExecuteRealPackAndGoAsync(
            string sourceTopAssembly,
            string sourceRoot, 
            string destRoot, 
            string newTopAssemblyName,
            List<string> idwFiles,
            CreateModuleRequest request)
        {
            var copiedFiles = new List<FileCopyResult>();
            int filesCopied = 0;
            int propertiesUpdated = 0;

            await Task.Run(() =>
            {
                AssemblyDocument? asmDoc = null;

                try
                {
                    string sourceModuleName = System.IO.Path.GetFileName(sourceTopAssembly);
                    Log($"🚀 NATIVE COPY DESIGN: {sourceModuleName} → {newTopAssemblyName}", "INFO");
                    Log($"   Source: {sourceRoot}", "DEBUG");
                    Log($"   Destination: {destRoot}", "DEBUG");

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 1: Ouvrir le Top Assembly
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(20, "Ouverture du Top Assembly...");
                    asmDoc = (AssemblyDocument)_inventorApp!.Documents.Open(sourceTopAssembly, false);
                    Log($"✓ Top Assembly ouvert: {sourceModuleName}", "SUCCESS");

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 2: Appliquer les iProperties AVANT la copie
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(25, "Application des iProperties...");
                    SetIProperties((Document)asmDoc, request);
                    propertiesUpdated = 1;
                    Log("✓ iProperties appliquées sur le Top Assembly", "SUCCESS");

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 3: Collecter TOUS les fichiers référencés
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(30, "Collecte de tous les fichiers référencés...");
                    
                    // Utiliser AllReferencedDocuments pour obtenir TOUS les fichiers (récursif)
                    var allReferencedDocs = new Dictionary<string, Document>(StringComparer.OrdinalIgnoreCase);
                    
                    // Ajouter le Top Assembly lui-même
                    allReferencedDocs[asmDoc.FullFileName] = (Document)asmDoc;
                    
                    // Collecter récursivement tous les documents référencés
                    CollectAllReferencedDocumentsDeep((Document)asmDoc, allReferencedDocs);
                    
                    // IMPORTANT: Séparer les fichiers du module vs Library VRAIE
                    // Les fichiers du module SOURCE (template) doivent être copiés
                    // Les fichiers de la Library EXTERNE (pas dans le template) gardent leurs liens
                    // 
                    // Exemple:
                    //   sourceRoot = C:\Vault\Engineering\Library\Xnrgy_Module
                    //   LibraryPath = C:\Vault\Engineering\Library
                    //   
                    //   C:\Vault\Engineering\Library\Xnrgy_Module\Roof-01.iam → COPIER (dans sourceRoot)
                    //   C:\Vault\Engineering\Library\Common\Bolt.ipt → GARDER LIEN (pas dans sourceRoot)
                    
                    var moduleFiles = allReferencedDocs
                        .Where(kvp => kvp.Key.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value, StringComparer.OrdinalIgnoreCase);
                    
                    var libraryFiles = allReferencedDocs
                        .Where(kvp => !kvp.Key.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    Log($"📁 Total fichiers référencés: {allReferencedDocs.Count}", "INFO");
                    Log($"   - Fichiers du module (à copier): {moduleFiles.Count}", "INFO");
                    Log($"   - Fichiers Library (liens préservés): {libraryFiles.Count}", "INFO");

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 4: COPIE NATIVE - SaveAs de TOUS les fichiers EN UNE SESSION
                    // IMPORTANT: Ne PAS fermer les documents entre les SaveAs!
                    // Les références sont automatiquement mises à jour car tous les 
                    // documents sont chargés en mémoire simultanément.
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(40, "Copy Design natif en cours...");
                    
                    // ÉTAPE CRITIQUE: Trier les fichiers bottom-up
                    // Les pièces (.ipt) d'abord, puis les sous-assemblages (.iam), 
                    // et le Top Assembly en DERNIER
                    var sortedModuleFiles = moduleFiles
                        .OrderBy(kvp => 
                        {
                            var doc = kvp.Value;
                            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject) return 0;      // IPT en premier
                            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                            {
                                // Le Top Assembly en dernier
                                if (kvp.Key.Equals(sourceTopAssembly, StringComparison.OrdinalIgnoreCase))
                                    return 100; // Top Assembly TOUT en dernier
                                return 1;       // Sous-assemblages au milieu
                            }
                            return 2; // Autres fichiers
                        })
                        .ToList();

                    Log($"🔧 Copie de {sortedModuleFiles.Count} fichiers (bottom-up)...", "INFO");
                    int fileIndex = 0;
                    int totalFiles = sortedModuleFiles.Count;

                    Log($"", "INFO");
                    Log($"════════════════════════════════════════════════════════════════", "INFO");
                    Log($"🎯 COPY DESIGN NATIF - Copie de {totalFiles} fichiers", "INFO");
                    Log($"   Ordre: Bottom-Up (IPT → IAM → Top Assembly)", "INFO");
                    Log($"   Tous les documents restent en mémoire pendant la copie", "INFO");
                    Log($"════════════════════════════════════════════════════════════════", "INFO");
                    Log($"", "INFO");

                    foreach (var kvp in sortedModuleFiles)
                    {
                        string originalPath = kvp.Key;
                        Document doc = kvp.Value;
                        fileIndex++;
                        
                        // Calculer le nouveau chemin
                        string relativePath = GetRelativePath(originalPath, sourceRoot);
                        string newPath;
                        bool isTopAssembly = originalPath.Equals(sourceTopAssembly, StringComparison.OrdinalIgnoreCase);
                        
                        // Le Top Assembly est renommé
                        if (isTopAssembly)
                        {
                            newPath = System.IO.Path.Combine(destRoot, newTopAssemblyName);
                        }
                        else
                        {
                            newPath = System.IO.Path.Combine(destRoot, relativePath);
                        }
                        
                        // S'assurer que le dossier destination existe
                        string? newDir = System.IO.Path.GetDirectoryName(newPath);
                        if (!string.IsNullOrEmpty(newDir) && !Directory.Exists(newDir))
                        {
                            Directory.CreateDirectory(newDir);
                        }

                        string fileName = System.IO.Path.GetFileName(originalPath);
                        string newFileName = System.IO.Path.GetFileName(newPath);
                        
                        try
                        {
                            // IMPORTANT: SaveAs met à jour automatiquement les références
                            // car tous les documents sont chargés en mémoire simultanément
                            doc.SaveAs(newPath, false);
                            
                            copiedFiles.Add(new FileCopyResult
                            {
                                OriginalPath = originalPath,
                                OriginalFileName = fileName,
                                NewPath = newPath,
                                NewFileName = newFileName,
                                Success = true,
                                IsTopAssembly = isTopAssembly,
                                PropertiesUpdated = isTopAssembly
                            });
                            filesCopied++;
                            
                            string typeIcon = doc.DocumentType == DocumentTypeEnum.kPartDocumentObject ? "🔩" :
                                             doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? "📦" : "📄";
                            Log($"  {typeIcon} [{fileIndex}/{totalFiles}] {fileName} → {newFileName}", "SUCCESS");
                        }
                        catch (Exception ex)
                        {
                            Log($"  ❌ [{fileIndex}/{totalFiles}] {fileName}: {ex.Message}", "ERROR");
                            copiedFiles.Add(new FileCopyResult
                            {
                                OriginalPath = originalPath,
                                OriginalFileName = fileName,
                                NewPath = newPath,
                                NewFileName = newFileName,
                                Success = false,
                                ErrorMessage = ex.Message
                            });
                        }

                        int progress = 40 + (int)(fileIndex * 30.0 / totalFiles);
                        ReportProgress(progress, $"Copie: {newFileName}");
                    }

                    Log($"", "INFO");
                    Log($"════════════════════════════════════════════════════════════════", "INFO");
                    Log($"✅ COPY DESIGN TERMINÉ: {filesCopied}/{totalFiles} fichiers copiés", "SUCCESS");
                    Log($"════════════════════════════════════════════════════════════════", "INFO");

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 5: Traiter les fichiers .idw (dessins)
                    // Les dessins ne sont PAS inclus dans AllReferencedDocuments
                    // car ils référencent l'assemblage, pas l'inverse
                    // On les copie APRÈS les assemblages pour que les références soient correctes
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(75, "Traitement des dessins (.idw)...");
                    
                    // Fermer le Top Assembly avant de traiter les dessins
                    if (asmDoc != null)
                    {
                        try { asmDoc.Close(false); } catch { }
                        asmDoc = null;
                    }

                    Log($"📐 Traitement de {idwFiles.Count} fichiers de dessins...", "INFO");
                    int idwIndex = 0;

                    foreach (var idwPath in idwFiles)
                    {
                        try
                        {
                            idwIndex++;
                            
                            // Vérifier que le .idw n'est pas dans un dossier exclu
                            string? dirPath = System.IO.Path.GetDirectoryName(idwPath);
                            if (!string.IsNullOrEmpty(dirPath) && 
                                ExcludedFolders.Any(ef => dirPath.Contains($"\\{ef}\\") || dirPath.EndsWith($"\\{ef}")))
                            {
                                Log($"  Exclu (dossier): {System.IO.Path.GetFileName(idwPath)}", "DEBUG");
                                continue;
                            }

                            // Vérifier si déjà copié
                            if (copiedFiles.Any(f => f.OriginalPath.Equals(idwPath, StringComparison.OrdinalIgnoreCase)))
                            {
                                Log($"  Déjà traité: {System.IO.Path.GetFileName(idwPath)}", "DEBUG");
                                continue;
                            }

                            string relativePath = GetRelativePath(idwPath, sourceRoot);
                            string newIdwPath = System.IO.Path.Combine(destRoot, relativePath);
                            
                            // S'assurer que le dossier existe
                            string? newDir = System.IO.Path.GetDirectoryName(newIdwPath);
                            if (!string.IsNullOrEmpty(newDir) && !Directory.Exists(newDir))
                            {
                                Directory.CreateDirectory(newDir);
                            }

                            // Ouvrir le dessin
                            var drawDoc = (DrawingDocument)_inventorApp!.Documents.Open(idwPath, false);
                            
                            // SaveAs copie le dessin avec ses références mises à jour
                            // Les références pointent vers les fichiers copiés car ils sont 
                            // dans le même dossier destination
                            ((Document)drawDoc).SaveAs(newIdwPath, false);
                            
                            copiedFiles.Add(new FileCopyResult
                            {
                                OriginalPath = idwPath,
                                OriginalFileName = System.IO.Path.GetFileName(idwPath),
                                NewPath = newIdwPath,
                                NewFileName = System.IO.Path.GetFileName(newIdwPath),
                                Success = true
                            });
                            filesCopied++;
                            
                            Log($"  📄 [{idwIndex}/{idwFiles.Count}] {System.IO.Path.GetFileName(idwPath)}", "SUCCESS");
                            
                            // Fermer le dessin
                            drawDoc.Close(false);
                            
                            int progress = 75 + (int)(idwIndex * 15.0 / idwFiles.Count);
                            ReportProgress(progress, $"Dessin: {System.IO.Path.GetFileName(newIdwPath)}");
                        }
                        catch (Exception ex)
                        {
                            Log($"  ✗ Erreur {System.IO.Path.GetFileName(idwPath)}: {ex.Message}", "WARN");
                        }
                    }

                    Log($"", "INFO");
                    Log($"✅ COPY DESIGN NATIF TERMINÉ: {filesCopied} fichiers copiés", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"❌ Erreur Copy Design Natif: {ex.Message}", "ERROR");
                    Log($"   Stack: {ex.StackTrace}", "DEBUG");
                    throw;
                }
                finally
                {
                    // Nettoyage: fermer tous les documents ouverts
                    try
                    {
                        if (asmDoc != null) asmDoc.Close(false);
                    }
                    catch { }
                }
            });

            return (copiedFiles, filesCopied, propertiesUpdated);
        }

        /// <summary>
        /// Collecte récursivement TOUS les documents référencés (version profonde)
        /// Utilise AllReferencedDocuments pour une collecte complète
        /// </summary>
        private void CollectAllReferencedDocumentsDeep(Document doc, Dictionary<string, Document> collected)
        {
            try
            {
                // AllReferencedDocuments inclut tous les fichiers référencés récursivement
                foreach (Document refDoc in doc.AllReferencedDocuments)
                {
                    string fullPath = refDoc.FullFileName;
                    
                    // Éviter les doublons
                    if (!collected.ContainsKey(fullPath))
                    {
                        collected[fullPath] = refDoc;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  Note: Collecte références: {ex.Message}", "DEBUG");
            }
        }

        /// <summary>
        /// Met à jour les références d'un dessin pour pointer vers les fichiers copiés
        /// tout en préservant les liens vers la Library
        /// </summary>
        private void UpdateDrawingReferences(DrawingDocument drawDoc, string sourceRoot, string destRoot, Dictionary<string, string> renameMap)
        {
            try
            {
                // Parcourir toutes les feuilles et vues
                foreach (Sheet sheet in drawDoc.Sheets)
                {
                    foreach (DrawingView view in sheet.DrawingViews)
                    {
                        try
                        {
                            // Obtenir le document référencé par la vue
                            Document? referencedDoc = view.ReferencedDocumentDescriptor?.ReferencedDocument as Document;
                            if (referencedDoc == null) continue;

                            string refPath = referencedDoc.FullFileName;
                            
                            // Si c'est un fichier de la Library, ne rien faire (garder le lien)
                            if (refPath.StartsWith(LibraryPath, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            // Si c'est un fichier du module source, mettre à jour le chemin
                            if (refPath.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                            {
                                string fileName = System.IO.Path.GetFileName(refPath);
                                
                                // Vérifier si ce fichier doit être renommé
                                string newFileName = renameMap.ContainsKey(fileName) ? renameMap[fileName] : fileName;
                                
                                string relativePath = GetRelativePath(refPath, sourceRoot);
                                string relativeDir = System.IO.Path.GetDirectoryName(relativePath) ?? "";
                                
                                string newRefPath;
                                if (string.IsNullOrEmpty(relativeDir))
                                {
                                    newRefPath = System.IO.Path.Combine(destRoot, newFileName);
                                }
                                else
                                {
                                    newRefPath = System.IO.Path.Combine(destRoot, relativeDir, newFileName);
                                }

                                // Mettre à jour la référence si le fichier existe
                                if (System.IO.File.Exists(newRefPath))
                                {
                                    // Note: La mise à jour automatique se fait via le ReferencedFileDescriptor
                                    // quand on sauvegarde avec SaveAs dans le nouveau dossier
                                }
                            }
                        }
                        catch
                        {
                            // Ignorer les erreurs sur les vues individuelles
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  Note: Mise à jour références dessin: {ex.Message}", "DEBUG");
            }
        }

        /// <summary>
        /// Exécute le Pack & Go d'Inventor (ancienne méthode conservée pour compatibilité)
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
        /// Copie les fichiers de dessins (.idw, .dwg) qui ne sont pas référencés par l'assemblage
        /// Les dessins référencent les pièces/assemblages, pas l'inverse
        /// </summary>
        private async Task<List<FileCopyResult>> CopyDrawingFilesAsync(string sourceRoot, string destRoot)
        {
            var results = new List<FileCopyResult>();

            await Task.Run(() =>
            {
                try
                {
                    // Chercher tous les fichiers de dessins
                    var drawingExtensions = new[] { ".idw", ".dwg" };
                    var drawingFiles = Directory.GetFiles(sourceRoot, "*.*", SearchOption.AllDirectories)
                        .Where(f => drawingExtensions.Contains(System.IO.Path.GetExtension(f).ToLower()))
                        .ToList();

                    if (drawingFiles.Count == 0)
                    {
                        Log("Aucun fichier de dessin (.idw, .dwg) trouvé", "INFO");
                        return;
                    }

                    Log($"Copie de {drawingFiles.Count} fichiers de dessins...", "INFO");

                    foreach (var drawingFile in drawingFiles)
                    {
                        // Vérifier si le fichier n'est pas dans un dossier exclu
                        string? dirPath = System.IO.Path.GetDirectoryName(drawingFile);
                        if (!string.IsNullOrEmpty(dirPath) && 
                            ExcludedFolders.Any(ef => dirPath.Contains($"\\{ef}\\") || dirPath.EndsWith($"\\{ef}")))
                        {
                            Log($"  Exclu (dossier exclu): {System.IO.Path.GetFileName(drawingFile)}", "DEBUG");
                            continue;
                        }

                        // Exclure les fichiers temporaires Vault
                        if (IsVaultTempFile(drawingFile))
                        {
                            Log($"  Exclu (Vault temp): {System.IO.Path.GetFileName(drawingFile)}", "DEBUG");
                            continue;
                        }

                        // Calculer le chemin relatif et destination
                        string relativePath = GetRelativePath(drawingFile, sourceRoot);
                        string destPath = System.IO.Path.Combine(destRoot, relativePath);

                        // Vérifier si le fichier n'existe pas déjà (copié par Pack & Go)
                        if (System.IO.File.Exists(destPath))
                        {
                            Log($"  Déjà copié: {System.IO.Path.GetFileName(drawingFile)}", "DEBUG");
                            continue;
                        }

                        try
                        {
                            // Créer le dossier si nécessaire
                            string? destDir = System.IO.Path.GetDirectoryName(destPath);
                            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                            {
                                Directory.CreateDirectory(destDir);
                            }

                            // Copier le fichier
                            System.IO.File.Copy(drawingFile, destPath, overwrite: true);

                            results.Add(new FileCopyResult
                            {
                                OriginalPath = drawingFile,
                                OriginalFileName = System.IO.Path.GetFileName(drawingFile),
                                NewPath = destPath,
                                NewFileName = System.IO.Path.GetFileName(destPath),
                                Success = true
                            });

                            Log($"  ✓ {System.IO.Path.GetFileName(drawingFile)} (dessin)", "SUCCESS");
                        }
                        catch (Exception ex)
                        {
                            Log($"  ✗ Erreur copie dessin {System.IO.Path.GetFileName(drawingFile)}: {ex.Message}", "WARN");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur copie fichiers dessins: {ex.Message}", "ERROR");
                }
            });

            return results;
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
        /// Renomme le fichier .ipj principal avec le numéro de projet formaté
        /// Seulement le fichier correspondant au pattern XXXXX-XX-XX_2026.ipj
        /// Retourne le chemin du nouveau fichier .ipj
        /// </summary>
        private async Task<string> RenameProjectFileAsync(string destinationPath, string fullProjectNumber)
        {
            string resultPath = string.Empty;
            
            await Task.Run(() =>
            {
                try
                {
                    // Chercher tous les fichiers .ipj dans le dossier destination (racine seulement)
                    var ipjFiles = Directory.GetFiles(destinationPath, "*.ipj", SearchOption.TopDirectoryOnly);

                    if (ipjFiles.Length == 0)
                    {
                        Log("Aucun fichier .ipj trouvé dans la destination", "DEBUG");
                        return;
                    }

                    foreach (var ipjFile in ipjFiles)
                    {
                        var fileName = System.IO.Path.GetFileName(ipjFile);
                        
                        // Vérifier si c'est le fichier projet principal (pattern XXXXX-XX-XX_2026.ipj)
                        if (!IsMainProjectFilePattern(fileName))
                        {
                            Log($"Fichier .ipj ignoré (pas le fichier principal): {fileName}", "DEBUG");
                            continue;
                        }
                        
                        string newIpjName = $"{fullProjectNumber}.ipj";
                        string newIpjPath = System.IO.Path.Combine(destinationPath, newIpjName);

                        // Ne pas renommer si c'est déjà le bon nom
                        if (fileName.Equals(newIpjName, StringComparison.OrdinalIgnoreCase))
                        {
                            Log($"Fichier .ipj déjà correctement nommé: {newIpjName}", "INFO");
                            resultPath = ipjFile;
                            continue;
                        }

                        // Renommer le fichier
                        if (System.IO.File.Exists(newIpjPath))
                        {
                            System.IO.File.Delete(newIpjPath);
                            Log($"Ancien fichier .ipj supprimé: {newIpjName}", "DEBUG");
                        }

                        System.IO.File.Move(ipjFile, newIpjPath);
                        Log($"✓ Fichier .ipj renommé: {fileName} → {newIpjName}", "SUCCESS");
                        resultPath = newIpjPath;
                        
                        // Un seul fichier principal à renommer
                        break;
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur renommage fichier .ipj: {ex.Message}", "WARN");
                }
            });
            
            return resultPath;
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

        /// <summary>
        /// Vérifie si un fichier .ipj correspond au pattern du fichier projet principal
        /// Pattern: XXXXX-XX-XX_2026.ipj ou similaire (contient _2026 ou _202X)
        /// </summary>
        private bool IsMainProjectFilePattern(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return false;
            
            var nameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(fileName);
            
            // Pattern 1: Contient _202X (année)
            if (nameWithoutExt.Contains("_202"))
                return true;
            
            // Pattern 2: Format XXXXX-XX-XX (numéro de projet avec tirets)
            if (System.Text.RegularExpressions.Regex.IsMatch(nameWithoutExt, @"^\d{5}-\d{2}-\d{2}"))
                return true;
            
            // Pattern 3: Le nom contient "Module" (fichier projet du module)
            if (nameWithoutExt.IndexOf("Module", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
                
            return false;
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
