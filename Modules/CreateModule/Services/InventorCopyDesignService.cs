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
using XnrgyEngineeringAutomationTools.Modules.CreateModule.Models;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Modules.CreateModule.Services
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
                Log($"[>] Switch vers projet template: {System.IO.Path.GetFileName(templateIpjPath)}", "INFO");

                DesignProjectManager designProjectManager = _inventorApp.DesignProjectManager;

                // Sauvegarder le projet actif actuel
                try
                {
                    DesignProject activeProject = designProjectManager.ActiveDesignProject;
                    if (activeProject != null)
                    {
                        _originalProjectPath = activeProject.FullFileName;
                        Log($"[i] Projet actuel sauvegardé: {System.IO.Path.GetFileName(_originalProjectPath)}", "DEBUG");
                    }
                }
                catch
                {
                    _originalProjectPath = null;
                }

                // Vérifier que le fichier IPJ du template existe
                if (!System.IO.File.Exists(templateIpjPath))
                {
                    Log($"[-] Fichier IPJ template introuvable: {templateIpjPath}", "ERROR");
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
                        Log($"[+] Projet template trouvé dans la collection", "DEBUG");
                        break;
                    }
                }

                // Si pas trouvé, le charger
                if (templateProject == null)
                {
                    Log($"[i] Chargement du projet template: {System.IO.Path.GetFileName(templateIpjPath)}", "DEBUG");
                    templateProject = projectsCollection.AddExisting(templateIpjPath);
                }

                // Activer le projet template
                if (templateProject != null)
                {
                    templateProject.Activate();
                    Thread.Sleep(1000); // Attendre que le switch soit effectif
                    Log($"[+] Projet template activé: {System.IO.Path.GetFileName(templateIpjPath)}", "SUCCESS");
                    return true;
                }
                else
                {
                    Log("[-] Impossible de charger le projet template", "ERROR");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur switch projet template: {ex.Message}", "ERROR");
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
                Log($"[>] Restauration projet original: {System.IO.Path.GetFileName(_originalProjectPath)}", "INFO");

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
                    Log($"[+] Projet original restauré: {System.IO.Path.GetFileName(_originalProjectPath)}", "SUCCESS");
                    return true;
                }
                else
                {
                    Log($"[!] Impossible de restaurer le projet original", "WARN");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"[!] Erreur restauration projet: {ex.Message}", "WARN");
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
                Log($"[>] Switch vers nouveau projet: {System.IO.Path.GetFileName(newIpjPath)}", "INFO");

                DesignProjectManager designProjectManager = _inventorApp.DesignProjectManager;

                // Vérifier que le fichier IPJ existe
                if (!System.IO.File.Exists(newIpjPath))
                {
                    Log($"[-] Fichier IPJ introuvable: {newIpjPath}", "ERROR");
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
                        Log($"[+] Nouveau projet trouvé dans la collection", "DEBUG");
                        break;
                    }
                }

                // Si pas trouvé, le charger
                if (newProject == null)
                {
                    Log($"[i] Chargement du nouveau projet: {System.IO.Path.GetFileName(newIpjPath)}", "DEBUG");
                    newProject = projectsCollection.AddExisting(newIpjPath);
                }

                // Activer le nouveau projet
                if (newProject != null)
                {
                    newProject.Activate();
                    Thread.Sleep(1000);
                    Log($"[+] Nouveau projet activé: {System.IO.Path.GetFileName(newIpjPath)}", "SUCCESS");
                    return true;
                }
                else
                {
                    Log("[-] Impossible de charger le nouveau projet", "ERROR");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur switch nouveau projet: {ex.Message}", "ERROR");
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
                    Log($"[i] Ouverture du nouveau module: {System.IO.Path.GetFileName(topAssemblyPath)}", "INFO");

                    // 1. Ouvrir le Top Assembly
                    Document topAssemblyDoc = _inventorApp.Documents.Open(topAssemblyPath, true);
                    
                    if (topAssemblyDoc == null)
                    {
                        Log($"[-] Impossible d'ouvrir: {topAssemblyPath}", "ERROR");
                        return;
                    }

                    Log($"[+] Document ouvert: {System.IO.Path.GetFileName(topAssemblyPath)}", "SUCCESS");

                    // 2. Appliquer les iProperties
                    Log($"[>] Application des iProperties...", "INFO");
                    ApplyIPropertiesToDocument(topAssemblyDoc, request);

                    // 3. Appliquer les paramètres Inventor (si c'est un assemblage)
                    if (topAssemblyDoc is AssemblyDocument assemblyDoc)
                    {
                        Log($"[>] Application des paramètres Inventor...", "INFO");
                        ApplyInventorParameters(assemblyDoc, request);
                    }

                    // 4. Update All (rebuild de l'assemblage)
                    Log($"[>] Update All (rebuild)...", "INFO");
                    try
                    {
                        topAssemblyDoc.Update2(true); // true = full update
                        Log($"[+] Update terminé", "SUCCESS");
                    }
                    catch (Exception updateEx)
                    {
                        Log($"[!] Erreur pendant Update: {updateEx.Message}", "WARN");
                    }

                    // 4.5 Préparer la vue: cacher workfeatures, vue ISO, zoom all
                    Log($"[>] Préparation de la vue...", "INFO");
                    try
                    {
                        PrepareViewForDesigner(topAssemblyDoc);
                        Log($"[+] Vue préparée (ISO, Zoom All, Workfeatures cachés)", "SUCCESS");
                    }
                    catch (Exception viewEx)
                    {
                        Log($"[!] Note: Préparation vue: {viewEx.Message}", "DEBUG");
                    }

                    // 5. Save All
                    Log($"[i] Save All...", "INFO");
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
                        Log($"[+] Sauvegarde terminée", "SUCCESS");
                    }
                    catch (Exception saveEx)
                    {
                        Log($"[!] Erreur pendant Save: {saveEx.Message}", "WARN");
                    }

                    // 6. Activer le document (le mettre au premier plan)
                    try
                    {
                        topAssemblyDoc.Activate();
                        _inventorApp.Visible = true; // S'assurer qu'Inventor est visible
                        Log($"[+] Module prêt pour le dessinateur: {System.IO.Path.GetFileName(topAssemblyPath)}", "SUCCESS");
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    Log($"[-] Erreur finalisation module: {ex.Message}", "ERROR");
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
                SetProperty(customProps, "Numero_de_Projet", request.FullProjectNumber, true);
                SetProperty(customProps, "Initiale_du_Dessinateur", request.InitialeDessinateur, true);
                SetProperty(customProps, "Initiale_du_Co_Dessinateur", request.InitialeCoDessinateur, true);
                SetProperty(customProps, "Creation_Date", request.CreationDateFormatted, true);
                
                if (!string.IsNullOrEmpty(request.JobTitle))
                {
                    SetProperty(customProps, "Job_Title", request.JobTitle, true);
                }

                Log($"[+] iProperties appliquées: Project={request.Project}, Ref={request.Reference}, Module={request.Module}", "SUCCESS");
            }
            catch (Exception ex)
            {
                Log($"[!] Erreur application iProperties: {ex.Message}", "WARN");
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

                Log($"[+] Paramètres Inventor appliqués", "SUCCESS");
            }
            catch (Exception ex)
            {
                Log($"[!] Erreur application paramètres: {ex.Message}", "WARN");
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
                    Log($"[i] Fermeture de {docCount} document(s)...", "DEBUG");
                    documents.CloseAll(false); // false = ne pas sauvegarder
                    Thread.Sleep(500);
                }
            }
            catch (Exception ex)
            {
                Log($"[!] Erreur fermeture documents: {ex.Message}", "DEBUG");
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
                    Log("[!] Aucun fichier .ipj trouvé dans le template", "WARN");
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
                    Log($"[i] Fichier IPJ principal trouvé: {System.IO.Path.GetFileName(mainIpj)}", "SUCCESS");
                    return mainIpj;
                }

                // Sinon prendre le premier .ipj disponible
                Log($"[i] Fichier IPJ utilisé: {System.IO.Path.GetFileName(ipjFiles[0])}", "INFO");
                return ipjFiles[0];
            }
            catch (Exception ex)
            {
                Log($"[-] Erreur recherche fichier IPJ: {ex.Message}", "ERROR");
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

                // Trouver le fichier Top Assembly (.iam)
                // Pour templates: chercher Module_.iam
                // Pour projets existants: chercher n'importe quel .iam à la racine
                var topAssemblyFile = request.FilesToCopy
                    .FirstOrDefault(f => f.IsTopAssembly || 
                                         f.OriginalFileName.Equals("Module_.iam", StringComparison.OrdinalIgnoreCase));

                if (topAssemblyFile == null)
                {
                    // Chercher n'importe quel .iam dans la racine (pas dans les sous-dossiers Equipment/Floor/Wall)
                    topAssemblyFile = request.FilesToCopy
                        .FirstOrDefault(f => System.IO.Path.GetExtension(f.OriginalPath).ToLower() == ".iam" &&
                                             !f.OriginalPath.Contains("1-Equipment") &&
                                             !f.OriginalPath.Contains("2-Floor") &&
                                             !f.OriginalPath.Contains("3-Wall") &&
                                             string.IsNullOrEmpty(System.IO.Path.GetDirectoryName(f.RelativePath)));
                }

                if (topAssemblyFile == null)
                {
                    throw new Exception("Aucun fichier Top Assembly (.iam) trouvé à la racine du projet source");
                }

                string sourceTopAssembly = topAssemblyFile.OriginalPath;
                string sourceFolderRoot = System.IO.Path.GetDirectoryName(sourceTopAssembly) ?? "";
                
                // Nouveau nom pour le Top Assembly: numéro formaté (ex: 123450101.iam)
                string newTopAssemblyName = $"{request.FullProjectNumber}.iam";
                
                Log($"Top Assembly source: {topAssemblyFile.OriginalFileName}", "INFO");
                Log($"Nouveau nom: {newTopAssemblyName}", "INFO");
                Log($"Destination: {request.DestinationPath}", "INFO");

                // ÉTAPE 0: CRITIQUE - Switch vers le projet IPJ
                // Pour TOUS les cas, utiliser l'IPJ du dossier source (que ce soit template ou projet existant)
                ReportProgress(2, "Recherche du projet IPJ...");
                
                string? sourceIpjPath = FindTemplateProjectFile(sourceFolderRoot);
                
                if (!string.IsNullOrEmpty(sourceIpjPath))
                {
                    Log($"[i] Fichier IPJ source trouvé: {System.IO.Path.GetFileName(sourceIpjPath)}", "INFO");
                    ReportProgress(5, "Activation du projet source...");
                    projectSwitched = SwitchToTemplateProject(sourceIpjPath);
                    
                    if (!projectSwitched)
                    {
                        Log("[!] Impossible de switcher vers le projet source, tentative de copie simple", "WARN");
                    }
                }
                else
                {
                    Log("[!] Aucun fichier IPJ trouvé dans le dossier source, copie sans switch de projet", "WARN");
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
                    result.CopiedFiles.Select(f => f.OriginalPath).ToList(),
                    request);
                result.FilesCopied += orphanResult.Count;
                foreach (var fileResult in orphanResult)
                {
                    result.CopiedFiles.Add(fileResult);
                }

                // ═══════════════════════════════════════════════════════════════════════════
                // ÉTAPE 4bis: Mise à jour des références vers les fichiers copiés (incluant supprimés)
                // Maintenant que TOUS les fichiers sont copiés (actifs + orphelins/supprimés),
                // on peut mettre à jour toutes les références correctement
                // ═══════════════════════════════════════════════════════════════════════════
                ReportProgress(85, "Mise a jour des references...");
                
                await Task.Run(() =>
                {
                    try
                    {
                        Log($"[>] Mise a jour des references dans les assemblages...", "INFO");
                        
                        // Construire le renameMapByName à partir de tous les fichiers copiés
                        var renameMapByName2 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var cf in result.CopiedFiles)
                        {
                            if (!string.IsNullOrEmpty(cf.OriginalPath) && !string.IsNullOrEmpty(cf.NewPath))
                            {
                                string origName = System.IO.Path.GetFileName(cf.OriginalPath).ToLowerInvariant();
                                string newName = System.IO.Path.GetFileName(cf.NewPath);
                                if (!origName.Equals(newName, StringComparison.OrdinalIgnoreCase))
                                {
                                    renameMapByName2[origName] = newName;
                                }
                            }
                        }
                        
                        // Construire le pathMapping complet
                        var pathMapping2 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var cf in result.CopiedFiles)
                        {
                            if (!string.IsNullOrEmpty(cf.OriginalPath) && !string.IsNullOrEmpty(cf.NewPath))
                            {
                                pathMapping2[cf.OriginalPath] = cf.NewPath;
                            }
                        }
                        
                        // Mettre à jour les références dans tous les assemblages (.iam)
                        var allAssemblies = result.CopiedFiles
                            .Where(f => f.Success && f.NewPath.ToLowerInvariant().EndsWith(".iam"))
                            .OrderBy(f => System.IO.Path.GetFileName(f.NewPath).Equals($"{request.FullProjectNumber}.iam", StringComparison.OrdinalIgnoreCase) ? 1 : 0)
                            .ToList();
                        
                        int totalRefs = 0;
                        foreach (var asmFile in allAssemblies)
                        {
                            try
                            {
                                int refs = UpdateAssemblyReferencesComplete(
                                    asmFile.NewPath,
                                    sourceFolderRoot,
                                    request.DestinationPath,
                                    pathMapping2,
                                    renameMapByName2);
                                totalRefs += refs;
                            }
                            catch (Exception)
                            {
                                // Ignorer silencieusement
                            }
                        }
                        
                        Log($"[+] {totalRefs} references MAJ dans {allAssemblies.Count} assemblages", "SUCCESS");
                        
                        // Mettre à jour les références dans tous les dessins (.idw)
                        // Les dessins peuvent référencer directement des sous-assemblages supprimés
                        var allDrawings = result.CopiedFiles
                            .Where(f => f.Success && f.NewPath.ToLowerInvariant().EndsWith(".idw"))
                            .ToList();
                        
                        int totalDrawingRefs = 0;
                        foreach (var idwFile in allDrawings)
                        {
                            try
                            {
                                int refs = UpdateAssemblyReferencesComplete(
                                    idwFile.NewPath,
                                    sourceFolderRoot,
                                    request.DestinationPath,
                                    pathMapping2,
                                    renameMapByName2);
                                totalDrawingRefs += refs;
                            }
                            catch (Exception)
                            {
                                // Ignorer silencieusement
                            }
                        }
                        
                        if (totalDrawingRefs > 0)
                        {
                            Log($"[+] {totalDrawingRefs} references MAJ dans {allDrawings.Count} dessins", "SUCCESS");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"[!] Erreur MAJ refs: {ex.Message}", "WARN");
                    }
                });

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
                string newIpjPath = await RenameProjectFileAsync(request.DestinationPath, request.FullProjectNumber, null, request.Source);

                // ÉTAPE 7: Switch vers le nouveau projet IPJ créé
                ReportProgress(94, "Activation du nouveau projet...");
                if (!string.IsNullOrEmpty(newIpjPath) && System.IO.File.Exists(newIpjPath))
                {
                    Log($"[>] Switch vers le nouveau projet: {System.IO.Path.GetFileName(newIpjPath)}", "INFO");
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
                    Log($"[!] Top Assembly non trouvé: {newTopAssemblyPath}", "WARN");
                }

                result.Success = result.FilesCopied > 0;
                result.EndTime = DateTime.Now;
                result.DestinationPath = request.DestinationPath;
                result.NewTopAssemblyPath = newTopAssemblyPath;

                ReportProgress(100, $"[+] Copy Design terminé: {result.FilesCopied} fichiers");
                Log($"=== COPY DESIGN TERMINÉ: {result.FilesCopied} fichiers copiés ===", "SUCCESS");
                Log($"[i] Module ouvert et prêt pour le dessinateur: {newTopAssemblyPath}", "SUCCESS");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Log($"ERREUR Copy Design: {ex.Message}", "ERROR");
                ReportProgress(0, $"[+] Erreur: {ex.Message}");
                
                // En cas d'erreur seulement, restaurer le projet original
                if (projectSwitched && !string.IsNullOrEmpty(_originalProjectPath))
                {
                    Log("[>] Restauration du projet original suite à erreur...", "INFO");
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
        /// Utilise les NewFileName définis dans le request pour appliquer préfixes/suffixes/renommages.
        /// </summary>
        private async Task<List<FileCopyResult>> CopyOrphanInventorFilesAsync(
            string sourceRoot, 
            string destRoot, 
            List<string> alreadyCopiedPaths,
            CreateModuleRequest request)
        {
            var results = new List<FileCopyResult>();

            // Créer un dictionnaire pour lookup rapide des renommages
            var renameMap = request.FilesToCopy
                .Where(f => !string.IsNullOrEmpty(f.NewFileName) && f.NewFileName != f.OriginalFileName)
                .ToDictionary(f => f.OriginalPath.ToLowerInvariant(), f => f.NewFileName);

            await Task.Run(() =>
            {
                try
                {
                    // Trouver tous les fichiers Inventor dans le template
                    var allInventorFiles = Directory.GetFiles(sourceRoot, "*.*", SearchOption.AllDirectories)
                        .Where(f => IsInventorFile(f) && !IsVaultTempFile(f))
                        .ToList();

                    Log($"[i] Fichiers Inventor trouvés dans le template: {allInventorFiles.Count}", "DEBUG");

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

                            // Exclure UNIQUEMENT les IPT Typical Drawing (partagés entre projets)
                            if (f.StartsWith(IPTTypicalDrawingPath, StringComparison.OrdinalIgnoreCase))
                                return false;

                            return true;
                        })
                        .ToList();

                    if (orphanFiles.Count == 0)
                    {
                        Log("Aucun fichier Inventor orphelin à copier", "DEBUG");
                        return;
                    }

                    Log($"[>] Copie de {orphanFiles.Count} fichier(s) Inventor à copier (copie simple)...", "INFO");

                    int copiedCount = 0;
                    int skippedCount = 0;
                    int renamedCount = 0;

                    foreach (var orphanFile in orphanFiles)
                    {
                        try
                        {
                            string relativePath = GetRelativePath(orphanFile, sourceRoot);
                            string originalFileName = System.IO.Path.GetFileName(orphanFile);
                            
                            // Vérifier si ce fichier a un nouveau nom défini (préfixe/suffixe/renommage)
                            string newFileName = originalFileName;
                            if (renameMap.TryGetValue(orphanFile.ToLowerInvariant(), out string? customName) && !string.IsNullOrEmpty(customName))
                            {
                                newFileName = customName;
                                renamedCount++;
                            }
                            
                            // Construire le chemin destination avec le nouveau nom
                            string relativeDir = System.IO.Path.GetDirectoryName(relativePath) ?? "";
                            string destPath;
                            if (string.IsNullOrEmpty(relativeDir))
                            {
                                destPath = System.IO.Path.Combine(destRoot, newFileName);
                            }
                            else
                            {
                                destPath = System.IO.Path.Combine(destRoot, relativeDir, newFileName);
                            }

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
                                OriginalFileName = originalFileName,
                                NewPath = destPath,
                                NewFileName = newFileName,
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

                    Log($"✓ {copiedCount} fichiers Inventor copiés ({renamedCount} renommés, {skippedCount} ignorés car déjà présents)", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur copie fichiers orphelins: {ex.Message}", "ERROR");
                }
            });

            return results;
        }

        /// <summary>
        /// Chemin des IPT Typical Drawing partagés entre projets - NE PAS COPIER, garder les liens
        /// C'est le SEUL chemin à exclure, pas tout C:\Vault\Engineering\Library
        /// </summary>
        private static readonly string IPTTypicalDrawingPath = @"C:\Vault\Engineering\Library\Cabinet\IPT_Typical_Drawing";

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
                    Log($"[>] NATIVE COPY DESIGN: {sourceModuleName} -> {newTopAssemblyName}", "INFO");
                    Log($"   Source: {sourceRoot}", "DEBUG");
                    Log($"   Destination: {destRoot}", "DEBUG");

                    // ══════════════════════════════════════════════════════════════════
                    // PRÉPARATION: Créer un dictionnaire pour les renommages (préfixe/suffixe)
                    // IMPORTANT: Utilisé pour tous les fichiers enfants, pas seulement les orphelins
                    // Normaliser les chemins en minuscules pour éviter les problèmes de casse
                    // ══════════════════════════════════════════════════════════════════
                    var renameMapByPath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    var renameMapByName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    
                    foreach (var f in request.FilesToCopy.Where(f => !string.IsNullOrEmpty(f.NewFileName) && f.NewFileName != f.OriginalFileName))
                    {
                        // Normaliser le chemin (remplacer / par \ et utiliser lowercase)
                        string normalizedPath = f.OriginalPath.Replace("/", "\\").ToLowerInvariant();
                        if (!renameMapByPath.ContainsKey(normalizedPath))
                        {
                            renameMapByPath[normalizedPath] = f.NewFileName;
                        }
                        
                        // Aussi indexer par nom de fichier seul (fallback)
                        string fileName = System.IO.Path.GetFileName(f.OriginalPath).ToLowerInvariant();
                        if (!renameMapByName.ContainsKey(fileName))
                        {
                            renameMapByName[fileName] = f.NewFileName;
                        }
                    }
                    
                    if (renameMapByPath.Count > 0)
                    {
                        Log($"[i] {renameMapByPath.Count} fichiers avec renommage (prefix/suffix)", "INFO");
                        foreach (var kvp in renameMapByPath.Take(5)) // Log les 5 premiers pour debug
                        {
                            Log($"    {System.IO.Path.GetFileName(kvp.Key)} -> {kvp.Value}", "DEBUG");
                        }
                        if (renameMapByPath.Count > 5)
                        {
                            Log($"    ... et {renameMapByPath.Count - 5} autres", "DEBUG");
                        }
                    }

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 1: Ouvrir le Top Assembly
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(20, "Ouverture du Top Assembly...");
                    asmDoc = (AssemblyDocument)_inventorApp!.Documents.Open(sourceTopAssembly, false);
                    Log($"[+] Top Assembly ouvert: {sourceModuleName}", "SUCCESS");

                    // NOTE: Les iProperties seront appliquées APRÈS le switch vers le nouveau IPJ
                    // dans FinalizeNewModuleAsync() - c'est la bonne séquence

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 2: Collecter TOUS les fichiers référencés
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(30, "Collecte de tous les fichiers référencés...");
                    
                    // Utiliser AllReferencedDocuments pour obtenir TOUS les fichiers (récursif)
                    var allReferencedDocs = new Dictionary<string, Document>(StringComparer.OrdinalIgnoreCase);
                    
                    // Ajouter le Top Assembly lui-même
                    allReferencedDocs[asmDoc.FullFileName] = (Document)asmDoc;
                    
                    // Collecter récursivement tous les documents référencés
                    CollectAllReferencedDocumentsDeep((Document)asmDoc, allReferencedDocs);
                    
                    // IMPORTANT: Séparer les fichiers du module vs IPT Typical Drawing
                    // Les fichiers du module SOURCE (template) doivent être copiés
                    // Les fichiers IPT Typical Drawing (partagés) gardent leurs liens
                    // 
                    // Exemple:
                    //   sourceRoot = C:\Vault\Engineering\Library\Xnrgy_Module
                    //   IPTTypicalDrawingPath = C:\Vault\Engineering\Library\Cabinet\IPT_Typical_Drawing
                    //   
                    //   C:\Vault\Engineering\Library\Xnrgy_Module\Roof-01.iam → COPIER (dans sourceRoot)
                    //   C:\Vault\Engineering\Library\Cabinet\IPT_Typical_Drawing\Bolt.ipt → GARDER LIEN
                    
                    var moduleFiles = allReferencedDocs
                        .Where(kvp => kvp.Key.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value, StringComparer.OrdinalIgnoreCase);
                    
                    var libraryFiles = allReferencedDocs
                        .Where(kvp => !kvp.Key.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    Log($"[i] Total fichiers référencés: {allReferencedDocs.Count}", "INFO");
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

                    Log($"[>] Copie de {sortedModuleFiles.Count} fichiers (bottom-up)...", "INFO");
                    int fileIndex = 0;
                    int totalFiles = sortedModuleFiles.Count;

                    Log($"", "INFO");
                    Log($"════════════════════════════════════════════════════════════════", "INFO");
                    Log($"[i] COPY DESIGN NATIF - Copie de {totalFiles} fichiers", "INFO");
                    Log($"   Ordre: Bottom-Up (IPT -> IAM -> Top Assembly)", "INFO");
                    Log($"   Tous les documents restent en memoire pendant la copie", "INFO");
                    Log($"════════════════════════════════════════════════════════════════", "INFO");
                    Log($"", "INFO");

                    foreach (var kvp in sortedModuleFiles)
                    {
                        string originalPath = kvp.Key;
                        Document doc = kvp.Value;
                        fileIndex++;
                        
                        // Calculer le nouveau chemin
                        string relativePath = GetRelativePath(originalPath, sourceRoot);
                        string originalFileName = System.IO.Path.GetFileName(originalPath);
                        string newPath;
                        string newFileName;
                        bool isTopAssembly = originalPath.Equals(sourceTopAssembly, StringComparison.OrdinalIgnoreCase);
                        
                        // Le Top Assembly est renommé avec le numéro de projet
                        if (isTopAssembly)
                        {
                            newFileName = newTopAssemblyName;
                            newPath = System.IO.Path.Combine(destRoot, newTopAssemblyName);
                        }
                        else
                        {
                            // IMPORTANT: Vérifier si ce fichier a un renommage (préfixe/suffixe)
                            // Normaliser le chemin pour la recherche
                            string normalizedOriginalPath = originalPath.Replace("/", "\\").ToLowerInvariant();
                            string fileNameLower = originalFileName.ToLowerInvariant();
                            
                            // D'abord chercher par chemin complet, puis par nom de fichier
                            if (renameMapByPath.TryGetValue(normalizedOriginalPath, out string? customName) && !string.IsNullOrEmpty(customName))
                            {
                                newFileName = customName;
                                Log($"    [>] Renommage (chemin): {originalFileName} -> {newFileName}", "DEBUG");
                            }
                            else if (renameMapByName.TryGetValue(fileNameLower, out string? customNameByFileName) && !string.IsNullOrEmpty(customNameByFileName))
                            {
                                newFileName = customNameByFileName;
                                Log($"    [>] Renommage (nom): {originalFileName} -> {newFileName}", "DEBUG");
                            }
                            else
                            {
                                // Pas de renommage, utiliser le nom original
                                newFileName = originalFileName;
                            }
                            
                            // Construire le chemin avec le nouveau nom dans le bon dossier relatif
                            string? relativeDir = System.IO.Path.GetDirectoryName(relativePath);
                            if (string.IsNullOrEmpty(relativeDir))
                            {
                                newPath = System.IO.Path.Combine(destRoot, newFileName);
                            }
                            else
                            {
                                newPath = System.IO.Path.Combine(destRoot, relativeDir, newFileName);
                            }
                        }
                        
                        // S'assurer que le dossier destination existe
                        string? newDir = System.IO.Path.GetDirectoryName(newPath);
                        if (!string.IsNullOrEmpty(newDir) && !Directory.Exists(newDir))
                        {
                            Directory.CreateDirectory(newDir);
                        }
                        
                        try
                        {
                            // IMPORTANT: SaveAs met à jour automatiquement les références
                            // car tous les documents sont chargés en mémoire simultanément
                            doc.SaveAs(newPath, false);
                            
                            copiedFiles.Add(new FileCopyResult
                            {
                                OriginalPath = originalPath,
                                OriginalFileName = originalFileName,
                                NewPath = newPath,
                                NewFileName = newFileName,
                                Success = true,
                                IsTopAssembly = isTopAssembly,
                                PropertiesUpdated = isTopAssembly
                            });
                            filesCopied++;
                            
                            string typeIcon = doc.DocumentType == DocumentTypeEnum.kPartDocumentObject ? "[#]" :
                                             doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? "[+]" : "[i]";
                            Log($"  {typeIcon} [{fileIndex}/{totalFiles}] {originalFileName} -> {newFileName}", "SUCCESS");
                        }
                        catch (Exception ex)
                        {
                            Log($"  [-] [{fileIndex}/{totalFiles}] {originalFileName}: {ex.Message}", "ERROR");
                            copiedFiles.Add(new FileCopyResult
                            {
                                OriginalPath = originalPath,
                                OriginalFileName = originalFileName,
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
                    Log($"[+] COPY DESIGN TERMINE: {filesCopied}/{totalFiles} fichiers copies", "SUCCESS");
                    Log($"════════════════════════════════════════════════════════════════", "INFO");

                    // ══════════════════════════════════════════════════════════════════
                    // CRITIQUE: FERMER TOUS LES DOCUMENTS AVANT de continuer
                    // Les références seront mises à jour dans la 2ème passe (après copie des fichiers orphelins)
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(72, "Fermeture des documents...");
                    Log($"[>] Fermeture de tous les documents...", "DEBUG");
                    try
                    {
                        _inventorApp.Documents.CloseAll(true); // true = skip save prompts
                    }
                    catch (Exception closeEx)
                    {
                        Log($"[!] Erreur fermeture: {closeEx.Message}", "DEBUG");
                    }
                    System.Threading.Thread.Sleep(500); // Attendre qu'Inventor soit stable
                    asmDoc = null; // Le document n'est plus valide

                    // ══════════════════════════════════════════════════════════════════
                    // PHASE 5: Traiter les fichiers .idw (dessins)
                    // Les dessins ne sont PAS inclus dans AllReferencedDocuments
                    // car ils référencent l'assemblage, pas l'inverse
                    // On doit mettre à jour leurs références AVANT de les copier
                    // ══════════════════════════════════════════════════════════════════
                    ReportProgress(75, "Traitement des dessins (.idw)...");

                    // Créer un dictionnaire des chemins source → destination pour les références
                    var pathMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var cf in copiedFiles)
                    {
                        if (!string.IsNullOrEmpty(cf.OriginalPath) && !string.IsNullOrEmpty(cf.NewPath))
                        {
                            pathMapping[cf.OriginalPath] = cf.NewPath;
                        }
                    }
                    // Ajouter le mapping du Top Assembly
                    string originalTopPath = System.IO.Path.Combine(sourceRoot, "Module_.iam");
                    string newTopPath = System.IO.Path.Combine(destRoot, newTopAssemblyName);
                    pathMapping[originalTopPath] = newTopPath;
                    
                    // Créer un dictionnaire des renommages (nom fichier uniquement)
                    // Pour le Top Assembly Module_.iam → nouveau nom ET tous les fichiers avec préfixe/suffixe
                    var idwRenameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    idwRenameMap["Module_.iam"] = newTopAssemblyName;
                    
                    // Ajouter aussi les renommages pour les fichiers référencés (préfixe/suffixe)
                    // Utiliser renameMapByName qui contient les noms de fichiers sans chemin
                    foreach (var kvp in renameMapByName)
                    {
                        if (!idwRenameMap.ContainsKey(kvp.Key))
                        {
                            idwRenameMap[kvp.Key] = kvp.Value;
                        }
                    }

                    Log($"[i] Traitement de {idwFiles.Count} fichiers de dessins...", "INFO");
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
                                Log($"  Deja traite: {System.IO.Path.GetFileName(idwPath)}", "DEBUG");
                                continue;
                            }

                            string relativePath = GetRelativePath(idwPath, sourceRoot);
                            string idwOriginalFileName = System.IO.Path.GetFileName(idwPath);
                            string idwNewFileName;
                            string newIdwPath;
                            
                            // Vérifier si ce fichier IDW a un renommage (préfixe/suffixe)
                            // Chercher d'abord par chemin, puis par nom de fichier
                            string normalizedIdwPath = idwPath.Replace("/", "\\").ToLowerInvariant();
                            string idwFileNameLower = idwOriginalFileName.ToLowerInvariant();
                            
                            if (renameMapByPath.TryGetValue(normalizedIdwPath, out string? customIdwName) && !string.IsNullOrEmpty(customIdwName))
                            {
                                idwNewFileName = customIdwName;
                            }
                            else if (renameMapByName.TryGetValue(idwFileNameLower, out string? customIdwNameByName) && !string.IsNullOrEmpty(customIdwNameByName))
                            {
                                idwNewFileName = customIdwNameByName;
                            }
                            else
                            {
                                idwNewFileName = idwOriginalFileName;
                            }
                            
                            // Construire le nouveau chemin
                            string? relativeDir = System.IO.Path.GetDirectoryName(relativePath);
                            if (string.IsNullOrEmpty(relativeDir))
                            {
                                newIdwPath = System.IO.Path.Combine(destRoot, idwNewFileName);
                            }
                            else
                            {
                                newIdwPath = System.IO.Path.Combine(destRoot, relativeDir, idwNewFileName);
                            }
                            
                            // S'assurer que le dossier existe
                            string? newDir = System.IO.Path.GetDirectoryName(newIdwPath);
                            if (!string.IsNullOrEmpty(newDir) && !Directory.Exists(newDir))
                            {
                                Directory.CreateDirectory(newDir);
                            }

                            // Ouvrir le dessin
                            var drawDoc = (DrawingDocument)_inventorApp!.Documents.Open(idwPath, false);
                            
                            // ═══════════════════════════════════════════════════════════
                            // IMPORTANT: Mettre à jour les références du dessin
                            // AVANT de faire le SaveAs
                            // ═══════════════════════════════════════════════════════════
                            UpdateDrawingReferencesWithMapping(drawDoc, sourceRoot, destRoot, pathMapping, idwRenameMap);
                            
                            // SaveAs vers la nouvelle destination
                            ((Document)drawDoc).SaveAs(newIdwPath, false);
                            
                            copiedFiles.Add(new FileCopyResult
                            {
                                OriginalPath = idwPath,
                                OriginalFileName = idwOriginalFileName,
                                NewPath = newIdwPath,
                                NewFileName = idwNewFileName,
                                Success = true
                            });
                            filesCopied++;
                            
                            Log($"  [i] [{idwIndex}/{idwFiles.Count}] {idwOriginalFileName} -> {idwNewFileName}", "SUCCESS");
                            
                            // Fermer le dessin
                            drawDoc.Close(false);
                            
                            int progress = 75 + (int)(idwIndex * 15.0 / idwFiles.Count);
                            ReportProgress(progress, $"Dessin: {idwNewFileName}");
                        }
                        catch (Exception ex)
                        {
                            Log($"  [-] Erreur {System.IO.Path.GetFileName(idwPath)}: {ex.Message}", "WARN");
                        }
                    }

                    Log($"", "INFO");
                    Log($"[+] COPY DESIGN NATIF TERMINE: {filesCopied} fichiers copies", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"[-] Erreur Copy Design Natif: {ex.Message}", "ERROR");
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
        /// Utilise ReferencedFileDescriptor.ReplaceReference() pour changer les liens
        /// </summary>
        private void UpdateDrawingReferencesWithMapping(
            DrawingDocument drawDoc, 
            string sourceRoot, 
            string destRoot, 
            Dictionary<string, string> pathMapping,
            Dictionary<string, string> renameMap)
        {
            try
            {
                Document doc = (Document)drawDoc;
                
                // Utiliser ReferencedFileDescriptors pour mettre à jour toutes les références
                foreach (ReferencedFileDescriptor refDesc in doc.ReferencedFileDescriptors)
                {
                    try
                    {
                        string refPath = refDesc.FullFileName;
                        
                        // Si c'est un fichier IPT Typical Drawing (partagé), garder le lien original
                        if (refPath.StartsWith(IPTTypicalDrawingPath, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        // Vérifier si on a un mapping direct
                        if (pathMapping.TryGetValue(refPath, out string? newPath))
                        {
                            if (System.IO.File.Exists(newPath))
                            {
                                // Utiliser PutLogicalFileNameUsingFull pour changer la référence
                                refDesc.PutLogicalFileNameUsingFull(newPath);
                                Log($"    ↪ {System.IO.Path.GetFileName(refPath)} → {System.IO.Path.GetFileName(newPath)}", "DEBUG");
                            }
                            continue;
                        }

                        // Si c'est un fichier du module source, calculer le nouveau chemin
                        if (refPath.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        {
                            string fileName = System.IO.Path.GetFileName(refPath);
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
                                // Remplacer le nom de fichier dans le chemin relatif
                                newRefPath = System.IO.Path.Combine(destRoot, relativeDir, newFileName);
                            }

                            if (System.IO.File.Exists(newRefPath))
                            {
                                // Utiliser PutLogicalFileNameUsingFull pour changer la référence
                                refDesc.PutLogicalFileNameUsingFull(newRefPath);
                                Log($"    ↪ {System.IO.Path.GetFileName(refPath)} → {System.IO.Path.GetFileName(newRefPath)}", "DEBUG");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"    Note: Référence {System.IO.Path.GetFileName(refDesc.FullFileName)}: {ex.Message}", "DEBUG");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  Note: Mise à jour références dessin: {ex.Message}", "DEBUG");
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
                            
                            // Si c'est un fichier IPT Typical Drawing (partagé), ne rien faire (garder le lien)
                            if (refPath.StartsWith(IPTTypicalDrawingPath, StringComparison.OrdinalIgnoreCase))
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
                Log($"  Note: Mise a jour references dessin: {ex.Message}", "DEBUG");
            }
        }

        /// <summary>
        /// Met à jour les références d'un assemblage pour pointer vers les fichiers renommés
        /// Utilisé quand des fichiers ont été copiés avec préfixe/suffixe
        /// </summary>
        private void UpdateAssemblyReferences(string assemblyPath, Dictionary<string, string> pathMapping)
        {
            if (_inventorApp == null) return;
            
            try
            {
                // Ouvrir l'assemblage
                var doc = _inventorApp.Documents.Open(assemblyPath, false);
                bool modified = false;
                
                // Parcourir tous les ReferencedFileDescriptors
                foreach (ReferencedFileDescriptor refDesc in doc.ReferencedFileDescriptors)
                {
                    try
                    {
                        string refPath = refDesc.FullFileName;
                        
                        // Vérifier si cette référence a un nouveau chemin dans le mapping
                        if (pathMapping.TryGetValue(refPath, out string? newPath) && !string.IsNullOrEmpty(newPath))
                        {
                            // Vérifier que le nouveau fichier existe
                            if (System.IO.File.Exists(newPath) && !refPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                            {
                                refDesc.PutLogicalFileNameUsingFull(newPath);
                                modified = true;
                                Log($"    [>] {System.IO.Path.GetFileName(refPath)} -> {System.IO.Path.GetFileName(newPath)}", "DEBUG");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"    [!] Note reference {System.IO.Path.GetFileName(refDesc.FullFileName)}: {ex.Message}", "DEBUG");
                    }
                }
                
                // Sauvegarder si modifié
                if (modified)
                {
                    doc.Save();
                    Log($"  [+] References MAJ: {System.IO.Path.GetFileName(assemblyPath)}", "DEBUG");
                }
                
                // Fermer le document
                doc.Close(false);
            }
            catch (Exception ex)
            {
                Log($"  [!] Erreur MAJ assemblage {System.IO.Path.GetFileName(assemblyPath)}: {ex.Message}", "WARN");
            }
        }

        /// <summary>
        /// Met à jour les références d'un assemblage avec un mapping COMPLET
        /// Spécialement conçu pour les composants SUPPRIMÉS qui ne sont pas chargés en mémoire
        /// Utilise à la fois le mapping exact ET une correspondance par chemin source/dest
        /// </summary>
        private int UpdateAssemblyReferencesComplete(
            string assemblyPath, 
            string sourceRoot, 
            string destRoot,
            Dictionary<string, string> pathMapping,
            Dictionary<string, string> renameMapByName)
        {
            if (_inventorApp == null) return 0;
            
            int updatedCount = 0;
            
            try
            {
                // Ouvrir l'assemblage (INVISIBLE pour performance)
                var doc = _inventorApp.Documents.Open(assemblyPath, false);
                bool modified = false;
                
                // Parcourir TOUS les ReferencedFileDescriptors (inclut les supprimés!)
                int refCount = 0;
                foreach (ReferencedFileDescriptor refDesc in doc.ReferencedFileDescriptors)
                {
                    refCount++;
                    try
                    {
                        string refPath = refDesc.FullFileName;
                        string refFileName = System.IO.Path.GetFileName(refPath);
                        string refFileNameLower = refFileName.ToLowerInvariant();
                        string? newPath = null;
                        
                        // MÉTHODE 1: Chercher dans le mapping exact (chemin complet source -> dest)
                        if (pathMapping.TryGetValue(refPath, out string? exactPath) && !string.IsNullOrEmpty(exactPath))
                        {
                            newPath = exactPath;
                        }
                        // MÉTHODE 2: Référence pointe vers sourceRoot (composants supprimés)
                        else if (refPath.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase))
                        {
                            string relativePath = GetRelativePath(refPath, sourceRoot);
                            string newFileName;
                            
                            // Vérifier si le fichier doit être renommé
                            if (renameMapByName.TryGetValue(refFileNameLower, out string? renamedName) && !string.IsNullOrEmpty(renamedName))
                            {
                                newFileName = renamedName;
                            }
                            else
                            {
                                newFileName = refFileName;
                            }
                            
                            string? relativeDir = System.IO.Path.GetDirectoryName(relativePath);
                            if (string.IsNullOrEmpty(relativeDir))
                            {
                                newPath = System.IO.Path.Combine(destRoot, newFileName);
                            }
                            else
                            {
                                newPath = System.IO.Path.Combine(destRoot, relativeDir, newFileName);
                            }
                            
                            // Si le chemin calculé n'existe pas, RECHERCHER le fichier dans toute la destination
                            if (!System.IO.File.Exists(newPath))
                            {
                                // Chercher le fichier renommé dans toute la destination
                                var foundFiles = System.IO.Directory.GetFiles(destRoot, newFileName, System.IO.SearchOption.AllDirectories);
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
                        // MÉTHODE 3: Référence pointe vers destRoot avec ANCIEN nom (composants actifs après SaveAs)
                        else if (refPath.StartsWith(destRoot, StringComparison.OrdinalIgnoreCase))
                        {
                            // Le fichier est déjà dans destRoot, vérifier si le nom doit être mis à jour
                            if (renameMapByName.TryGetValue(refFileNameLower, out string? renamedName) && !string.IsNullOrEmpty(renamedName))
                            {
                                // Le fichier a un nouveau nom avec préfixe/suffixe
                                string relativePath = GetRelativePath(refPath, destRoot);
                                string? relativeDir = System.IO.Path.GetDirectoryName(relativePath);
                                
                                if (string.IsNullOrEmpty(relativeDir))
                                {
                                    newPath = System.IO.Path.Combine(destRoot, renamedName);
                                }
                                else
                                {
                                    newPath = System.IO.Path.Combine(destRoot, relativeDir, renamedName);
                                }
                            }
                        }
                        
                        // Appliquer la mise à jour si un nouveau chemin valide est trouvé
                        if (!string.IsNullOrEmpty(newPath) && 
                            System.IO.File.Exists(newPath) && 
                            !refPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                        {
                            refDesc.PutLogicalFileNameUsingFull(newPath);
                            modified = true;
                            updatedCount++;
                        }
                    }
                    catch (Exception)
                    {
                        // Ignorer silencieusement les erreurs de références non modifiables
                    }
                }
                
                // Sauvegarder si modifié
                if (modified)
                {
                    doc.Save();
                    Log($"  [+] {updatedCount} refs MAJ: {System.IO.Path.GetFileName(assemblyPath)}", "SUCCESS");
                }
                
                // Fermer le document
                doc.Close(false);
            }
            catch (Exception ex)
            {
                Log($"  [!] Erreur MAJ: {System.IO.Path.GetFileName(assemblyPath)}: {ex.Message}", "WARN");
            }
            
            return updatedCount;
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
                    
                    Log($"Pack & Go: {referencedDocs.Count + 1} fichiers Inventor a traiter", "INFO");
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
        /// Utilise les NewFileName définis dans le request pour appliquer préfixes/suffixes/renommages.
        /// </summary>
        private async Task<List<FileCopyResult>> CopyNonInventorFilesAsync(string sourceRoot, string destRoot, CreateModuleRequest request)
        {
            var results = new List<FileCopyResult>();

            // Créer un dictionnaire pour lookup rapide des renommages
            var renameMap = request.FilesToCopy
                .Where(f => !string.IsNullOrEmpty(f.NewFileName) && f.NewFileName != f.OriginalFileName)
                .ToDictionary(f => f.OriginalPath.ToLowerInvariant(), f => f.NewFileName);

            await Task.Run(() =>
            {
                try
                {
                    // Parcourir tous les fichiers du dossier source
                    var allFiles = Directory.GetFiles(sourceRoot, "*.*", SearchOption.AllDirectories);
                    
                    Log($"Analyse: {allFiles.Length} fichiers trouvés dans le template", "INFO");
                    int renamedCount = 0;

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

                        // Calculer le chemin relatif
                        string relativePath = GetRelativePath(filePath, sourceRoot);
                        string originalFileName = System.IO.Path.GetFileName(filePath);
                        
                        // Vérifier si ce fichier a un nouveau nom défini (préfixe/suffixe/renommage)
                        string newFileName = originalFileName;
                        if (renameMap.TryGetValue(filePath.ToLowerInvariant(), out string? customName) && !string.IsNullOrEmpty(customName))
                        {
                            newFileName = customName;
                            renamedCount++;
                        }
                        
                        // Construire le chemin destination avec le nouveau nom
                        string relativeDir = System.IO.Path.GetDirectoryName(relativePath) ?? "";
                        string destPath;
                        if (string.IsNullOrEmpty(relativeDir))
                        {
                            destPath = System.IO.Path.Combine(destRoot, newFileName);
                        }
                        else
                        {
                            destPath = System.IO.Path.Combine(destRoot, relativeDir, newFileName);
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
                            System.IO.File.Copy(filePath, destPath, overwrite: true);

                            results.Add(new FileCopyResult
                            {
                                OriginalPath = filePath,
                                OriginalFileName = originalFileName,
                                NewPath = destPath,
                                NewFileName = newFileName,
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

        /// <summary>
        /// Prépare la vue pour le dessinateur:
        /// - Met le Design View sur "Default" (pas "Primary")
        /// - Désactive l'affichage de tous les Workfeatures (plans, axes, points, UCS)
        /// - Fait un Zoom All pour voir tout le modèle
        /// </summary>
        private void PrepareViewForDesigner(Document doc)
        {
            try
            {
                // 1. Mettre le Design View sur "Default" (au lieu de "Primary")
                try
                {
                    if (doc is AssemblyDocument asmDoc)
                    {
                        var designViewReps = asmDoc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations;
                        
                        // Chercher et activer "Default"
                        foreach (DesignViewRepresentation dvr in designViewReps)
                        {
                            if (dvr.Name.Equals("Default", StringComparison.OrdinalIgnoreCase) ||
                                dvr.Name.Equals("Défaut", StringComparison.OrdinalIgnoreCase))
                            {
                                dvr.Activate();
                                Log($"  ✓ Design View: {dvr.Name}", "DEBUG");
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"  Note: Design View: {ex.Message}", "DEBUG");
                }

                // 2. Désactiver l'affichage de tous les Workfeatures (All Work Features OFF)
                try
                {
                    if (doc is AssemblyDocument asmDoc)
                    {
                        // Désactiver les plans d'origine
                        foreach (WorkPlane wp in asmDoc.ComponentDefinition.WorkPlanes)
                        {
                            try { wp.Visible = false; } catch { }
                        }
                        // Désactiver les axes d'origine
                        foreach (WorkAxis wa in asmDoc.ComponentDefinition.WorkAxes)
                        {
                            try { wa.Visible = false; } catch { }
                        }
                        // Désactiver les points d'origine
                        foreach (WorkPoint wpt in asmDoc.ComponentDefinition.WorkPoints)
                        {
                            try { wpt.Visible = false; } catch { }
                        }
                        
                        // Désactiver l'Origin Folder (contient XY, XZ, YZ planes, X, Y, Z axes, Center Point)
                        try
                        {
                            asmDoc.ComponentDefinition.WorkPlanes["XY Plane"].Visible = false;
                            asmDoc.ComponentDefinition.WorkPlanes["XZ Plane"].Visible = false;
                            asmDoc.ComponentDefinition.WorkPlanes["YZ Plane"].Visible = false;
                        }
                        catch { }
                        
                        try
                        {
                            asmDoc.ComponentDefinition.WorkAxes["X Axis"].Visible = false;
                            asmDoc.ComponentDefinition.WorkAxes["Y Axis"].Visible = false;
                            asmDoc.ComponentDefinition.WorkAxes["Z Axis"].Visible = false;
                        }
                        catch { }
                        
                        try
                        {
                            asmDoc.ComponentDefinition.WorkPoints["Center Point"].Visible = false;
                        }
                        catch { }
                    }
                    Log("  ✓ Workfeatures cachés", "DEBUG");
                }
                catch (Exception ex)
                {
                    Log($"  Note: Masquage workfeatures: {ex.Message}", "DEBUG");
                }

                // 3. Vue Isométrique (ISO)
                try
                {
                    View? activeView = _inventorApp?.ActiveView;
                    if (activeView != null)
                    {
                        Camera camera = activeView.Camera;
                        camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
                        camera.Apply();
                        Log("  ✓ Vue ISO appliquée", "DEBUG");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  Note: Vue ISO: {ex.Message}", "DEBUG");
                }

                // 4. Zoom All (Fit)
                try
                {
                    View? activeView = _inventorApp?.ActiveView;
                    if (activeView != null)
                    {
                        activeView.Fit();
                        Log("  ✓ Zoom All (Fit)", "DEBUG");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  Note: Zoom All: {ex.Message}", "DEBUG");
                }
            }
            catch (Exception ex)
            {
                Log($"  Erreur préparation vue: {ex.Message}", "DEBUG");
            }
        }

        #endregion

        #region Project File Renaming

        /// <summary>
        /// Renomme le fichier .ipj principal avec le numéro de projet formaté
        /// Seulement le fichier correspondant au pattern XXXXX-XX-XX_2026.ipj
        /// Si aucun .ipj n'est trouvé (cas projet existant), copie depuis le template
        /// Pour les projets existants, accepte TOUT fichier .ipj à la racine (pas de pattern requis)
        /// Retourne le chemin du nouveau fichier .ipj
        /// </summary>
        private async Task<string> RenameProjectFileAsync(string destinationPath, string fullProjectNumber, string templatePath = null, CreateModuleSource source = CreateModuleSource.FromTemplate)
        {
            string resultPath = string.Empty;
            bool isFromExistingProject = source == CreateModuleSource.FromExistingProject;
            
            await Task.Run(() =>
            {
                try
                {
                    // Chercher tous les fichiers .ipj dans le dossier destination (racine seulement)
                    var ipjFiles = Directory.GetFiles(destinationPath, "*.ipj", SearchOption.TopDirectoryOnly);

                    if (ipjFiles.Length == 0)
                    {
                        // Aucun .ipj trouvé - Copier depuis le template (cas projet existant)
                        Log("Aucun fichier .ipj dans la destination - Création depuis template...", "INFO");
                        
                        // Chercher le .ipj dans le template par défaut
                        string defaultTemplatePath = @"C:\Vault\Engineering\Library\Xnrgy_Module";
                        string sourceTemplatePath = !string.IsNullOrEmpty(templatePath) ? templatePath : defaultTemplatePath;
                        
                        if (Directory.Exists(sourceTemplatePath))
                        {
                            var templateIpjFiles = Directory.GetFiles(sourceTemplatePath, "*.ipj", SearchOption.TopDirectoryOnly);
                            if (templateIpjFiles.Length > 0)
                            {
                                // Prendre le premier .ipj du template
                                string templateIpj = templateIpjFiles[0];
                                string newIpjName = $"{fullProjectNumber}.ipj";
                                string newIpjPath = System.IO.Path.Combine(destinationPath, newIpjName);
                                
                                // Copier et renommer
                                System.IO.File.Copy(templateIpj, newIpjPath, true);
                                Log($"[+] Fichier .ipj créé depuis template: {newIpjName}", "SUCCESS");
                                resultPath = newIpjPath;
                            }
                            else
                            {
                                Log("[!] Aucun .ipj trouvé dans le template", "WARN");
                            }
                        }
                        else
                        {
                            Log($"[!] Dossier template non trouvé: {sourceTemplatePath}", "WARN");
                        }
                        return;
                    }

                    foreach (var ipjFile in ipjFiles)
                    {
                        var fileName = System.IO.Path.GetFileName(ipjFile);
                        
                        // Pour les projets existants: accepter TOUT fichier .ipj à la racine
                        // Pour les templates: vérifier si c'est le fichier projet principal (pattern XXXXX-XX-XX_2026.ipj)
                        if (!isFromExistingProject && !IsMainProjectFilePattern(fileName))
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
            
            // Pattern 4: Format XXXXXXXXX (9 chiffres) - fichier déjà renommé au format projet
            // Ex: 123450101.ipj (Projet 12345, Ref 01, Module 01)
            if (System.Text.RegularExpressions.Regex.IsMatch(nameWithoutExt, @"^\d{9}$"))
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
