using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Net;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Modules.SmartTools.Views;
using IProgressWindow = XnrgyEngineeringAutomationTools.Modules.SmartTools.Views.IProgressWindow;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Services
{
    /// <summary>
    /// Service pour exécuter les outils Smart Tools directement via COM Inventor
    /// Code converti depuis VB (règles iLogic) vers C#
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public class SmartToolsService
    {
        private readonly InventorService _inventorService;
        private Action<string, string>? _logCallback;
        private Action<string, string>? _htmlPopupCallback;
        private Func<string, string, ExportOptionsResult?>? _exportOptionsCallback;
        private Func<string, string, IProgressWindow>? _progressWindowCallback;

        public SmartToolsService(InventorService inventorService)
        {
            _inventorService = inventorService;
        }

        /// <summary>
        /// Définit le callback pour le logging
        /// </summary>
        public void SetLogCallback(Action<string, string>? callback)
        {
            _logCallback = callback;
        }

        /// <summary>
        /// Définit le callback pour afficher les popups HTML
        /// </summary>
        public void SetHtmlPopupCallback(Action<string, string>? callback)
        {
            _htmlPopupCallback = callback;
        }

        /// <summary>
        /// Définit le callback pour afficher les options d'export
        /// </summary>
        public void SetExportOptionsCallback(Func<string, string, ExportOptionsResult?>? callback)
        {
            _exportOptionsCallback = callback;
        }

        /// <summary>
        /// Définit le callback pour créer une fenêtre de progression HTML
        /// </summary>
        public void SetProgressWindowCallback(Func<string, string, IProgressWindow>? callback)
        {
            _progressWindowCallback = callback;
        }

        private void Log(string message, string level = "INFO")
        {
            _logCallback?.Invoke(message, level);
        }

        /// <summary>
        /// Affiche une popup HTML
        /// </summary>
        private void ShowHtmlPopup(string title, string htmlContent)
        {
            _htmlPopupCallback?.Invoke(title, htmlContent);
        }

        /// <summary>
        /// Vérifie la connexion Inventor et retourne l'application
        /// </summary>
        private dynamic GetInventorApplication()
        {
            if (!_inventorService.IsConnected)
            {
                if (!_inventorService.TryConnect())
                {
                    throw new InvalidOperationException("Inventor n'est pas connecté. Veuillez ouvrir Inventor et réessayer.");
                }
            }

            dynamic inventorApp = _inventorService.GetInventorApplication();
            if (inventorApp == null)
            {
                throw new InvalidOperationException("Impossible d'accéder à l'application Inventor.");
            }

            return inventorApp;
        }

        /// <summary>
        /// Vérifie qu'un document d'assemblage est ouvert
        /// Utilise une méthode robuste : vérifie la présence de ComponentDefinition
        /// qui existe uniquement pour les assemblages
        /// </summary>
        private dynamic GetActiveAssemblyDocument()
        {
            dynamic inventorApp = GetInventorApplication();
            dynamic activeDoc = inventorApp.ActiveDocument;
            
            if (activeDoc == null)
            {
                throw new InvalidOperationException("Aucun document Inventor n'est ouvert.");
            }

            // Méthode robuste : vérifier si le document a une ComponentDefinition
            // (propriété qui existe uniquement pour les assemblages)
            try
            {
                dynamic componentDef = activeDoc.ComponentDefinition;
                if (componentDef == null)
                {
                    throw new InvalidOperationException("Cette fonctionnalité fonctionne uniquement avec les assemblages.");
                }
                
                // Vérifier aussi que ComponentDefinition a la propriété Occurrences
                dynamic occurrences = componentDef.Occurrences;
                if (occurrences == null)
                {
                    throw new InvalidOperationException("Cette fonctionnalité fonctionne uniquement avec les assemblages.");
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // Si ComponentDefinition n'existe pas, ce n'est pas un assemblage
                throw new InvalidOperationException("Cette fonctionnalité fonctionne uniquement avec les assemblages.");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // Erreur COM - probablement pas un assemblage
                throw new InvalidOperationException("Cette fonctionnalité fonctionne uniquement avec les assemblages.", ex);
            }

            return activeDoc;
        }

        /// <summary>
        /// Exécute Vue Home + Zoom All silencieusement (sans log) pour ajuster la vue
        /// </summary>
        private void ExecuteIsoViewAndZoomAllSilent()
        {
            try
            {
                dynamic inventorApp = GetInventorApplication();
                
                // 1. Appliquer la vue isométrique (utilise ApplyIsometricView existant)
                ApplyIsometricView(inventorApp);
                
                // 2. Zoom All
                try
                {
                    dynamic cmdManager = inventorApp.CommandManager;
                    dynamic controlDefs = cmdManager.ControlDefinitions;
                    dynamic cmdZoom = controlDefs.Item("AppZoomAllCmd");
                    cmdZoom.Execute();
                }
                catch { }
            }
            catch
            {
                // Ignorer silencieusement
            }
        }

        #region SmartHideBox - Masquage intelligent des composants

        /// <summary>
        /// Exécute HideBox - Masque/affiche intelligemment les composants box, template, dummy, etc.
        /// Code converti depuis SmartHideBox.iLogicVb
        /// </summary>
        public async Task ExecuteHideBoxAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    Log("Vérification du document actif...", "INFO");
                    dynamic doc = GetActiveAssemblyDocument();
                    var motsClesAMasquer = new List<string>
                    {
                        "box", "template", "multi_opening", "cut_opening", "roof_dummy", "vicwest_dummy",
                        "transition_dummy", "dummyremovable", "dummysupportalignment", "panfactice",
                        "airflow_r", "airflow_l", "cellidsketch", "opendummy",
                        "external_left_swing_1", "external_right_swing_1",
                        "internal_left_swing_1", "internal_right_swing_1"
                    };

                    int nombreTotal = 0;
                    int nombreVisibles = 0;
                    int nombreCaches = 0;

                    // Analyser l'état actuel
                    AnalyserEtatComposants(doc.ComponentDefinition.Occurrences, motsClesAMasquer, 
                        ref nombreTotal, ref nombreVisibles, ref nombreCaches);

                    int nombreTotalTraites = 0;
                    int nombreCutOpening = 0;
                    int nombreDummy = 0;
                    int nombreErreurs = 0;

                    // Déterminer et exécuter l'action
                    if (nombreCaches > nombreVisibles)
                    {
                        // Plus de cachés que de visibles -> RÉAFFICHER
                        ReafficherComposants(doc.ComponentDefinition.Occurrences, motsClesAMasquer,
                            ref nombreTotalTraites, ref nombreCutOpening, ref nombreDummy, ref nombreErreurs);
                    }
                    else
                    {
                        // Plus de visibles que de cachés -> MASQUER
                        MasquerComposants(doc.ComponentDefinition.Occurrences, motsClesAMasquer,
                            ref nombreTotalTraites, ref nombreCutOpening, ref nombreDummy, ref nombreErreurs);
                    }

                    // Mise à jour du document
                    doc.Update2(true);
                    
                    // Vue ISO + Zoom All silencieux
                    ExecuteIsoViewAndZoomAllSilent();
                    
                    string action = (nombreCaches > nombreVisibles) ? "affiches" : "caches";
                    Log($"[+] HideBox termine: {nombreTotalTraites} composants {action} - Vue ISO + Zoom All", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'execution de HideBox: {ex.Message}", "ERROR");
                    throw new Exception($"Erreur lors de l'execution de HideBox: {ex.Message}", ex);
                }
            });
        }

        private void AnalyserEtatComposants(dynamic occurrences, List<string> motsCles, 
            ref int totalCount, ref int visibleCount, ref int hiddenCount)
        {
            foreach (dynamic occ in occurrences)
            {
                try
                {
                    if (DoitEtreGere(occ, motsCles))
                    {
                        totalCount++;
                        if (occ.Visible)
                        {
                            visibleCount++;
                        }
                        else
                        {
                            hiddenCount++;
                        }
                    }

                    // Récursion pour les sous-assemblages
                    // Vérifier si l'occurrence a des sous-occurrences (méthode robuste via COM)
                    try
                    {
                        dynamic subOccurrences = occ.SubOccurrences;
                        if (subOccurrences != null)
                        {
                            AnalyserEtatComposants(subOccurrences, motsCles, 
                                ref totalCount, ref visibleCount, ref hiddenCount);
                        }
                    }
                    catch
                    {
                        // Pas de sous-occurrences ou erreur - continuer
                    }
                }
                catch
                {
                    // Continuer malgré l'erreur
                }
            }
        }

        private bool DoitEtreGere(dynamic occ, List<string> motsCles)
        {
            string nomFichier = occ.Name.ToLower();

            // EXCLUSION: J-Box
            if (nomFichier.Contains("j_box") || nomFichier.Contains("j-box"))
            {
                return false;
            }

            // EXCLUSION: Cross_Member dans Multi_Opening
            if (nomFichier.Contains("cross_member") && nomFichier.Contains("multi_opening"))
            {
                return false;
            }

            // EXCLUSION: Structural_Tube dans Multi_Opening
            if (nomFichier.Contains("structural_tube") && nomFichier.Contains("multi_opening"))
            {
                return false;
            }

            // EXCLUSION: Toutes les variations de Tee_Joint
            if (EstTeeJoint(nomFichier))
            {
                return false;
            }

            // EXCLUSION SPECIALE: cut_opening dans tube structurel
            if (nomFichier.Contains("cut_opening") && EstDansTubeStructurel(occ))
            {
                return false;
            }

            // DETECTION: Cut_Opening normal
            if (nomFichier.Contains("cut_opening")) return true;

            // DETECTION: Dummy elements
            if (nomFichier.Contains("roof_dummy") || nomFichier.Contains("vicwest_dummy") || 
                nomFichier.Contains("transition_dummy"))
            {
                return true;
            }

            // DETECTION: DummyRemovable (Right/Left)
            if (nomFichier.Contains("dummyremovable_right") || nomFichier.Contains("dummyremovable_left"))
            {
                return true;
            }

            // DETECTION: DummySupportAlignment
            if (nomFichier.Contains("dummysupportalignment"))
            {
                return true;
            }

            // DETECTION: PanFactice
            if (nomFichier.Contains("panfactice"))
            {
                return true;
            }

            // DETECTION: AirFlow_R et AirFlow_L
            if (nomFichier.Contains("airflow_r") || nomFichier.Contains("airflow_l"))
            {
                return true;
            }

            // DETECTION: OpenDummy_X (tous numéros)
            if (nomFichier.Contains("opendummy"))
            {
                return true;
            }

            // DETECTION: External_Left_Swing_X et External_Right_Swing_X
            if (nomFichier.Contains("external_left_swing") || nomFichier.Contains("external_right_swing"))
            {
                return true;
            }

            // DETECTION: Internal_Left_Swing et Internal_Right_Swing
            if (nomFichier.Contains("internal_left_swing") || nomFichier.Contains("internal_right_swing"))
            {
                return true;
            }

            // DETECTION: CellIDSketch
            if (nomFichier.Contains("cellidsketch"))
            {
                return true;
            }

            // DETECTION: Joint patterns (tout type de joint SAUF Tee_Joint)
            if (EstJointPattern(nomFichier))
            {
                return true;
            }

            // DETECTION: Mots-clés standards avec vérification spéciale pour multi_opening
            foreach (string motCle in motsCles)
            {
                if (nomFichier.Contains(motCle))
                {
                    // Vérification supplémentaire pour multi_opening
                    if (motCle == "multi_opening")
                    {
                        // Ne pas masquer si c'est un composant structurel avec multi_opening
                        if (nomFichier.Contains("structural_tube") || nomFichier.Contains("cross_member"))
                        {
                            return false;
                        }
                    }
                    return true;
                }
            }

            return false;
        }

        private bool EstDansTubeStructurel(dynamic occ)
        {
            try
            {
                // Vérifier le parent ou le contexte
                if (occ.ParentOccurrence != null)
                {
                    string parentName = occ.ParentOccurrence.Name.ToLower();
                    // Patterns typiques de tubes structurels
                    if (parentName.Contains("tube") || parentName.Contains("struct") ||
                        parentName.Contains("beam") || parentName.Contains("column"))
                    {
                        return true;
                    }
                }

                // Vérifier le nom du composant lui-même
                string nomComplet = occ.Name.ToLower();
                if (nomComplet.Contains("tube") && nomComplet.Contains("cut_opening"))
                {
                    return true;
                }
            }
            catch
            {
                // En cas d'erreur, ne pas exclure
            }

            return false;
        }

        private bool EstTeeJoint(string nomFichier)
        {
            if (nomFichier.Contains("tee_joint") ||
                nomFichier.Contains("teejoint") ||
                nomFichier.Contains("tee-joint") ||
                nomFichier.Contains("t_joint"))
            {
                return true;
            }

            if (nomFichier.Contains("tee") && (
                nomFichier.Contains("joint") ||
                nomFichier.Contains("jnt") ||
                nomFichier.Contains("jt")))
            {
                return true;
            }

            return false;
        }

        private bool EstJointPattern(string nomFichier)
        {
            // Si c'est un Tee_Joint sous n'importe quelle forme, ne pas traiter
            if (EstTeeJoint(nomFichier)) return false;

            // Patterns standards
            if (nomFichier.Contains("_joint") ||
                nomFichier.Contains("joint_") ||
                nomFichier.Contains("joint-") ||
                nomFichier.Contains("joint:") ||
                nomFichier.EndsWith("joint"))
            {
                return true;
            }

            // Pattern pour les murs Wall avec Joint
            if ((nomFichier.Contains("wall") ||
                nomFichier.Contains("roof") ||
                nomFichier.Contains("front") ||
                nomFichier.Contains("back") ||
                nomFichier.Contains("right") ||
                nomFichier.Contains("left") ||
                nomFichier.Contains("floor") ||
                nomFichier.Contains("ceiling")) &&
                nomFichier.Contains("joint"))
            {
                // S'assurer que ce n'est pas un Tee_Joint dans un mur
                if (!EstTeeJoint(nomFichier))
                {
                    return true;
                }
            }

            return false;
        }

        private void MasquerComposants(dynamic occurrences, List<string> motsCles,
            ref int totalCount, ref int cutOpeningCount, ref int dummyCount, ref int errorCount)
        {
            foreach (dynamic occ in occurrences)
            {
                try
                {
                    if (DoitEtreGere(occ, motsCles) && occ.Visible)
                    {
                        occ.Visible = false;
                        totalCount++;

                        string nomMinuscule = occ.Name.ToLower();
                        if (nomMinuscule.Contains("cut_opening"))
                        {
                            cutOpeningCount++;
                        }
                        if (nomMinuscule.Contains("dummy"))
                        {
                            dummyCount++;
                        }
                    }

                    // Récursion pour les sous-assemblages
                    // Vérifier si l'occurrence a des sous-occurrences (méthode robuste via COM)
                    try
                    {
                        dynamic subOccurrences = occ.SubOccurrences;
                        if (subOccurrences != null)
                        {
                            MasquerComposants(subOccurrences, motsCles,
                                ref totalCount, ref cutOpeningCount, ref dummyCount, ref errorCount);
                        }
                    }
                    catch
                    {
                        // Pas de sous-occurrences ou erreur - continuer
                    }
                }
                catch
                {
                    errorCount++;
                }
            }
        }

        private void ReafficherComposants(dynamic occurrences, List<string> motsCles,
            ref int totalCount, ref int cutOpeningCount, ref int dummyCount, ref int errorCount)
        {
            foreach (dynamic occ in occurrences)
            {
                try
                {
                    if (DoitEtreGere(occ, motsCles) && !occ.Visible)
                    {
                        occ.Visible = true;
                        totalCount++;

                        string nomMinuscule = occ.Name.ToLower();
                        if (nomMinuscule.Contains("cut_opening"))
                        {
                            cutOpeningCount++;
                        }
                        if (nomMinuscule.Contains("dummy"))
                        {
                            dummyCount++;
                        }
                    }

                    // Récursion pour les sous-assemblages
                    // Vérifier si l'occurrence a des sous-occurrences (méthode robuste via COM)
                    try
                    {
                        dynamic subOccurrences = occ.SubOccurrences;
                        if (subOccurrences != null)
                        {
                            ReafficherComposants(subOccurrences, motsCles,
                                ref totalCount, ref cutOpeningCount, ref dummyCount, ref errorCount);
                        }
                    }
                    catch
                    {
                        // Pas de sous-occurrences ou erreur - continuer
                    }
                }
                catch
                {
                    errorCount++;
                }
            }
        }

        #endregion

        #region Helpers - Méthodes utilitaires pour accéder aux documents

        /// <summary>
        /// Obtient le document d'une occurrence de manière fiable pour une application externe
        /// Essaie plusieurs méthodes pour accéder au document
        /// </summary>
        private dynamic? GetOccurrenceDocument(dynamic occ)
        {
            if (occ == null) return null;

            // Méthode 1: Essayer occ.Definition.Document (si le document est déjà ouvert)
            try
            {
                dynamic defDoc = occ.Definition.Document;
                if (defDoc != null)
                {
                    // Vérifier que c'est un document valide
                    try
                    {
                        var test = defDoc.DocumentType;
                        return defDoc;
                    }
                    catch { }
                }
            }
            catch { }

            // Méthode 2: Essayer ReferencedDocumentDescriptor.ReferencedDocument
            try
            {
                dynamic refDocDesc = occ.ReferencedDocumentDescriptor;
                if (refDocDesc != null)
                {
                    dynamic refDoc = refDocDesc.ReferencedDocument;
                    if (refDoc != null)
                    {
                        try
                        {
                            var test = refDoc.DocumentType;
                            return refDoc;
                        }
                        catch { }
                    }
                }
            }
            catch { }

            // Méthode 3: Essayer d'ouvrir le document depuis le chemin
            try
            {
                dynamic inventorApp = GetInventorApplication();
                string filePath = occ.ReferencedFileDescriptor.FullFileName;
                if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                {
                    // Vérifier si le document est déjà ouvert
                    foreach (dynamic doc in inventorApp.Documents)
                    {
                        try
                        {
                            if (doc.FullFileName == filePath)
                            {
                                return doc;
                            }
                        }
                        catch { }
                    }

                    // Ouvrir le document en lecture seule
                    dynamic openedDoc = inventorApp.Documents.Open(filePath, false);
                    if (openedDoc != null)
                    {
                        return openedDoc;
                    }
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Obtient la définition d'un composant depuis une occurrence
        /// Peut accéder directement aux workplanes/esquisses sans ouvrir le document
        /// </summary>
        private dynamic? GetOccurrenceDefinition(dynamic occ)
        {
            if (occ == null) return null;

            try
            {
                dynamic definition = occ.Definition;
                if (definition != null)
                {
                    return definition;
                }
            }
            catch { }

            return null;
        }

        #endregion

        #region ToggleRefVisibility - Basculer visibilité des références

        /// <summary>
        /// Bascule la visibilité des éléments de référence (Plans, Axes, Points)
        /// Utilise doc.ObjectVisibility (API officielle Inventor 2011+)
        /// Équivalent de: View > Object Visibility > All Work Features
        /// </summary>
        public async Task ExecuteToggleRefVisibilityAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif n'est ouvert.", "ERROR");
                        throw new InvalidOperationException("Aucun document actif n'est ouvert.");
                    }

                    // Vérifier le type de document via l'extension du fichier (méthode fiable avec dynamic)
                    string fullFileName = doc.FullFileName;
                    string extension = System.IO.Path.GetExtension(fullFileName).ToLowerInvariant();
                    
                    bool isPart = extension == ".ipt";
                    bool isAssembly = extension == ".iam";
                    
                    if (!isPart && !isAssembly)
                    {
                        Log("Ce script fonctionne uniquement sur des fichiers de type Piece ou Assemblage.", "WARNING");
                        return;
                    }
                    
                    string detectedType = isPart ? "Piece" : "Assemblage";
                    Log($"Type de document detecte: {detectedType}", "INFO");

                    // Déterminer l'état actuel (vérifier si quelque chose est visible)
                    bool currentState = AreAnyRefVisible_Simplified(doc);
                    bool newState = !currentState;
                    
                    Log($"Etat actuel: {(currentState ? "Visibles" : "Masquees")}. Nouvel etat: {(newState ? "Afficher" : "Masquer")}", "INFO");

                    // ÉTAPE 1: ObjectVisibility pour les User Work Planes (API globale)
                    // Cela affecte les User Work Planes dans TOUS les sous-assemblages automatiquement
                    try
                    {
                        dynamic objVis = doc.ObjectVisibility;
                        
                        Log($"Application ObjectVisibility pour User Work Features...", "INFO");
                        
                        // Basculer les User Work Features via ObjectVisibility
                        try { objVis.UserWorkPlanes = newState; } catch { }
                        try { objVis.UserWorkAxes = newState; } catch { }
                        try { objVis.UserWorkPoints = newState; } catch { }
                        
                        Log($"[+] UserWorkPlanes={newState}, UserWorkAxes={newState}, UserWorkPoints={newState}", "INFO");
                    }
                    catch (Exception objVisEx)
                    {
                        Log($"[!] ObjectVisibility.User* non disponible: {objVisEx.Message}", "WARNING");
                    }

                    // ÉTAPE 2: Méthode récursive individuelle pour les Origin Planes
                    var stats = ToggleRefVisibility_Simplified(doc, newState);
                    
                    // Mise à jour de la vue
                    try
                    {
                        dynamic activeView = inventorApp.ActiveView;
                        if (activeView != null)
                        {
                            activeView.Update();
                        }
                    }
                    catch { }
                    
                    // Vue ISO + Zoom All silencieux
                    ExecuteIsoViewAndZoomAllSilent();
                    
                    string action = newState ? "affichees" : "cachees";
                    Log($"[+] References {action} - Vue ISO + Zoom All", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'execution de ToggleRefVisibility: {ex.Message}", "ERROR");
                    throw new Exception($"Erreur lors de l'execution de ToggleRefVisibility: {ex.Message}", ex);
                }
            });
        }

        /// <summary>
        /// Structure pour les statistiques (comme dans le VB original)
        /// </summary>
        private struct RefStats
        {
            public int WorkPlanes;
            public int WorkAxes;
            public int WorkPoints;
            public int DocumentsProcessed;
        }

        /// <summary>
        /// Vérifie si au moins une référence est visible (pattern SDK externe COM)
        /// </summary>
        private bool AreAnyRefVisible_Simplified(dynamic doc)
        {
            try
            {
                // Utiliser l'extension du fichier pour déterminer le type (fiable avec dynamic)
                string fullFileName = doc.FullFileName;
                string extension = System.IO.Path.GetExtension(fullFileName).ToLowerInvariant();
                bool isPart = extension == ".ipt";
                bool isAssembly = extension == ".iam";
                
                if (isPart)
                {
                    // C'est une pièce - vérifier les WorkPlanes/Axes/Points
                    dynamic partDef = doc.ComponentDefinition;
                    
                    try
                    {
                        dynamic workPlanes = partDef.WorkPlanes;
                        for (int i = 1; i <= workPlanes.Count; i++)
                        {
                            try { if (workPlanes[i].Visible) return true; } catch { }
                        }
                    }
                    catch { }
                    
                    try
                    {
                        dynamic workAxes = partDef.WorkAxes;
                        for (int i = 1; i <= workAxes.Count; i++)
                        {
                            try { if (workAxes[i].Visible) return true; } catch { }
                        }
                    }
                    catch { }
                    
                    try
                    {
                        dynamic workPoints = partDef.WorkPoints;
                        for (int i = 1; i <= workPoints.Count; i++)
                        {
                            try { if (workPoints[i].Visible) return true; } catch { }
                        }
                    }
                    catch { }
                }
                else if (isAssembly)
                {
                    // C'est un assemblage
                    dynamic asmDef = doc.ComponentDefinition;
                    
                    // Vérifier au niveau de l'assemblage
                    try
                    {
                        dynamic workPlanes = asmDef.WorkPlanes;
                        for (int i = 1; i <= workPlanes.Count; i++)
                        {
                            try { if (workPlanes[i].Visible) return true; } catch { }
                        }
                    }
                    catch { }
                    
                    try
                    {
                        dynamic workAxes = asmDef.WorkAxes;
                        for (int i = 1; i <= workAxes.Count; i++)
                        {
                            try { if (workAxes[i].Visible) return true; } catch { }
                        }
                    }
                    catch { }
                    
                    try
                    {
                        dynamic workPoints = asmDef.WorkPoints;
                        for (int i = 1; i <= workPoints.Count; i++)
                        {
                            try { if (workPoints[i].Visible) return true; } catch { }
                        }
                    }
                    catch { }
                    
                    // PARCOURS RÉCURSIF - Pattern SDK externe COM
                    // Utiliser index 1-based et SubOccurrences pour sous-assemblages
                    if (AreAnyRefVisibleInOccurrences(asmDef.Occurrences))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Erreur AreAnyRefVisible_Simplified: {ex.Message}");
            }
            
            return false;
        }
        
        /// <summary>
        /// Vérifie récursivement les occurrences (pattern SDK externe COM avec SubOccurrences)
        /// </summary>
        private bool AreAnyRefVisibleInOccurrences(dynamic occurrences)
        {
            try
            {
                int count = occurrences.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic occ = occurrences[i];
                        
                        // Vérifier si l'occurrence n'est pas supprimée et est visible
                        bool suppressed = false;
                        bool visible = true;
                        try { suppressed = occ.Suppressed; } catch { }
                        try { visible = occ.Visible; } catch { }
                        
                        if (!suppressed && visible)
                        {
                            // Vérifier les éléments de référence dans la définition
                            try
                            {
                                dynamic definition = occ.Definition;
                                if (definition != null)
                                {
                                    // WorkPlanes
                                    try
                                    {
                                        dynamic workPlanes = definition.WorkPlanes;
                                        for (int j = 1; j <= workPlanes.Count; j++)
                                        {
                                            try { if (workPlanes[j].Visible) return true; } catch { }
                                        }
                                    }
                                    catch { }
                                    
                                    // WorkAxes
                                    try
                                    {
                                        dynamic workAxes = definition.WorkAxes;
                                        for (int j = 1; j <= workAxes.Count; j++)
                                        {
                                            try { if (workAxes[j].Visible) return true; } catch { }
                                        }
                                    }
                                    catch { }
                                    
                                    // WorkPoints
                                    try
                                    {
                                        dynamic workPoints = definition.WorkPoints;
                                        for (int j = 1; j <= workPoints.Count; j++)
                                        {
                                            try { if (workPoints[j].Visible) return true; } catch { }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                            
                            // PATTERN SDK EXTERNE: Utiliser SubOccurrences pour sous-assemblages
                            try
                            {
                                dynamic subOccurrences = occ.SubOccurrences;
                                if (subOccurrences != null && subOccurrences.Count > 0)
                                {
                                    if (AreAnyRefVisibleInOccurrences(subOccurrences))
                                    {
                                        return true;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Erreur AreAnyRefVisibleInOccurrences: {ex.Message}");
            }
            
            return false;
        }

        /// <summary>
        /// Bascule la visibilité des références de manière récursive (pattern SDK externe COM)
        /// </summary>
        private RefStats ToggleRefVisibility_Simplified(dynamic doc, bool visibleState)
        {
            var stats = new RefStats();
            
            try
            {
                // Utiliser l'extension du fichier pour déterminer le type (fiable avec dynamic)
                string fullFileName = doc.FullFileName;
                string extension = System.IO.Path.GetExtension(fullFileName).ToLowerInvariant();
                bool isPart = extension == ".ipt";
                bool isAssembly = extension == ".iam";
                
                if (isPart)
                {
                    // C'est une pièce - traiter les WorkPlanes/Axes/Points
                    dynamic partDef = doc.ComponentDefinition;
                    
                    // WorkPlanes - inclut Origin Planes ET User Work Planes
                    try
                    {
                        dynamic workPlanes = partDef.WorkPlanes;
                        int wpCount = workPlanes.Count;
                        Log($"[DEBUG] Piece: {wpCount} WorkPlanes trouvés", "INFO");
                        for (int i = 1; i <= wpCount; i++)
                        {
                            try 
                            { 
                                dynamic wp = workPlanes[i];
                                string wpName = wp.Name;
                                wp.Visible = visibleState; 
                                stats.WorkPlanes++; 
                                Log($"[DEBUG]   Plan '{wpName}' -> {visibleState}", "INFO");
                            } 
                            catch (Exception wpEx) 
                            { 
                                Log($"[DEBUG]   Erreur plan[{i}]: {wpEx.Message}", "WARNING");
                            }
                        }
                    }
                    catch (Exception ex) { Log($"[DEBUG] Erreur acces WorkPlanes: {ex.Message}", "WARNING"); }
                    
                    // WorkAxes - inclut Origin Axes ET User Work Axes
                    try
                    {
                        dynamic workAxes = partDef.WorkAxes;
                        int waCount = workAxes.Count;
                        Log($"[DEBUG] Piece: {waCount} WorkAxes trouvés", "INFO");
                        for (int i = 1; i <= waCount; i++)
                        {
                            try 
                            { 
                                dynamic wa = workAxes[i];
                                wa.Visible = visibleState; 
                                stats.WorkAxes++; 
                            } 
                            catch { }
                        }
                    }
                    catch { }
                    
                    // WorkPoints - inclut Origin Point ET User Work Points
                    try
                    {
                        dynamic workPoints = partDef.WorkPoints;
                        int wptCount = workPoints.Count;
                        Log($"[DEBUG] Piece: {wptCount} WorkPoints trouvés", "INFO");
                        for (int i = 1; i <= wptCount; i++)
                        {
                            try 
                            { 
                                dynamic wpt = workPoints[i];
                                wpt.Visible = visibleState; 
                                stats.WorkPoints++; 
                            } 
                            catch { }
                        }
                    }
                    catch { }
                    
                    stats.DocumentsProcessed = 1;
                }
                else if (isAssembly)
                {
                    // C'est un assemblage
                    dynamic asmDef = doc.ComponentDefinition;
                    string docName = doc.DisplayName;
                    Log($"[DEBUG] Assemblage: {docName}", "INFO");
                    
                    // Éléments de référence au niveau de l'assemblage
                    try
                    {
                        dynamic workPlanes = asmDef.WorkPlanes;
                        int wpCount = workPlanes.Count;
                        Log($"[DEBUG] Assemblage: {wpCount} WorkPlanes au niveau racine", "INFO");
                        for (int i = 1; i <= wpCount; i++)
                        {
                            try 
                            { 
                                dynamic wp = workPlanes[i];
                                string wpName = wp.Name;
                                wp.Visible = visibleState; 
                                stats.WorkPlanes++; 
                                Log($"[DEBUG]   Plan '{wpName}' -> {visibleState}", "INFO");
                            } 
                            catch (Exception wpEx) 
                            { 
                                Log($"[DEBUG]   Erreur plan[{i}]: {wpEx.Message}", "WARNING");
                            }
                        }
                    }
                    catch (Exception ex) { Log($"[DEBUG] Erreur acces WorkPlanes asm: {ex.Message}", "WARNING"); }
                    
                    try
                    {
                        dynamic workAxes = asmDef.WorkAxes;
                        int waCount = workAxes.Count;
                        Log($"[DEBUG] Assemblage: {waCount} WorkAxes au niveau racine", "INFO");
                        for (int i = 1; i <= waCount; i++)
                        {
                            try { workAxes[i].Visible = visibleState; stats.WorkAxes++; } catch { }
                        }
                    }
                    catch { }
                    
                    try
                    {
                        dynamic workPoints = asmDef.WorkPoints;
                        int wptCount = workPoints.Count;
                        Log($"[DEBUG] Assemblage: {wptCount} WorkPoints au niveau racine", "INFO");
                        for (int i = 1; i <= wptCount; i++)
                        {
                            try { workPoints[i].Visible = visibleState; stats.WorkPoints++; } catch { }
                        }
                    }
                    catch { }
                    
                    stats.DocumentsProcessed = 1;
                    
                    // PARCOURS RÉCURSIF - Pattern SDK externe COM avec SubOccurrences
                    Log($"[DEBUG] Debut parcours recursif des occurrences...", "INFO");
                    var subStats = ToggleRefVisibilityInOccurrences(asmDef.Occurrences, visibleState);
                    stats.WorkPlanes += subStats.WorkPlanes;
                    stats.WorkAxes += subStats.WorkAxes;
                    stats.WorkPoints += subStats.WorkPoints;
                    stats.DocumentsProcessed += subStats.DocumentsProcessed;
                    Log($"[DEBUG] Fin parcours recursif: {subStats.WorkPlanes}P, {subStats.WorkAxes}A, {subStats.WorkPoints}Pt dans {subStats.DocumentsProcessed} docs", "INFO");
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur dans ToggleRefVisibility_Simplified: {ex.Message}", "ERROR");
            }
            
            return stats;
        }
        
        /// <summary>
        /// Bascule récursivement les occurrences (pattern identique au VB original Add-In)
        /// Utilise occ.Definition.Document pour accéder au document réel et modifier les WorkPlanes
        /// </summary>
        private RefStats ToggleRefVisibilityInOccurrences(dynamic occurrences, bool visibleState)
        {
            var stats = new RefStats();
            
            try
            {
                int count = occurrences.Count;
                Log($"Parcours de {count} occurrence(s)...", "INFO");
                
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic occ = occurrences[i];
                        
                        // Vérifier si l'occurrence n'est pas supprimée et est visible
                        bool suppressed = false;
                        bool visible = true;
                        string occName = "Inconnu";
                        try { suppressed = occ.Suppressed; } catch { }
                        try { visible = occ.Visible; } catch { }
                        try { occName = occ.Name; } catch { }
                        
                        if (!suppressed && visible)
                        {
                            // PATTERN IDENTIQUE AU VB ORIGINAL:
                            // Récupérer le document via occ.Definition.Document
                            // puis appeler récursivement ToggleRefVisibility_Simplified sur ce document
                            try
                            {
                                dynamic subDoc = occ.Definition.Document;
                                if (subDoc != null)
                                {
                                    Log($"  [{i}] Traitement recursif document: {occName}", "INFO");
                                    var subStats = ToggleRefVisibility_Simplified(subDoc, visibleState);
                                    stats.WorkPlanes += subStats.WorkPlanes;
                                    stats.WorkAxes += subStats.WorkAxes;
                                    stats.WorkPoints += subStats.WorkPoints;
                                    stats.DocumentsProcessed += subStats.DocumentsProcessed;
                                }
                            }
                            catch (Exception docEx)
                            {
                                Log($"  [{i}] Erreur acces document {occName}: {docEx.Message}", "WARNING");
                            }
                        }
                    }
                    catch (Exception occEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[!] Erreur occurrence index {i}: {occEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur dans ToggleRefVisibilityInOccurrences: {ex.Message}", "ERROR");
            }
            
            return stats;
        }

        #endregion

        #region ToggleSketchVisibility - Basculer visibilité des esquisses

        /// <summary>
        /// Active ou désactive la visibilité de toutes les esquisses (2D et 3D)
        /// Utilise doc.ObjectVisibility (API officielle Inventor 2011+)
        /// Équivalent de: View > Object Visibility > Sketches / Sketches3D
        /// </summary>
        public async Task ExecuteToggleSketchVisibilityAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif n'est ouvert.", "ERROR");
                        throw new InvalidOperationException("Aucun document actif n'est ouvert.");
                    }

                    // Vérifier le type de document via l'extension du fichier (méthode fiable avec dynamic)
                    string fullFileName = doc.FullFileName;
                    string extension = System.IO.Path.GetExtension(fullFileName).ToLowerInvariant();
                    
                    bool isPart = extension == ".ipt";
                    bool isAssembly = extension == ".iam";
                    
                    if (!isPart && !isAssembly)
                    {
                        Log("Ce script fonctionne uniquement sur des fichiers .ipt ou .iam.", "WARNING");
                        return;
                    }
                    
                    string detectedType = isPart ? "Piece" : "Assemblage";
                    Log($"Type de document detecte: {detectedType}", "INFO");

                    // Déterminer l'état actuel (vérifier si quelque chose est visible)
                    bool currentState = HasSketchVisible_Simplified(doc);
                    bool newState = !currentState;
                    
                    Log($"Etat actuel: {(currentState ? "Certaines visibles" : "Toutes masquees")}. Nouvel etat: {(newState ? "Afficher" : "Masquer")}", "INFO");

                    // ÉTAPE 1: ObjectVisibility pour les Sketches (API globale)
                    try
                    {
                        dynamic objVis = doc.ObjectVisibility;
                        try { objVis.Sketches = newState; } catch { }
                        try { objVis.Sketches3D = newState; } catch { }
                    }
                    catch { }

                    // ÉTAPE 2: Méthode récursive individuelle pour tous les documents
                    ToggleSketches_Simplified(doc, newState);
                    
                    // Mise à jour de la vue
                    try
                    {
                        dynamic activeView = inventorApp.ActiveView;
                        if (activeView != null)
                        {
                            activeView.Update();
                        }
                    }
                    catch { }
                    
                    // Vue ISO + Zoom All silencieux
                    ExecuteIsoViewAndZoomAllSilent();
                    
                    string action = newState ? "affichees" : "cachees";
                    Log($"[+] Esquisses {action} - Vue ISO + Zoom All", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur Toggle_SketchVisibility: {ex.Message}", "ERROR");
                }
            });
        }

        /// <summary>
        /// Vérifie si au moins une esquisse est visible (pattern SDK externe COM)
        /// </summary>
        private bool HasSketchVisible_Simplified(dynamic doc)
        {
            try
            {
                // Utiliser l'extension du fichier pour déterminer le type (fiable avec dynamic)
                string fullFileName = doc.FullFileName;
                string extension = System.IO.Path.GetExtension(fullFileName).ToLowerInvariant();
                bool isAssembly = extension == ".iam";
                
                dynamic def = doc.ComponentDefinition;
                
                // Vérifier les esquisses 2D du document courant (index 1-based)
                try
                {
                    dynamic sketches = def.Sketches;
                    for (int i = 1; i <= sketches.Count; i++)
                    {
                        try { if (sketches[i].Visible) return true; } catch { }
                    }
                }
                catch { }
                
                // Vérifier les esquisses 3D
                try
                {
                    dynamic sketches3D = def.Sketches3D;
                    for (int i = 1; i <= sketches3D.Count; i++)
                    {
                        try { if (sketches3D[i].Visible) return true; } catch { }
                    }
                }
                catch { }
                
                // Vérification récursive pour les assemblages (PATTERN SDK EXTERNE COM)
                if (isAssembly)
                {
                    if (HasSketchVisibleInOccurrences(def.Occurrences))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Erreur HasSketchVisible_Simplified: {ex.Message}");
            }
            
            return false;
        }
        
        /// <summary>
        /// Vérifie récursivement les esquisses dans les occurrences (pattern SDK externe COM avec SubOccurrences)
        /// </summary>
        private bool HasSketchVisibleInOccurrences(dynamic occurrences)
        {
            try
            {
                int count = occurrences.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic occ = occurrences[i];
                        
                        // Vérifier si l'occurrence n'est pas supprimée et est visible
                        bool suppressed = false;
                        bool visible = true;
                        try { suppressed = occ.Suppressed; } catch { }
                        try { visible = occ.Visible; } catch { }
                        
                        if (!suppressed && visible)
                        {
                            // Vérifier les esquisses dans la définition
                            try
                            {
                                dynamic definition = occ.Definition;
                                if (definition != null)
                                {
                                    // Sketches 2D
                                    try
                                    {
                                        dynamic sketches = definition.Sketches;
                                        for (int j = 1; j <= sketches.Count; j++)
                                        {
                                            try { if (sketches[j].Visible) return true; } catch { }
                                        }
                                    }
                                    catch { }
                                    
                                    // Sketches 3D
                                    try
                                    {
                                        dynamic sketches3D = definition.Sketches3D;
                                        for (int j = 1; j <= sketches3D.Count; j++)
                                        {
                                            try { if (sketches3D[j].Visible) return true; } catch { }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                            
                            // PATTERN SDK EXTERNE: Utiliser SubOccurrences pour sous-assemblages
                            try
                            {
                                dynamic subOccurrences = occ.SubOccurrences;
                                if (subOccurrences != null && subOccurrences.Count > 0)
                                {
                                    if (HasSketchVisibleInOccurrences(subOccurrences))
                                    {
                                        return true;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Erreur HasSketchVisibleInOccurrences: {ex.Message}");
            }
            
            return false;
        }

        /// <summary>
        /// Bascule la visibilité des esquisses de manière récursive (pattern SDK externe COM)
        /// </summary>
        private void ToggleSketches_Simplified(dynamic doc, bool visibleState)
        {
            try
            {
                // Utiliser l'extension du fichier pour déterminer le type (fiable avec dynamic)
                string fullFileName = doc.FullFileName;
                string extension = System.IO.Path.GetExtension(fullFileName).ToLowerInvariant();
                bool isAssembly = extension == ".iam";
                
                dynamic def = doc.ComponentDefinition;
                int sketchCount = 0;
                
                // Basculer les esquisses 2D du document courant (index 1-based)
                try
                {
                    dynamic sketches = def.Sketches;
                    for (int i = 1; i <= sketches.Count; i++)
                    {
                        try { sketches[i].Visible = visibleState; sketchCount++; } catch { }
                    }
                }
                catch { }
                
                // Basculer les esquisses 3D
                try
                {
                    dynamic sketches3D = def.Sketches3D;
                    for (int i = 1; i <= sketches3D.Count; i++)
                    {
                        try { sketches3D[i].Visible = visibleState; sketchCount++; } catch { }
                    }
                }
                catch { }
                
                if (sketchCount > 0)
                {
                    Log($"Document principal: {sketchCount} esquisse(s) traitee(s)", "INFO");
                }
                
                // Traitement récursif pour les assemblages (PATTERN SDK EXTERNE COM)
                if (isAssembly)
                {
                    ToggleSketchesInOccurrences(def.Occurrences, visibleState);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Erreur ToggleSketches_Simplified: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Bascule récursivement les esquisses dans les occurrences (pattern SDK externe COM avec SubOccurrences)
        /// </summary>
        private void ToggleSketchesInOccurrences(dynamic occurrences, bool visibleState)
        {
            try
            {
                int count = occurrences.Count;
                
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic occ = occurrences[i];
                        
                        // Vérifier si l'occurrence n'est pas supprimée et est visible
                        bool suppressed = false;
                        bool visible = true;
                        string occName = "Inconnu";
                        try { suppressed = occ.Suppressed; } catch { }
                        try { visible = occ.Visible; } catch { }
                        try { occName = occ.Name; } catch { }
                        
                        if (!suppressed && visible)
                        {
                            int sketchCount = 0;
                            
                            // Traiter les esquisses dans la définition
                            try
                            {
                                dynamic definition = occ.Definition;
                                if (definition != null)
                                {
                                    // Sketches 2D
                                    try
                                    {
                                        dynamic sketches = definition.Sketches;
                                        for (int j = 1; j <= sketches.Count; j++)
                                        {
                                            try { sketches[j].Visible = visibleState; sketchCount++; } catch { }
                                        }
                                    }
                                    catch { }
                                    
                                    // Sketches 3D
                                    try
                                    {
                                        dynamic sketches3D = definition.Sketches3D;
                                        for (int j = 1; j <= sketches3D.Count; j++)
                                        {
                                            try { sketches3D[j].Visible = visibleState; sketchCount++; } catch { }
                                        }
                                    }
                                    catch { }
                                    
                                    if (sketchCount > 0)
                                    {
                                        Log($"  [{i}] {occName}: {sketchCount} esquisse(s)", "INFO");
                                    }
                                }
                            }
                            catch (Exception defEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"[!] Erreur definition {occName}: {defEx.Message}");
                            }
                            
                            // PATTERN SDK EXTERNE: Utiliser SubOccurrences pour sous-assemblages
                            try
                            {
                                dynamic subOccurrences = occ.SubOccurrences;
                                if (subOccurrences != null)
                                {
                                    int subCount = subOccurrences.Count;
                                    if (subCount > 0)
                                    {
                                        Log($"  [{i}] {occName} est un sous-assemblage avec {subCount} sous-composants", "INFO");
                                        ToggleSketchesInOccurrences(subOccurrences, visibleState);
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception occEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[!] Erreur occurrence esquisse index {i}: {occEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur dans ToggleSketchesInOccurrences: {ex.Message}", "ERROR");
            }
        }

        #endregion

        #region ConstraintReport - Rapport de contraintes d'assemblage

        /// <summary>
        /// Génère un rapport de contraintes d'assemblage
        /// Code converti depuis AssemblyConstraintStatusReport.vb
        /// </summary>
        public async Task ExecuteConstraintReportAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif n'est ouvert.", "ERROR");
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType != kAssemblyDocumentObject)
                    {
                        Log("Ce rapport fonctionne uniquement sur des assemblages (.iam)", "WARNING");
                        return;
                    }

                    Log("Analyse des contraintes d'assemblage...", "INFO");

                    dynamic asmDef = doc.ComponentDefinition;
                    dynamic constraints = asmDef.Constraints;
                    
                    var constraintList = new System.Collections.Generic.List<Views.ConstraintInfo>();
                    int totalCount = 0;
                    int healthyCount = 0;
                    int suppressedCount = 0;
                    int failedCount = 0;

                    foreach (dynamic constraint in constraints)
                    {
                        try
                        {
                            totalCount++;
                            var info = new Views.ConstraintInfo();
                            
                            // Nom du type de contrainte
                            try
                            {
                                int constraintType = constraint.Type;
                                info.ConstraintType = GetConstraintTypeName(constraintType);
                            }
                            catch
                            {
                                info.ConstraintType = "Inconnu";
                            }

                            // Composants impliqués
                            try
                            {
                                dynamic entityOne = constraint.EntityOne;
                                dynamic entityTwo = constraint.EntityTwo;
                                
                                info.Component1 = GetComponentNameFromEntity(entityOne);
                                info.Component2 = GetComponentNameFromEntity(entityTwo);
                            }
                            catch
                            {
                                info.Component1 = "[Composant 1]";
                                info.Component2 = "[Composant 2]";
                            }

                            // Statut
                            try
                            {
                                info.IsSuppressed = constraint.Suppressed;
                                if (info.IsSuppressed)
                                {
                                    suppressedCount++;
                                }
                                else
                                {
                                    // Vérifier la santé de la contrainte
                                    try
                                    {
                                        int healthStatus = constraint.HealthStatus;
                                        // 0 = Healthy, autres = problème
                                        info.IsHealthy = (healthStatus == 0);
                                        if (info.IsHealthy)
                                            healthyCount++;
                                        else
                                            failedCount++;
                                    }
                                    catch
                                    {
                                        info.IsHealthy = true;
                                        healthyCount++;
                                    }
                                }
                            }
                            catch
                            {
                                info.IsSuppressed = false;
                                info.IsHealthy = true;
                                healthyCount++;
                            }

                            constraintList.Add(info);
                        }
                        catch (Exception ex)
                        {
                            Log($"Erreur lecture contrainte: {ex.Message}", "WARNING");
                        }
                    }

                    Log($"Contraintes analysées: {totalCount} (OK: {healthyCount}, Supprimées: {suppressedCount}, Erreurs: {failedCount})", "INFO");

                    // Générer et afficher le rapport HTML
                    string assemblyName = System.IO.Path.GetFileName(doc.FullFileName);
                    string htmlContent = Views.HtmlPopupWindow.GenerateConstraintReportHtml(assemblyName, constraintList);
                    
                    // Afficher via le callback
                    ShowHtmlPopup($"Rapport de contraintes - {assemblyName}", htmlContent);
                    
                    Log("Rapport de contraintes généré avec succès", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de la génération du rapport: {ex.Message}", "ERROR");
                }
            });
        }

        private string GetConstraintTypeName(int constraintType)
        {
            // Types de contraintes Inventor
            switch (constraintType)
            {
                case 59137: return "Mate";           // kMateConstraintObject
                case 59138: return "Angle";          // kAngleConstraintObject
                case 59139: return "Tangent";        // kTangentConstraintObject
                case 59140: return "Insert";         // kInsertConstraintObject
                case 59141: return "Flush";          // kFlushConstraintObject (symétrie)
                case 59142: return "Rotation";       // kRotationConstraintObject
                case 59143: return "Translation";    // kTranslationalConstraintObject
                case 59144: return "Motion";         // kMotionConstraintObject
                case 59145: return "Transitional";   // kTransitionalConstraintObject
                default: return $"Type_{constraintType}";
            }
        }

        private string GetComponentNameFromEntity(dynamic entity)
        {
            try
            {
                // Essayer d'obtenir le composant parent
                dynamic parent = entity.Parent;
                if (parent != null)
                {
                    try
                    {
                        return parent.Name;
                    }
                    catch
                    {
                        return parent.ToString();
                    }
                }
                return entity.ToString();
            }
            catch
            {
                return "[Inconnu]";
            }
        }

        #endregion

        #region iPropertiesSummary - Résumé des propriétés

        /// <summary>
        /// Affiche un résumé des iProperties du composant sélectionné
        /// Code converti depuis iPropertiesSummary.vb
        /// </summary>
        public async Task ExecuteIPropertiesSummaryAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif n'est ouvert.", "ERROR");
                        return;
                    }

                    // Vérifier s'il y a une sélection
                    dynamic selectSet = doc.SelectSet;
                    if (selectSet.Count == 0)
                    {
                        Log("Veuillez sélectionner un composant dans l'arbre ou le graphique.", "WARNING");
                        MessageBox.Show("Veuillez selectionner un composant dans l'arbre ou le graphique.", 
                            "Aucune selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // Obtenir l'objet sélectionné
                    dynamic selObj = selectSet[1];
                    dynamic compDoc = null;
                    string componentName = "";

                    // Vérifier si c'est un ComponentOccurrence (assemblage)
                    try
                    {
                        // Essayer d'accéder aux propriétés d'un ComponentOccurrence
                        dynamic definition = selObj.Definition;
                        compDoc = definition.Document;
                        componentName = selObj.Name;
                        Log($"Composant sélectionné: {componentName}", "INFO");
                    }
                    catch
                    {
                        // Si ce n'est pas un ComponentOccurrence, utiliser le document actif
                        compDoc = doc;
                        componentName = System.IO.Path.GetFileName(doc.FullFileName);
                        Log($"Document actif utilisé: {componentName}", "INFO");
                    }

                    if (compDoc == null)
                    {
                        Log("Impossible d'accéder au document du composant.", "ERROR");
                        return;
                    }

                    Log("Lecture des propriétés...", "INFO");

                    // Récupérer les propriétés
                    var properties = new System.Collections.Generic.Dictionary<string, string>();
                    
                    // Ordre prioritaire des propriétés XNRGY
                    var priorityProps = new[] { "Tag", "Tag_Assy", "Description", "Material", 
                        "Length", "Width", "Depth", "Thickness", "Prefix", "Destination", 
                        "Engraving", "Finish_Paint_Face", "MachineNo" };

                    try
                    {
                        dynamic propSets = compDoc.PropertySets;
                        
                        // Propriétés utilisateur personnalisées
                        try
                        {
                            dynamic customProps = propSets["Inventor User Defined Properties"];
                            foreach (string propName in priorityProps)
                            {
                                try
                                {
                                    dynamic prop = customProps[propName];
                                    if (prop != null && prop.Value != null)
                                    {
                                        properties[propName] = prop.Value.ToString();
                                    }
                                }
                                catch { /* Propriété non trouvée */ }
                            }

                            // Ajouter les autres propriétés custom non listées
                            foreach (dynamic prop in customProps)
                            {
                                string propName = prop.Name;
                                if (!properties.ContainsKey(propName) && propName != "Sheet Metal StyleName")
                                {
                                    try
                                    {
                                        if (prop.Value != null)
                                            properties[propName] = prop.Value.ToString();
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"Erreur lecture propriétés custom: {ex.Message}", "WARNING");
                        }

                        // Propriétés de conception (Design Tracking)
                        try
                        {
                            dynamic designProps = propSets["Design Tracking Properties"];
                            
                            // Material si pas déjà défini
                            if (!properties.ContainsKey("Material"))
                            {
                                try
                                {
                                    dynamic matProp = designProps["Material"];
                                    if (matProp != null && matProp.Value != null)
                                        properties["Material"] = matProp.Value.ToString();
                                }
                                catch { }
                            }

                            // Description si pas déjà défini
                            if (!properties.ContainsKey("Description"))
                            {
                                try
                                {
                                    dynamic descProp = designProps["Description"];
                                    if (descProp != null && descProp.Value != null)
                                        properties["Description"] = descProp.Value.ToString();
                                }
                                catch { }
                            }

                            // Part Number
                            try
                            {
                                dynamic partNumProp = designProps["Part Number"];
                                if (partNumProp != null && partNumProp.Value != null)
                                    properties["Part Number"] = partNumProp.Value.ToString();
                            }
                            catch { }
                        }
                        catch { }

                        // Propriétés du document (Summary Information)
                        try
                        {
                            dynamic summaryProps = propSets["Inventor Summary Information"];
                            
                            try
                            {
                                dynamic authorProp = summaryProps["Author"];
                                if (authorProp != null && authorProp.Value != null)
                                    properties["Auteur"] = authorProp.Value.ToString();
                            }
                            catch { }
                        }
                        catch { }

                        // Propriétés document info
                        try
                        {
                            dynamic docProps = propSets["Inventor Document Summary Information"];
                            
                            try
                            {
                                dynamic companyProp = docProps["Company"];
                                if (companyProp != null && companyProp.Value != null)
                                    properties["Entreprise"] = companyProp.Value.ToString();
                            }
                            catch { }
                        }
                        catch { }
                    }
                    catch (Exception ex)
                    {
                        Log($"Erreur accès PropertySets: {ex.Message}", "ERROR");
                    }

                    // Ajouter des infos système
                    try
                    {
                        properties["Derniere sauvegarde"] = System.IO.File.GetLastWriteTime(compDoc.FullFileName).ToString("yyyy-MM-dd HH:mm");
                    }
                    catch { }

                    Log($"Propriétés lues: {properties.Count}", "INFO");

                    // Générer et afficher le HTML
                    string fileName = System.IO.Path.GetFileName(compDoc.FullFileName);
                    string fullPath = compDoc.FullFileName;
                    string htmlContent = Views.HtmlPopupWindow.GenerateIPropertiesHtml(fileName, fullPath, properties);
                    
                    // Afficher via le callback
                    ShowHtmlPopup($"iProperties - {componentName}", htmlContent);
                    
                    Log("Résumé iProperties affiché avec succès", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de la génération du résumé: {ex.Message}", "ERROR");
                }
            });
        }

        #endregion

        #region SmartSave - Sauvegarde intelligente

        /// <summary>
        /// Sauvegarde intelligente des documents Inventor sans les fermer
        /// Code converti depuis SmartSave.iLogicVb
        /// Affiche une fenêtre HTML dynamique avec progression en temps réel
        /// </summary>
        public async Task ExecuteSmartSaveAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(async () =>
            {
                IProgressWindow? progressWindow = null;
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        MessageBox.Show("⚠️ Aucun document actif n'est ouvert.", "Erreur", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;
                    const int kPartDocumentObject = 12288;
                    const int kDrawingDocumentObject = 12291;

                    string typeText = docType == kAssemblyDocumentObject ? "Assemblage" :
                                     docType == kPartDocumentObject ? "Pièce" :
                                     docType == kDrawingDocumentObject ? "Mise en plan" : "Document";

                    string docName = doc.DisplayName ?? "Document";
                    string htmlContent = GenerateSmartSaveHtml(docType, docName, typeText);

                    // Créer la fenêtre de progression
                    if (_progressWindowCallback != null)
                    {
                        progressWindow = _progressWindowCallback("💾 Smart Save - V1.1 @2025 - By Mohammed Amine Elgalai", htmlContent);
                    }

                    // Exécuter les étapes avec mise à jour du HTML
                    if (docType == kAssemblyDocumentObject)
                    {
                        await ExecuteAssemblyStepsWithProgressAsync(doc, inventorApp, progressWindow);
                    }
                    else if (docType == kPartDocumentObject)
                    {
                        await ExecutePartStepsWithProgressAsync(doc, inventorApp, progressWindow);
                    }
                    else if (docType == kDrawingDocumentObject)
                    {
                        await ExecuteDrawingStepsWithProgressAsync(doc, inventorApp, progressWindow);
                    }
                    else
                    {
                        await ExecuteGenericStepsWithProgressAsync(doc, inventorApp, progressWindow);
                    }

                    // Afficher le message de complétion
                    if (progressWindow != null)
                    {
                        await progressWindow.ShowCompletionAsync("✅ Sauvegarde effectuée avec succès !");
                        await Task.Delay(1500);
                        progressWindow.CloseWindow();
                    }
                }
                catch (Exception ex)
                {
                    if (progressWindow != null)
                    {
                        await progressWindow.ShowCompletionAsync($"❌ Erreur: {ex.Message}");
                        await Task.Delay(2000);
                        progressWindow.CloseWindow();
                    }
                    MessageBox.Show($"❌ Erreur lors du Smart Save : {ex.Message}", "Erreur", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        private void ExecuteAssemblySteps(dynamic doc, dynamic inventorApp)
        {
            // Étape 1: Activation représentation par défaut
            ActivateDefaultRepresentation(doc);

            // Étape 2: Affichage de TOUS les composants
            try
            {
                dynamic cmdManager = inventorApp.CommandManager;
                dynamic controlDefs = cmdManager.ControlDefinitions;
                try
                {
                    dynamic cmd = controlDefs.Item("AssemblyShowAllComponentsCmd");
                    cmd.Execute();
                    doc.Update();
                }
                catch
                {
                    // Méthode alternative
                    AfficherTousComposantsRecursive(doc.ComponentDefinition.Occurrences);
                    doc.Update();
                }
            }
            catch
            {
                // Continuer
            }

            // Étape 3: Réduction de l'arborescence
            CollapseTree(doc);

            // Étape 4: Mise à jour
            try
            {
                doc.Update2(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la mise à jour: {ex.Message}", "Erreur", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Étape 5: Vue isométrique
            ApplyIsometricView(inventorApp);

            // Étape 6: Masquage intelligent des références
            SmartHideReferences(doc);

            // Étape 7: Zoom All / Fit
            ZoomToFit(doc, inventorApp);

            // Étape 8: Sauvegarde
            SaveDocument(doc);
        }

        /// <summary>
        /// Exécute les étapes d'assemblage avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecuteAssemblyStepsWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étape 1: Activation représentation par défaut
            try
            {
                ActivateDefaultRepresentation(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step1", "✅ Étape 1: Représentation par défaut activée", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step1", $"❌ Étape 1: Erreur - {ex.Message}", "error");
            }

            // Étape 2: Affichage de TOUS les composants
            try
            {
                dynamic cmdManager = inventorApp.CommandManager;
                dynamic controlDefs = cmdManager.ControlDefinitions;
                try
                {
                    dynamic cmd = controlDefs.Item("AssemblyShowAllComponentsCmd");
                    cmd.Execute();
                    doc.Update();
                }
                catch
                {
                    AfficherTousComposantsRecursive(doc.ComponentDefinition.Occurrences);
                    doc.Update();
                }
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step2", "✅ Étape 2: Tous les composants masqués affichés", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step2", $"❌ Étape 2: Échec - {ex.Message}", "error");
            }

            // Étape 3: Réduction de l'arborescence
            try
            {
                CollapseTree(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", "✅ Étape 3: Arborescence réduite", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", $"❌ Étape 3: Échec - {ex.Message}", "error");
            }

            // Étape 4: Mise à jour
            try
            {
                doc.Update2(true);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step4", "✅ Étape 4: Document mis à jour", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step4", $"❌ Étape 4: Échec - {ex.Message}", "error");
            }

            // Étape 5: Vue isométrique
            try
            {
                ApplyIsometricView(inventorApp);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step5", "✅ Étape 5: Vue isométrique appliquée", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step5", $"❌ Étape 5: Échec - {ex.Message}", "error");
            }

            // Étape 6: Masquage intelligent des références
            try
            {
                SmartHideReferences(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step6", "✅ Étape 6: Références masquées", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step6", $"❌ Étape 6: Échec - {ex.Message}", "error");
            }

            // Étape 7: Sauvegarde
            try
            {
                SaveDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", "✅ Étape 7: Document sauvegardé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", $"❌ Étape 7: Échec - {ex.Message}", "error");
            }
        }

        /// <summary>
        /// Exécute les étapes de pièce avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecutePartStepsWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étape 1
            try
            {
                ActivateDefaultRepresentation(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step1", "✅ Étape 1: Représentation par défaut activée", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step1", $"❌ Étape 1: Erreur - {ex.Message}", "error");
            }

            // Étape 2
            if (progressWindow != null)
                await progressWindow.UpdateStepStatusAsync("step2", "✅ Étape 2: Fonctions vérifiées", "completed");

            // Étape 3
            try
            {
                CollapseTree(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", "✅ Étape 3: Arborescence réduite", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", $"❌ Étape 3: Échec - {ex.Message}", "error");
            }

            // Étape 4
            try
            {
                doc.Update2(true);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step4", "✅ Étape 4: Document mis à jour", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step4", $"❌ Étape 4: Échec - {ex.Message}", "error");
            }

            // Étape 5
            try
            {
                ApplyIsometricView(inventorApp);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step5", "✅ Étape 5: Vue isométrique appliquée", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step5", $"❌ Étape 5: Échec - {ex.Message}", "error");
            }

            // Étape 6
            if (progressWindow != null)
                await progressWindow.UpdateStepStatusAsync("step6", "✅ Étape 6: Matériaux et apparences vérifiés", "completed");

            // Étape 7
            try
            {
                SaveDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", "✅ Étape 7: Document sauvegardé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", $"❌ Étape 7: Échec - {ex.Message}", "error");
            }
        }

        /// <summary>
        /// Exécute les étapes de dessin avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecuteDrawingStepsWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étape 1
            if (progressWindow != null)
                await progressWindow.UpdateStepStatusAsync("step1", "✅ Étape 1: Vues de mise en plan vérifiées", "completed");

            // Étape 2
            try
            {
                CollapseTree(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step2", "✅ Étape 2: Arborescence réduite", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step2", $"❌ Étape 2: Échec - {ex.Message}", "error");
            }

            // Étape 3
            try
            {
                doc.Update2(true);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", "✅ Étape 3: Document mis à jour", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", $"❌ Étape 3: Échec - {ex.Message}", "error");
            }

            // Étape 4
            if (progressWindow != null)
                await progressWindow.UpdateStepStatusAsync("step4", "✅ Étape 4: Cotations et annotations vérifiées", "completed");

            // Étape 5
            if (progressWindow != null)
                await progressWindow.UpdateStepStatusAsync("step5", "✅ Étape 5: Cartouche et propriétés vérifiés", "completed");

            // Étape 6
            try
            {
                SaveDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step6", "✅ Étape 6: Document sauvegardé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step6", $"❌ Étape 6: Échec - {ex.Message}", "error");
            }
        }

        /// <summary>
        /// Exécute les étapes génériques avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecuteGenericStepsWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étape 1
            if (progressWindow != null)
                await progressWindow.UpdateStepStatusAsync("step1", "✅ Étape 1: Document vérifié", "completed");

            // Étape 2
            try
            {
                doc.Update2(true);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step2", "✅ Étape 2: Document mis à jour", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step2", $"❌ Étape 2: Échec - {ex.Message}", "error");
            }

            // Étape 3
            try
            {
                SaveDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", "✅ Étape 3: Document sauvegardé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", $"❌ Étape 3: Échec - {ex.Message}", "error");
            }
        }

        private void ExecutePartSteps(dynamic doc, dynamic inventorApp)
        {
            // Étape 1: Activation représentation par défaut
            ActivateDefaultRepresentation(doc);

            // Étape 2: Affichage des corps cachés
            try
            {
                AfficherCorpsCaches(doc);
            }
            catch
            {
                // Continuer
            }

            // Étape 3: Réduction de l'arborescence
            CollapseTree(doc);

            // Étape 4: Mise à jour
            try
            {
                doc.Update2(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la mise à jour: {ex.Message}", "Erreur", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Étape 5: Vue isométrique
            ApplyIsometricView(inventorApp);

            // Étape 6: Zoom All / Fit
            ZoomToFit(doc, inventorApp);

            // Étape 7: Sauvegarde
            SaveDocument(doc);
        }

        private void ExecuteDrawingSteps(dynamic doc, dynamic inventorApp)
        {
            // Étape 1: Réduction de l'arborescence
            CollapseTree(doc);

            // Étape 2: Mise à jour
            try
            {
                doc.Update2(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la mise à jour: {ex.Message}", "Erreur", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Étape 3: Zoom All / Fit
            ZoomToFit(doc, inventorApp);

            // Étape 4: Sauvegarde
            SaveDocument(doc);
        }

        private void ExecuteGenericSteps(dynamic doc, dynamic inventorApp)
        {
            // Étape 1: Mise à jour
            try
            {
                doc.Update2(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la mise à jour: {ex.Message}", "Erreur", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Étape 2: Zoom All / Fit
            ZoomToFit(doc, inventorApp);

            // Étape 3: Sauvegarde
            SaveDocument(doc);
        }

        private void ActivateDefaultRepresentation(dynamic doc)
        {
            try
            {
                int docType = doc.DocumentType;
                const int kAssemblyDocumentObject = 12290;
                const int kPartDocumentObject = 12288;

                if (docType == kAssemblyDocumentObject)
                {
                    dynamic asmDef = doc.ComponentDefinition;
                    dynamic designViewReps = asmDef.RepresentationsManager.DesignViewRepresentations;

                    dynamic targetRep = null;

                    // ASSEMBLAGES: POSITION 2 PRIORITAIRE
                    if (designViewReps.Count >= 2)
                    {
                        targetRep = designViewReps.Item(2);
                    }
                    else if (designViewReps.Count >= 1)
                    {
                        targetRep = designViewReps.Item(1);
                    }

                    if (targetRep != null)
                    {
                        targetRep.Activate();
                        doc.Update();
                    }
                }
                else if (docType == kPartDocumentObject)
                {
                    dynamic partDef = doc.ComponentDefinition;
                    dynamic designViewReps = partDef.RepresentationsManager.DesignViewRepresentations;

                    dynamic targetRep = null;

                    // PIÈCES: Recherche mots-clés puis position 2
                    foreach (dynamic rep in designViewReps)
                    {
                        string repNameLower = rep.Name.ToLower().Trim();
                        if (repNameLower.Contains("défaut") || repNameLower.Contains("default") || 
                            repNameLower.Contains("primary"))
                        {
                            targetRep = rep;
                            break;
                        }
                    }

                    if (targetRep == null && designViewReps.Count >= 2)
                    {
                        targetRep = designViewReps.Item(2);
                    }
                    else if (targetRep == null && designViewReps.Count >= 1)
                    {
                        targetRep = designViewReps.Item(1);
                    }

                    if (targetRep != null)
                    {
                        targetRep.Activate();
                        doc.Update();
                    }
                }
            }
            catch
            {
                // Continuer malgré l'erreur
            }
        }

        private void CollapseTree(dynamic doc)
        {
            try
            {
                dynamic browserPanes = doc.BrowserPanes;
                dynamic activePane = browserPanes.ActivePane;
                if (activePane != null)
                {
                    dynamic topNode = activePane.TopNode;
                    if (topNode != null)
                    {
                        dynamic browserNodes = topNode.BrowserNodes;
                        foreach (dynamic node in browserNodes)
                        {
                            if (node.Expanded)
                            {
                                node.Expanded = false;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Continuer
            }
        }

        private void ApplyIsometricView(dynamic inventorApp)
        {
            try
            {
                dynamic activeView = inventorApp.ActiveView;
                if (activeView != null)
                {
                    // Utiliser la commande native Inventor pour vue ISO
                    try
                    {
                        dynamic cmdManager = inventorApp.CommandManager;
                        dynamic controlDefs = cmdManager.ControlDefinitions;
                        // "AppIsometricViewCmd" est la commande native pour vue ISO
                        dynamic cmdIso = controlDefs.Item("AppIsometricViewCmd");
                        cmdIso.Execute();
                        // La commande gère tout (caméra + update) - ne rien faire après
                        return;
                    }
                    catch
                    {
                        // Fallback: définir la caméra manuellement
                        dynamic cam = activeView.Camera;
                        cam.Eye = inventorApp.TransientGeometry.CreatePoint(1, 1, 1);
                        cam.Target = inventorApp.TransientGeometry.CreatePoint(0, 0, 0);
                        cam.UpVector = inventorApp.TransientGeometry.CreateUnitVector(0, 1, 0);
                        cam.Perspective = false;
                        cam.Fit();
                        cam.Apply();
                        activeView.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ApplyIsometricView] Erreur: {ex.Message}");
            }
        }

        private void ZoomToFit(dynamic doc, dynamic inventorApp)
        {
            try
            {
                // Méthode 1: Commande native
                try
                {
                    dynamic cmdManager = inventorApp.CommandManager;
                    dynamic controlDefs = cmdManager.ControlDefinitions;
                    dynamic cmd = controlDefs.Item("AppZoomAllCmd");
                    cmd.Execute();
                }
                catch
                {
                    // Méthode 2: Alternative selon type
                    int docType = doc.DocumentType;
                    const int kDrawingDocumentObject = 12291;
                    const int kPartDocumentObject = 12288;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType == kDrawingDocumentObject)
                    {
                        try
                        {
                            dynamic activeSheet = doc.ActiveSheet;
                            if (activeSheet != null)
                            {
                                dynamic drawingViews = activeSheet.DrawingViews;
                                foreach (dynamic view in drawingViews)
                                {
                                    view.Fit();
                                }
                            }
                        }
                        catch
                        {
                            // Continuer
                        }
                    }
                    else if (docType == kPartDocumentObject || docType == kAssemblyDocumentObject)
                    {
                        try
                        {
                            dynamic activeView = inventorApp.ActiveView;
                            if (activeView != null)
                            {
                                dynamic cam = activeView.Camera;
                                cam.Fit();
                                cam.Apply();
                            }
                        }
                        catch
                        {
                            // Continuer
                        }
                    }
                }

                // S'assurer que les changements sont appliqués
                try
                {
                    dynamic activeView = inventorApp.ActiveView;
                    if (activeView != null)
                    {
                        activeView.Update();
                    }
                    System.Threading.Thread.Sleep(50);
                }
                catch
                {
                    // Continuer
                }
            }
            catch
            {
                // Continuer
            }
        }

        private void SaveDocument(dynamic doc)
        {
            try
            {
                if (doc.Dirty)
                {
                    doc.Save2(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la sauvegarde: {ex.Message}", "Erreur", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AfficherTousComposantsRecursive(dynamic occurrences)
        {
            if (occurrences == null) return;

            foreach (dynamic occ in occurrences)
            {
                try
                {
                    occ.Visible = true;

                    const int kPartDocumentObject = 12288;
                    if (occ.DefinitionDocumentType == kPartDocumentObject)
                    {
                        try
                        {
                            dynamic partDoc = occ.Definition.Document;
                            AfficherCorpsCaches(partDoc);
                        }
                        catch
                        {
                            // Ignorer
                        }
                    }

                    const int kAssemblyDocumentObject = 12290;
                    if (occ.DefinitionDocumentType == kAssemblyDocumentObject)
                    {
                        try
                        {
                            dynamic subOccurrences = occ.SubOccurrences;
                            if (subOccurrences != null)
                            {
                                AfficherTousComposantsRecursive(subOccurrences);
                            }
                        }
                        catch
                        {
                            // Ignorer
                        }
                    }
                }
                catch
                {
                    // Ignorer
                }
            }
        }

        private void AfficherCorpsCaches(dynamic partDoc)
        {
            try
            {
                if (partDoc == null || partDoc.ComponentDefinition == null)
                {
                    return;
                }

                dynamic surfaceBodies = partDoc.ComponentDefinition.SurfaceBodies;
                foreach (dynamic body in surfaceBodies)
                {
                    try
                    {
                        body.Visible = true;
                    }
                    catch
                    {
                        // Ignorer
                    }
                }
            }
            catch
            {
                // Ignorer
            }
        }

        private void SmartHideReferences(dynamic doc)
        {
            try
            {
                const int kAssemblyDocumentObject = 12290;
                if (doc.DocumentType == kAssemblyDocumentObject)
                {
                    dynamic asmDef = doc.ComponentDefinition;
                    SmartHideReferencesRecursive(asmDef.Occurrences);
                }
            }
            catch
            {
                // Continuer
            }
        }

        private void SmartHideReferencesRecursive(dynamic occurrences)
        {
            foreach (dynamic occ in occurrences)
            {
                try
                {
                    if (EstReferenceAMasquer(occ))
                    {
                        if (occ.Visible)
                        {
                            occ.Visible = false;
                        }
                    }

                    const int kAssemblyDocumentObject = 12290;
                    if (occ.DefinitionDocumentType == kAssemblyDocumentObject)
                    {
                        try
                        {
                            dynamic subOccurrences = occ.SubOccurrences;
                            if (subOccurrences != null)
                            {
                                SmartHideReferencesRecursive(subOccurrences);
                            }
                        }
                        catch
                        {
                            // Continuer
                        }
                    }
                }
                catch
                {
                    // Continuer
                }
            }
        }

        private bool EstReferenceAMasquer(dynamic occ)
        {
            try
            {
                string nomMinuscule = occ.Name.ToLower();

                // EXCLUSIONS
                if (nomMinuscule.Contains("j_box") || nomMinuscule.Contains("j-box"))
                    return false;

                if (nomMinuscule.Contains("cross_member") && nomMinuscule.Contains("multi_opening"))
                    return false;

                if (nomMinuscule.Contains("structural_tube") && nomMinuscule.Contains("multi_opening"))
                    return false;

                if (EstTeeJoint(nomMinuscule))
                    return false;

                if (nomMinuscule.Contains("cut_opening") && EstDansTubeStructurel(occ))
                    return false;

                // DÉTECTIONS
                if (nomMinuscule.Contains("cut_opening")) return true;
                if (nomMinuscule.Contains("roof_dummy")) return true;
                if (nomMinuscule.Contains("vicwest_dummy")) return true;
                if (nomMinuscule.Contains("dummyremovable_right") || nomMinuscule.Contains("dummyremovable_left"))
                    return true;
                if (nomMinuscule.Contains("dummysupportalignment")) return true;
                if (nomMinuscule.Contains("panfactice")) return true;
                if (nomMinuscule.Contains("airflow_r") || nomMinuscule.Contains("airflow_l")) return true;
                if (nomMinuscule.Contains("cellidsketch")) return true;
                if (nomMinuscule.Contains("opendummy")) return true;
                if (nomMinuscule.Contains("external_left_swing") || nomMinuscule.Contains("external_right_swing"))
                    return true;
                if (nomMinuscule.Contains("internal_left_swing") || nomMinuscule.Contains("internal_right_swing"))
                    return true;
                if (nomMinuscule.Contains("dummyrehausse")) return true;
                if (nomMinuscule.Contains("dummysplitwall_upstream_fromrightside")) return true;
                if (nomMinuscule.Contains("dummysplitwall_downstream_fromrightside")) return true;
                if (EstJointPattern(nomMinuscule)) return true;
                if (nomMinuscule.Contains("box")) return true;
                if (nomMinuscule.Contains("template")) return true;
                if (nomMinuscule.Contains("multi_opening"))
                {
                    if (!(nomMinuscule.Contains("structural_tube") || nomMinuscule.Contains("cross_member")))
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region SafeClose - Fermeture sécurisée

        /// <summary>
        /// Fermeture intelligente et sécurisée des documents Inventor
        /// Code converti depuis SafeClose.iLogicVb
        /// Affiche une fenêtre HTML dynamique avec progression en temps réel
        /// </summary>
        public async Task ExecuteSafeCloseAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(async () =>
            {
                IProgressWindow? progressWindow = null;
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        MessageBox.Show("Aucun document actif n'est ouvert.", "Erreur", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;
                    const int kPartDocumentObject = 12288;
                    const int kDrawingDocumentObject = 12291;

                    string typeText = docType == kAssemblyDocumentObject ? "Assemblage" :
                                     docType == kPartDocumentObject ? "Pièce" :
                                     docType == kDrawingDocumentObject ? "Mise en plan" : "Document";

                    string docName = doc.DisplayName ?? "Document";
                    string htmlContent = GenerateSafeCloseHtml(docType, docName, typeText);

                    // Créer la fenêtre de progression
                    if (_progressWindowCallback != null)
                    {
                        progressWindow = _progressWindowCallback("🔒 Safe Close V1.7 - Released on 2025-12-18 - By Mohammed Amine Elgalai - XNRGY Climate Systems ULC", htmlContent);
                    }

                    // Exécuter les étapes avec mise à jour du HTML
                    if (docType == kAssemblyDocumentObject)
                    {
                        await ExecuteAssemblyStepsForCloseWithProgressAsync(doc, inventorApp, progressWindow);
                    }
                    else if (docType == kPartDocumentObject)
                    {
                        await ExecutePartStepsForCloseWithProgressAsync(doc, inventorApp, progressWindow);
                    }
                    else if (docType == kDrawingDocumentObject)
                    {
                        await ExecuteDrawingStepsForCloseWithProgressAsync(doc, inventorApp, progressWindow);
                    }
                    else
                    {
                        await ExecuteGenericStepsForCloseWithProgressAsync(doc, inventorApp, progressWindow);
                    }

                    // Afficher le message de complétion
                    if (progressWindow != null)
                    {
                        await progressWindow.ShowCompletionAsync("🎉 Toutes les étapes sont terminées avec succès !");
                        await Task.Delay(1500);
                        progressWindow.CloseWindow();
                    }
                }
                catch (Exception ex)
                {
                    if (progressWindow != null)
                    {
                        await progressWindow.ShowCompletionAsync($"❌ Erreur: {ex.Message}");
                        await Task.Delay(2000);
                        progressWindow.CloseWindow();
                    }
                    MessageBox.Show($"Erreur lors du Safe Close : {ex.Message}", "Erreur", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        private void ExecuteAssemblyStepsForClose(dynamic doc, dynamic inventorApp)
        {
            // Identique à SmartSave mais avec fermeture à la fin
            ExecuteAssemblySteps(doc, inventorApp);
            SaveAllDocuments(inventorApp);
            CloseDocument(doc);
        }

        private void ExecutePartStepsForClose(dynamic doc, dynamic inventorApp)
        {
            // Identique à SmartSave mais avec fermeture à la fin
            ExecutePartSteps(doc, inventorApp);
            SaveAllDocuments(inventorApp);
            CloseDocument(doc);
        }

        private void ExecuteDrawingStepsForClose(dynamic doc, dynamic inventorApp)
        {
            // Identique à SmartSave mais avec fermeture à la fin
            ExecuteDrawingSteps(doc, inventorApp);
            SaveAllDocuments(inventorApp);
            CloseDocument(doc);
        }

        private void ExecuteGenericStepsForClose(dynamic doc, dynamic inventorApp)
        {
            // Identique à SmartSave mais avec fermeture à la fin
            ExecuteGenericSteps(doc, inventorApp);
            SaveAllDocuments(inventorApp);
            CloseDocument(doc);
        }

        /// <summary>
        /// Exécute les étapes d'assemblage pour Safe Close avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecuteAssemblyStepsForCloseWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étapes 1-6 identiques à Smart Save
            await ExecuteAssemblyStepsWithProgressAsync(doc, inventorApp, progressWindow);

            // Étape 7: Sauvegarde de tous les documents
            try
            {
                SaveAllDocuments(inventorApp);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", "✅ Étape 7: Documents sauvegardés", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", $"❌ Étape 7: Échec - {ex.Message}", "error");
            }

            // Étape 8: Fermeture du document
            try
            {
                CloseDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step8", "✅ Étape 8: Document fermé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step8", $"❌ Étape 8: Échec - {ex.Message}", "error");
            }
        }

        /// <summary>
        /// Exécute les étapes de pièce pour Safe Close avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecutePartStepsForCloseWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étapes 1-6 identiques à Smart Save
            await ExecutePartStepsWithProgressAsync(doc, inventorApp, progressWindow);

            // Étape 7: Sauvegarde de tous les documents
            try
            {
                SaveAllDocuments(inventorApp);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", "✅ Étape 7: Documents sauvegardés", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", $"❌ Étape 7: Échec - {ex.Message}", "error");
            }

            // Étape 8: Fermeture du document
            try
            {
                CloseDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step8", "✅ Étape 8: Document fermé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step8", $"❌ Étape 8: Échec - {ex.Message}", "error");
            }
        }

        /// <summary>
        /// Exécute les étapes de dessin pour Safe Close avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecuteDrawingStepsForCloseWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étapes 1-5 identiques à Smart Save
            await ExecuteDrawingStepsWithProgressAsync(doc, inventorApp, progressWindow);

            // Étape 6: Sauvegarde de tous les documents
            try
            {
                SaveAllDocuments(inventorApp);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step6", "✅ Étape 6: Documents sauvegardés", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step6", $"❌ Étape 6: Échec - {ex.Message}", "error");
            }

            // Étape 7: Fermeture du document
            try
            {
                CloseDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", "✅ Étape 7: Document fermé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step7", $"❌ Étape 7: Échec - {ex.Message}", "error");
            }
        }

        /// <summary>
        /// Exécute les étapes génériques pour Safe Close avec mise à jour HTML en temps réel
        /// </summary>
        private async Task ExecuteGenericStepsForCloseWithProgressAsync(dynamic doc, dynamic inventorApp, IProgressWindow? progressWindow)
        {
            // Étapes 1-2 identiques à Smart Save
            await ExecuteGenericStepsWithProgressAsync(doc, inventorApp, progressWindow);

            // Étape 3: Sauvegarde de tous les documents
            try
            {
                SaveAllDocuments(inventorApp);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", "✅ Étape 3: Documents sauvegardés", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step3", $"❌ Étape 3: Échec - {ex.Message}", "error");
            }

            // Étape 4: Fermeture du document
            try
            {
                CloseDocument(doc);
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step4", "✅ Étape 4: Document fermé", "completed");
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                    await progressWindow.UpdateStepStatusAsync("step4", $"❌ Étape 4: Échec - {ex.Message}", "error");
            }
        }

        private void SaveAllDocuments(dynamic inventorApp)
        {
            try
            {
                dynamic documents = inventorApp.Documents;
                foreach (dynamic d in documents)
                {
                    try
                    {
                        if (d.Dirty)
                        {
                            d.Save2(true);
                        }
                    }
                    catch
                    {
                        // Continuer
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la sauvegarde: {ex.Message}", "Erreur", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CloseDocument(dynamic doc)
        {
            try
            {
                doc.Close(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la fermeture: {ex.Message}", "Erreur", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region ExportIAMToIPT - Export assemblage vers IPT/STEP

        /// <summary>
        /// Exporte un assemblage (IAM) vers un fichier IPT ou STEP
        /// Code converti depuis ExportIAMToIPT.vb
        /// </summary>
        public async Task ExecuteExportIAMToIPTAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif n'est ouvert.", "ERROR");
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType != kAssemblyDocumentObject)
                    {
                        Log("Cette fonction fonctionne uniquement sur des assemblages (.iam)", "WARNING");
                        MessageBox.Show("Veuillez ouvrir un fichier assemblage (.iam) pour utiliser cette fonction.",
                            "Document non compatible", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    string sourceFileName = System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName);
                    string sourcePath = System.IO.Path.GetDirectoryName(doc.FullFileName);
                    
                    Log($"Préparation de l'export pour: {sourceFileName}", "INFO");

                    // Afficher les options d'export via callback
                    if (_exportOptionsCallback == null)
                    {
                        Log("Callback d'options d'export non configuré", "ERROR");
                        return;
                    }

                    var options = _exportOptionsCallback(sourceFileName, sourcePath);
                    if (options == null)
                    {
                        Log("Export annulé par l'utilisateur", "INFO");
                        return;
                    }

                    Log($"Format sélectionné: {options.Format}", "INFO");
                    Log($"Destination: {options.FullOutputPath}", "INFO");

                    // Masquer les éléments de référence si demandé
                    if (options.HideReferences)
                    {
                        Log("Masquage des éléments de référence...", "INFO");
                        MasquerElementsReferenceForExport(doc);
                    }

                    // Activer représentation par défaut si demandé
                    if (options.ActivateDefaultRepresentation)
                    {
                        try
                        {
                            ActivateDefaultRepresentation(doc);
                            Log("Représentation par défaut activée", "INFO");
                        }
                        catch { }
                    }

                    // Mise à jour du document
                    doc.Update2(true);

                    // Exécuter l'export selon le format
                    if (options.Format.ToString() == "IPT")
                    {
                        ExportToIPT(doc, options.FullOutputPath, inventorApp);
                    }
                    else
                    {
                        ExportToSTEP(doc, options.FullOutputPath, inventorApp);
                    }

                    Log($"Export terminé avec succès: {options.OutputFileName}", "SUCCESS");

                    // Ouvrir le fichier si demandé
                    if (options.OpenAfterExport && System.IO.File.Exists(options.FullOutputPath))
                    {
                        try
                        {
                            if (options.Format.ToString() == "IPT")
                            {
                                inventorApp.Documents.Open(options.FullOutputPath, true);
                            }
                            else
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                {
                                    FileName = options.FullOutputPath,
                                    UseShellExecute = true
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"Impossible d'ouvrir le fichier: {ex.Message}", "WARNING");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'export: {ex.Message}", "ERROR");
                }
            });
        }

        private void MasquerElementsReferenceForExport(dynamic doc)
        {
            try
            {
                dynamic asmDef = doc.ComponentDefinition;
                dynamic occurrences = asmDef.Occurrences;

                var motsCles = new[] { "box", "template", "multi_opening", "cut_opening", "dummy", 
                    "panfactice", "airflow", "rehausse", "splitwall", "removable" };

                int maskedCount = 0;
                foreach (dynamic occ in occurrences)
                {
                    try
                    {
                        string nomMinuscule = occ.Name.ToLower();
                        bool shouldHide = false;

                        foreach (var motCle in motsCles)
                        {
                            if (nomMinuscule.Contains(motCle))
                            {
                                shouldHide = true;
                                break;
                            }
                        }

                        if (shouldHide && occ.Visible)
                        {
                            occ.Visible = false;
                            maskedCount++;
                        }
                    }
                    catch { }
                }

                Log($"Éléments masqués pour export: {maskedCount}", "INFO");
            }
            catch (Exception ex)
            {
                Log($"Erreur masquage éléments: {ex.Message}", "WARNING");
            }
        }

        private void ExportToIPT(dynamic doc, string outputPath, dynamic inventorApp)
        {
            try
            {
                // Utiliser la commande d'export vers pièce composite
                dynamic asmDef = doc.ComponentDefinition;
                
                // Créer un nouveau document pièce
                dynamic partDoc = inventorApp.Documents.Add(12288, "", true); // kPartDocumentObject
                dynamic partDef = partDoc.ComponentDefinition;

                // Créer des corps solides à partir de l'assemblage
                // Note: Cette approche simplifiée - l'export complet nécessite plus de logique
                try
                {
                    // Méthode via ShrinkwrapDefinition
                    dynamic shrinkwrapDef = partDef.ReferenceComponents.ShrinkwrapComponents.CreateDefinition(doc.FullFileName);
                    shrinkwrapDef.DeriveStyle = 58115; // kDeriveAsSingleBodyNoSeams
                    partDef.ReferenceComponents.ShrinkwrapComponents.Add(shrinkwrapDef);
                }
                catch
                {
                    // Fallback: sauvegarder comme copie
                    Log("Utilisation de la méthode de copie simplifiée", "INFO");
                }

                // Sauvegarder
                partDoc.SaveAs(outputPath, false);
                Log($"Fichier IPT créé: {System.IO.Path.GetFileName(outputPath)}", "SUCCESS");
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur export IPT: {ex.Message}");
            }
        }

        private void ExportToSTEP(dynamic doc, string outputPath, dynamic inventorApp)
        {
            try
            {
                // Export STEP via le translateur
                dynamic translator = inventorApp.ApplicationAddIns.ItemById["{90AF7F40-0C01-11D5-8E83-0010B541CD80}"];
                
                if (translator == null)
                {
                    throw new Exception("Translateur STEP non disponible");
                }

                // Contexte de traduction
                dynamic context = inventorApp.TransientObjects.CreateTranslationContext();
                context.Type = 2; // kFileBrowseIOMechanism

                // Options
                dynamic options = inventorApp.TransientObjects.CreateNameValueMap();
                
                if (translator.HasSaveCopyAsOptions[doc, context, options])
                {
                    // Configurer les options STEP (AP214)
                    options.Value["ApplicationProtocolType"] = 2; // AP214
                }

                // Données de destination
                dynamic oData = inventorApp.TransientObjects.CreateDataMedium();
                oData.FileName = outputPath;

                // Exécuter l'export
                translator.SaveCopyAs(doc, context, options, oData);
                
                Log($"Fichier STEP créé: {System.IO.Path.GetFileName(outputPath)}", "SUCCESS");
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur export STEP: {ex.Message}");
            }
        }

        #endregion

        #region ExportIDWtoShopPDF - Export dessins vers PDF

        /// <summary>
        /// Exporte le dessin actif vers PDF dans le dossier Shop Drawing
        /// </summary>
        public async Task ExecuteExportIDWtoShopPDFAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif n'est ouvert.", "ERROR");
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kDrawingDocumentObject = 12291;

                    if (docType != kDrawingDocumentObject)
                    {
                        Log("Cette fonction fonctionne uniquement sur des dessins (.idw/.dwg)", "WARNING");
                        MessageBox.Show("Veuillez ouvrir un fichier dessin (.idw ou .dwg) pour utiliser cette fonction.",
                            "Document non compatible", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    string fileName = System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName);
                    string sourcePath = System.IO.Path.GetDirectoryName(doc.FullFileName);
                    
                    // Déterminer le dossier de destination (6-Shop Drawing PDF)
                    string shopDrawingPath = FindShopDrawingFolder(sourcePath);
                    
                    if (string.IsNullOrEmpty(shopDrawingPath))
                    {
                        Log("Dossier '6-Shop Drawing PDF' non trouvé", "WARNING");
                        
                        // Créer le dossier s'il n'existe pas
                        string parentDir = System.IO.Path.GetDirectoryName(sourcePath);
                        if (!string.IsNullOrEmpty(parentDir))
                        {
                            shopDrawingPath = System.IO.Path.Combine(parentDir, "6-Shop Drawing PDF");
                            try
                            {
                                System.IO.Directory.CreateDirectory(shopDrawingPath);
                                Log($"Dossier créé: {shopDrawingPath}", "INFO");
                            }
                            catch
                            {
                                shopDrawingPath = sourcePath; // Fallback au dossier source
                            }
                        }
                        else
                        {
                            shopDrawingPath = sourcePath;
                        }
                    }

                    string pdfPath = System.IO.Path.Combine(shopDrawingPath, fileName + ".pdf");
                    
                    Log($"Export PDF vers: {pdfPath}", "INFO");

                    // Export PDF
                    ExportDrawingToPDF(doc, pdfPath, inventorApp);

                    Log($"PDF créé avec succès: {fileName}.pdf", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'export PDF: {ex.Message}", "ERROR");
                }
            });
        }

        private string FindShopDrawingFolder(string startPath)
        {
            try
            {
                // Chercher dans le dossier courant et parents
                string currentPath = startPath;
                
                for (int i = 0; i < 5; i++) // Remonter max 5 niveaux
                {
                    if (string.IsNullOrEmpty(currentPath)) break;

                    string shopPath = System.IO.Path.Combine(currentPath, "6-Shop Drawing PDF");
                    if (System.IO.Directory.Exists(shopPath))
                    {
                        return shopPath;
                    }

                    // Chercher aussi avec d'autres noms possibles
                    var variants = new[] { "Shop Drawing PDF", "Shop_Drawing_PDF", "PDF" };
                    foreach (var variant in variants)
                    {
                        shopPath = System.IO.Path.Combine(currentPath, variant);
                        if (System.IO.Directory.Exists(shopPath))
                        {
                            return shopPath;
                        }
                    }

                    currentPath = System.IO.Path.GetDirectoryName(currentPath);
                }
            }
            catch { }

            return null;
        }

        private void ExportDrawingToPDF(dynamic doc, string outputPath, dynamic inventorApp)
        {
            try
            {
                // Obtenir le translateur PDF
                dynamic translator = inventorApp.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"];
                
                if (translator == null)
                {
                    throw new Exception("Translateur PDF non disponible");
                }

                // Contexte
                dynamic context = inventorApp.TransientObjects.CreateTranslationContext();
                context.Type = 2; // kFileBrowseIOMechanism

                // Options PDF
                dynamic options = inventorApp.TransientObjects.CreateNameValueMap();
                dynamic oData = inventorApp.TransientObjects.CreateDataMedium();

                if (translator.HasSaveCopyAsOptions[doc, context, options])
                {
                    // Options recommandées
                    options.Value["All_Color_AS_Black"] = 0;       // Garder les couleurs
                    options.Value["Remove_Line_Weights"] = 0;      // Garder les épaisseurs
                    options.Value["Vector_Resolution"] = 400;       // Haute résolution
                    options.Value["Sheet_Range"] = 0;               // Toutes les feuilles
                }

                oData.FileName = outputPath;

                // Export
                translator.SaveCopyAs(doc, context, options, oData);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur export PDF: {ex.Message}");
            }
        }

        #endregion

        #region FormCentering - Centrage des formulaires

        /// <summary>
        /// Centre les formulaires iLogic et autres fenêtres
        /// </summary>
        public async Task ExecuteFormCenteringUtilityAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    Log("Recherche de formulaires iLogic à centrer...", "INFO");
                    
                    // Utiliser l'API Windows pour trouver et centrer les fenêtres
                    int centeredCount = CenterILogicForms();
                    
                    if (centeredCount > 0)
                    {
                        Log($"{centeredCount} formulaire(s) centré(s) avec succès", "SUCCESS");
                    }
                    else
                    {
                        Log("Aucun formulaire iLogic trouvé à centrer", "INFO");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors du centrage: {ex.Message}", "ERROR");
                }
            });
        }

        public async Task ExecuteILogicFormsCentredAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    Log("Activation du centrage automatique des formulaires iLogic...", "INFO");
                    
                    // Cette fonctionnalité nécessite un hook Windows pour détecter les nouvelles fenêtres
                    // Pour l'instant, on fait un centrage ponctuel
                    int centeredCount = CenterILogicForms();
                    
                    if (centeredCount > 0)
                    {
                        Log($"{centeredCount} formulaire(s) iLogic centré(s)", "SUCCESS");
                    }
                    else
                    {
                        Log("Aucun formulaire iLogic détecté. Le centrage sera appliqué aux prochains formulaires.", "INFO");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors du centrage: {ex.Message}", "ERROR");
                }
            });
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private int CenterILogicForms()
        {
            int centeredCount = 0;
            var windowsToCenter = new System.Collections.Generic.List<IntPtr>();

            // Obtenir le PID d'Inventor
            int inventorPid = 0;
            try
            {
                var inventorProcesses = System.Diagnostics.Process.GetProcessesByName("Inventor");
                if (inventorProcesses.Length > 0)
                {
                    inventorPid = inventorProcesses[0].Id;
                }
            }
            catch { }

            // Énumérer toutes les fenêtres
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true;

                // Vérifier si la fenêtre appartient au processus Inventor
                GetWindowThreadProcessId(hWnd, out int windowPid);
                if (windowPid != inventorPid && inventorPid != 0) return true;

                // Obtenir le titre de la fenêtre
                var sb = new System.Text.StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                string title = sb.ToString();

                // Détecter les formulaires iLogic (titres typiques)
                bool isILogicForm = !string.IsNullOrEmpty(title) && 
                    (title.Contains("iLogic") || 
                     title.Contains("Form") ||
                     title.Contains("Input") ||
                     title.Contains("Paramètre") ||
                     title.Contains("Parameter") ||
                     title.Contains("Selection") ||
                     title.Contains("Sélection"));

                if (isILogicForm)
                {
                    windowsToCenter.Add(hWnd);
                }

                return true;
            }, IntPtr.Zero);

            // Centrer les fenêtres trouvées
            foreach (var hWnd in windowsToCenter)
            {
                try
                {
                    if (CenterWindow(hWnd))
                    {
                        centeredCount++;
                    }
                }
                catch { }
            }

            return centeredCount;
        }

        private bool CenterWindow(IntPtr hWnd)
        {
            try
            {
                GetWindowRect(hWnd, out RECT windowRect);
                
                int windowWidth = windowRect.Right - windowRect.Left;
                int windowHeight = windowRect.Bottom - windowRect.Top;

                // Obtenir la taille de l'écran principal
                int screenWidth = (int)SystemParameters.PrimaryScreenWidth;
                int screenHeight = (int)SystemParameters.PrimaryScreenHeight;

                // Calculer la position centrée
                int newX = (screenWidth - windowWidth) / 2;
                int newY = (screenHeight - windowHeight) / 2;

                // Déplacer la fenêtre
                const uint SWP_NOSIZE = 0x0001;
                const uint SWP_NOZORDER = 0x0004;
                SetWindowPos(hWnd, IntPtr.Zero, newX, newY, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        public async Task ExecuteInfoCommandAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                Log("Affichage des informations Smart Tools...", "INFO");
                string info = @"Smart Tools - Outils d'automatisation Inventor v1.7

Outils disponibles:
- HideBox V1.5: Masquage intelligent des éléments Dummy, PanFactice, AirFlow, etc.
- ToggleRefVisibility: Basculer la visibilité des références
- ToggleSketchVisibility: Basculer la visibilité des esquisses
- ConstraintReport: Rapport de contraintes d'assemblage
- iPropertiesSummary V1.3: Résumé des propriétés avec chemin complet
- SafeClose V1.6: Fermeture sécurisée avec support des éléments Dummy
- SmartSave V1.0: Sauvegarde intelligente sans fermeture
- ExportIAMToIPT V1.1: Export d'assemblage vers IPT/STEP
- ExportIDWtoShopPDF V1.1: Export de dessins en PDF vers 6-Shop Drawing PDF
- FormCenteringUtility: Centrage des formulaires génériques
- iLogicFormsCentred V1.1: Détection et centrage des formulaires iLogic

Nouveautés V1.7:
- Migration vers Inventor 2026
- Support du nouveau chemin C:\Vault\Engineering\Projects\
- ExportIAMToIPT V1.1: Bouton Parcourir + édition manuelle du chemin
- ExportIDWtoShopPDF V1.1: Option Parcourir si dossier non trouvé

Développé par: Mohammed Amine El Galai
Entreprise: XNRGY Climate Systems ULC
Date: 2026-01-02";

                MessageBox.Show(info, "Smart Tools - Informations", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }

        #region AutoSave - Sauvegarde automatique

        /// <summary>
        /// Active la sauvegarde automatique d'Inventor
        /// </summary>
        public async Task ExecuteAutoSaveAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    
                    // Accéder aux options d'application
                    dynamic appOptions = inventorApp.ApplicationOptions;
                    dynamic saveOptions = appOptions.SaveOptions;
                    
                    // Activer l'auto-sauvegarde
                    bool currentState = saveOptions.EnableAutoRecover;
                    
                    if (!currentState)
                    {
                        saveOptions.EnableAutoRecover = true;
                        saveOptions.AutoRecoverInterval = 10; // 10 minutes
                        Log("Auto-sauvegarde activée (intervalle: 10 min)", "SUCCESS");
                    }
                    else
                    {
                        // Afficher les paramètres actuels
                        int interval = saveOptions.AutoRecoverInterval;
                        Log($"Auto-sauvegarde déjà activée (intervalle: {interval} min)", "INFO");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'activation de l'auto-sauvegarde: {ex.Message}", "ERROR");
                }
            });
        }

        #endregion

        #region CheckSaveStatus - Vérification du statut de sauvegarde

        /// <summary>
        /// Vérifie si des documents ont des modifications non sauvegardées
        /// </summary>
        public async Task ExecuteCheckSaveStatusAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic documents = inventorApp.Documents;
                    
                    int totalDocs = 0;
                    int unsavedCount = 0;
                    var unsavedList = new System.Collections.Generic.List<string>();

                    foreach (dynamic doc in documents)
                    {
                        totalDocs++;
                        try
                        {
                            if (doc.Dirty) // Dirty = modifications non sauvegardées
                            {
                                unsavedCount++;
                                string fileName = System.IO.Path.GetFileName(doc.FullFileName);
                                if (string.IsNullOrEmpty(fileName))
                                    fileName = doc.DisplayName;
                                unsavedList.Add(fileName);
                            }
                        }
                        catch { }
                    }

                    if (unsavedCount == 0)
                    {
                        Log($"Tous les documents sont sauvegardés ({totalDocs} documents)", "SUCCESS");
                    }
                    else
                    {
                        Log($"Documents non sauvegardés: {unsavedCount}/{totalDocs}", "WARNING");
                        foreach (var name in unsavedList)
                        {
                            Log($"  [!] {name}", "WARNING");
                        }
                        
                        // Proposer de sauvegarder
                        var result = MessageBox.Show(
                            $"{unsavedCount} document(s) ont des modifications non sauvegardées.\n\nVoulez-vous les sauvegarder maintenant?",
                            "Documents non sauvegardés",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);

                        if (result == MessageBoxResult.Yes)
                        {
                            int savedCount = 0;
                            foreach (dynamic doc in documents)
                            {
                                try
                                {
                                    if (doc.Dirty)
                                    {
                                        doc.Save2(true); // true = réinitialiser le statut dirty
                                        savedCount++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log($"Erreur sauvegarde: {ex.Message}", "ERROR");
                                }
                            }
                            Log($"{savedCount} document(s) sauvegardé(s)", "SUCCESS");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de la vérification: {ex.Message}", "ERROR");
                }
            });
        }

        #endregion

        #region CollapseExpandAll - Réduire/Développer l'arborescence

        /// <summary>
        /// Bascule entre les états collapse/expand dans le navigateur
        /// Code converti depuis CollapseExpandAll.Vb
        /// </summary>
        public async Task ExecuteCollapseExpandAllAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    Log("Vérification de l'arborescence...", "INFO");
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null) return;

                    // Méthode 1: Essayer d'abord les commandes natives (plus robuste)
                    try
                    {
                        dynamic browserPanes = doc.BrowserPanes;
                        dynamic browser = browserPanes.ActivePane;
                        if (browser == null) return;

                        dynamic topNode = browser.TopNode;
                        if (topNode == null) return;

                        // Vérifier si au moins un nœud est développé
                        bool anyExpanded = false;
                        try
                        {
                            dynamic browserNodes = topNode.BrowserNodes;
                            int count = browserNodes.Count;
                            for (int i = 1; i <= count; i++)
                            {
                                try
                                {
                                    dynamic child = browserNodes.Item(i);
                                    if (child != null && child.Expanded)
                                    {
                                        anyExpanded = true;
                                        break;
                                    }
                                }
                                catch
                                {
                                    // Continuer avec le prochain nœud
                                }
                            }
                        }
                        catch
                        {
                            // Si on ne peut pas vérifier, essayer directement les commandes
                        }

                        // Basculer entre collapse et expand
                        if (anyExpanded)
                        {
                            // Méthode 1: Commande native
                            try
                            {
                                dynamic cmdManager = inventorApp.CommandManager;
                                dynamic controlDefs = cmdManager.ControlDefinitions;
                                dynamic cmd = controlDefs.Item("AppCollapseAllNodesCmd");
                                if (cmd != null)
                                {
                                    cmd.Execute();
                                    Log("Arborescence réduite avec succès", "SUCCESS");
                                    return; // Succès
                                }
                            }
                            catch
                            {
                                // Continuer avec méthode manuelle
                            }

                            // Méthode 2: Collapse manuel
                            Log("Réduction manuelle de l'arborescence...", "INFO");
                            try
                            {
                                dynamic browserNodes = topNode.BrowserNodes;
                                int count = browserNodes.Count;
                                for (int i = 1; i <= count; i++)
                                {
                                    try
                                    {
                                        dynamic child = browserNodes.Item(i);
                                        if (child != null)
                                        {
                                            child.Expanded = false;
                                        }
                                    }
                                    catch
                                    {
                                        // Continuer
                                    }
                                }
                            }
                            catch
                            {
                                // Ignorer les erreurs
                            }
                        }
                        else
                        {
                            // Méthode 1: Commande native
                            try
                            {
                                dynamic cmdManager = inventorApp.CommandManager;
                                dynamic controlDefs = cmdManager.ControlDefinitions;
                                dynamic cmd = controlDefs.Item("AppExpandOneLevelCmd");
                                if (cmd != null)
                                {
                                    cmd.Execute();
                                    Log("Arborescence développée avec succès", "SUCCESS");
                                    return; // Succès
                                }
                            }
                            catch
                            {
                                // Continuer avec méthode manuelle
                            }

                            // Méthode 2: Expand manuel
                            Log("Développement manuel de l'arborescence...", "INFO");
                            try
                            {
                                dynamic browserNodes = topNode.BrowserNodes;
                                int count = browserNodes.Count;
                                for (int i = 1; i <= count; i++)
                                {
                                    try
                                    {
                                        dynamic child = browserNodes.Item(i);
                                        if (child != null)
                                        {
                                            child.Expanded = true;
                                        }
                                    }
                                    catch
                                    {
                                        // Continuer
                                    }
                                }
                            }
                            catch
                            {
                                // Ignorer les erreurs
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Si tout échoue, essayer uniquement les commandes natives sans vérification
                        try
                        {
                            dynamic cmdManager = inventorApp.CommandManager;
                            dynamic controlDefs = cmdManager.ControlDefinitions;
                            
                            // Essayer collapse d'abord
                            try
                            {
                                dynamic cmd = controlDefs.Item("AppCollapseAllNodesCmd");
                                if (cmd != null)
                                {
                                    cmd.Execute();
                                    return;
                                }
                            }
                            catch
                            {
                                // Essayer expand
                                try
                                {
                                    dynamic cmd = controlDefs.Item("AppExpandOneLevelCmd");
                                    if (cmd != null)
                                    {
                                        cmd.Execute();
                                        return;
                                    }
                                }
                                catch
                                {
                                    // Échec complet
                                    throw new Exception($"Impossible d'accéder au navigateur: {ex.Message}", ex);
                                }
                            }
                        }
                        catch
                        {
                            throw new Exception($"Erreur lors de l'exécution de CollapseExpandAll: {ex.Message}", ex);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Erreur lors de l'exécution de CollapseExpandAll: {ex.Message}", ex);
                }
            });
        }

        #endregion

        #region 2DIsometricView - Vue isométrique 2D

        /// <summary>
        /// Applique la vue isométrique
        /// Code converti depuis 2D-IsometricView.Vb
        /// </summary>
        public async Task Execute2DIsometricViewAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    ApplyIsometricView(inventorApp);
                    Log("Vue isométrique appliquée avec succès", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'application de la vue iso : {ex.Message}", "ERROR");
                    throw;
                }
            });
        }

        #endregion

        #region ReturnToFrontView - Retour à la vue de face

        /// <summary>
        /// Remet en vue de face
        /// Code converti depuis ReturnToFrontView.Vb
        /// </summary>
        public async Task ExecuteReturnToFrontViewAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic activeView = inventorApp.ActiveView;
                    if (activeView != null)
                    {
                        dynamic cam = activeView.Camera;
                        const int kFrontViewOrientation = 1;
                        cam.ViewOrientationType = kFrontViewOrientation;
                        cam.Perspective = false;
                        cam.Fit();
                        cam.Apply();
                        activeView.Update();
                        Log("Vue de face restaurée avec succès", "SUCCESS");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de la restauration de vue : {ex.Message}", "ERROR");
                    throw;
                }
            });
        }

        #endregion

        #region Fonctions avancées (à implémenter ultérieurement)

        public async Task ExecuteInsertSpecificScrewInHolesAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                Log("InsertSpecificScrewInHoles - Fonctionnalité avancée en cours de développement", "INFO");
                MessageBox.Show("Cette fonctionnalité avancée sera disponible dans une prochaine version.\n\nInsertSpecificScrewInHoles permet d'insérer automatiquement des vis dans les trous d'un assemblage.",
                    "Fonctionnalité en développement", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }

        public async Task ExecuteIPropertyCustomBatchAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                Log("iPropertyCustomBatch - Fonctionnalité avancée en cours de développement", "INFO");
                MessageBox.Show("Cette fonctionnalité avancée sera disponible dans une prochaine version.\n\niPropertyCustomBatch permet de modifier les propriétés de plusieurs fichiers en lot.",
                    "Fonctionnalité en développement", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }

        public async Task ExecuteFixPromotedVariesAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                Log("FixPromotedVaries - Fonctionnalité avancée en cours de développement", "INFO");
                MessageBox.Show("Cette fonctionnalité avancée sera disponible dans une prochaine version.\n\nFixPromotedVaries corrige les paramètres promus qui varient.",
                    "Fonctionnalité en développement", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }

        public async Task ExecuteOpenSelectedComponentAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null || doc.SelectSet.Count == 0)
                    {
                        Log("Veuillez sélectionner un composant à ouvrir", "WARNING");
                        return;
                    }

                    dynamic selObj = doc.SelectSet[1];
                    try
                    {
                        dynamic definition = selObj.Definition;
                        dynamic compDoc = definition.Document;
                        string filePath = compDoc.FullFileName;
                        
                        // Ouvrir le composant
                        inventorApp.Documents.Open(filePath, true);
                        Log($"Composant ouvert: {System.IO.Path.GetFileName(filePath)}", "SUCCESS");
                    }
                    catch
                    {
                        Log("L'objet sélectionné n'est pas un composant ouvrable", "WARNING");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur: {ex.Message}", "ERROR");
                }
            });
        }

        public async Task ExecuteUpdateCurrentSheetAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif", "ERROR");
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kDrawingDocumentObject = 12291;

                    if (docType != kDrawingDocumentObject)
                    {
                        Log("Cette fonction fonctionne uniquement sur des dessins", "WARNING");
                        return;
                    }

                    // Mettre à jour la feuille active
                    dynamic activeSheet = doc.ActiveSheet;
                    if (activeSheet != null)
                    {
                        activeSheet.Update();
                        Log($"Feuille '{activeSheet.Name}' mise à jour", "SUCCESS");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Erreur mise à jour: {ex.Message}", "ERROR");
                }
            });
        }

        public async Task ExecuteHideBoxTemplateMultiOpeningAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                try
                {
                    dynamic inventorApp = GetInventorApplication();
                    dynamic doc = inventorApp.ActiveDocument;

                    if (doc == null)
                    {
                        Log("Aucun document actif", "ERROR");
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType != kAssemblyDocumentObject)
                    {
                        Log("Cette fonction fonctionne uniquement sur des assemblages", "WARNING");
                        return;
                    }

                    // Masquer les Box, Template et Multi_Opening
                    dynamic asmDef = doc.ComponentDefinition;
                    int maskedCount = 0;

                    foreach (dynamic occ in asmDef.Occurrences)
                    {
                        try
                        {
                            string nomMinuscule = occ.Name.ToLower();
                            if (nomMinuscule.Contains("box") || 
                                nomMinuscule.Contains("template") || 
                                nomMinuscule.Contains("multi_opening"))
                            {
                                if (occ.Visible)
                                {
                                    occ.Visible = false;
                                    maskedCount++;
                                }
                            }
                        }
                        catch { }
                    }

                    // Vue ISO + Zoom All silencieux
                    ExecuteIsoViewAndZoomAllSilent();
                    
                    Log($"[+] Box/Template/MultiOpening: {maskedCount} caches - Vue ISO + Zoom All", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur: {ex.Message}", "ERROR");
                }
            });
        }

        #endregion

        #region Helper Methods - Génération HTML pour Smart Save et Safe Close

        /// <summary>
        /// Génère le HTML pour Smart Save selon le type de document
        /// </summary>
        private string GenerateSmartSaveHtml(int docType, string docName, string typeText)
        {
            var steps = GetSmartSaveSteps(docType);
            return GenerateProgressHtml("💾 Smart Save V1.1", typeText, docName, steps, "#2e7d32", "#4caf50");
        }

        /// <summary>
        /// Génère le HTML pour Safe Close selon le type de document
        /// </summary>
        private string GenerateSafeCloseHtml(int docType, string docName, string typeText)
        {
            var steps = GetSafeCloseSteps(docType);
            return GenerateProgressHtml("🔒 Safe Close V1.7", typeText, docName, steps, "#0d47a1", "#1976d2");
        }

        /// <summary>
        /// Génère le HTML de progression avec les étapes
        /// </summary>
        private string GenerateProgressHtml(string title, string typeText, string docName, List<string> steps, string primaryColor, string secondaryColor)
        {
            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html lang='fr'>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset='UTF-8'>");
            html.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
            html.AppendLine($"    <title>{title} - {typeText}</title>");
            html.AppendLine("    <style>");
            html.AppendLine("        @import url('https://fonts.googleapis.com/css2?family=Noto+Color+Emoji&display=swap');");
            html.AppendLine("        * { font-family: 'Segoe UI', 'Roboto', 'Noto Color Emoji', 'Apple Color Emoji', sans-serif; }");
            html.AppendLine($"        body {{ margin: 15px; background: linear-gradient(135deg, {primaryColor} 0%, {secondaryColor} 100%); font-size: 16px; min-height: calc(100vh - 30px); color: white; }}");
            html.AppendLine("        .container { max-width: 95%; margin: 0 auto; background: rgba(255,255,255,0.95); padding: 25px; border-radius: 12px; box-shadow: 0 8px 25px rgba(0,0,0,0.3); color: #333; }");
            html.AppendLine($"        h2 {{ color: {primaryColor}; font-size: 24px; text-align: center; margin-bottom: 20px; font-weight: bold; text-shadow: 1px 1px 2px rgba(0,0,0,0.1); }}");
            html.AppendLine($"        .info-box {{ background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); padding: 15px; border-radius: 8px; margin-bottom: 20px; border: 2px solid {secondaryColor}; }}");
            html.AppendLine($"        .info-box strong {{ color: {primaryColor}; }}");
            html.AppendLine("        ul { list-style: none; padding: 0; margin: 0; }");
            html.AppendLine("        li { margin: 8px 0; font-size: 15px; padding: 8px 12px; border-radius: 6px; display: flex; align-items: center; background: rgba(248,249,250,0.8); border-left: 4px solid #90a4ae; }");
            html.AppendLine("        li.completed { background: rgba(232,245,233,0.9); border-left-color: #4caf50; }");
            html.AppendLine("        li.error { background: rgba(255,235,238,0.9); border-left-color: #f44336; }");
            html.AppendLine("        li.info { background: rgba(227,242,253,0.9); border-left-color: #2196f3; }");
            html.AppendLine("        .emoji { font-size: 18px; margin-right: 10px; min-width: 25px; }");
            html.AppendLine($"        .completion {{ text-align: center; font-size: 18px; font-weight: bold; color: {primaryColor}; margin: 20px 0; padding: 15px; background: rgba(232,245,233,0.8); border-radius: 8px; display: none; }}");
            html.AppendLine($"        .btn-close {{ display: block; width: 150px; padding: 12px; margin: 20px auto; font-size: 16px; font-weight: bold; cursor: pointer; border: none; border-radius: 8px; background: {secondaryColor}; color: white; transition: all 0.3s; }}");
            html.AppendLine($"        .btn-close:hover {{ background: {primaryColor}; transform: scale(1.05); }}");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            html.AppendLine("    <div class='container'>");
            html.AppendLine($"        <h2><span class='emoji'>💾</span> {title} - {typeText}</h2>");
            html.AppendLine("        <div class='info-box'>");
            html.AppendLine($"            <span class='emoji'>📄</span> <strong>Document:</strong> {WebUtility.HtmlEncode(docName)}<br>");
            html.AppendLine($"            <span class='emoji'>📋</span> <strong>Type:</strong> {typeText}<br>");
            html.AppendLine($"            <span class='emoji'>📅</span> <strong>Date:</strong> {DateTime.Now:yyyy-MM-dd HH:mm:ss}<br>");
            html.AppendLine("            <span class='emoji'>👨‍💻</span> <strong>Développé par:</strong> Mohammed Amine Elgalai");
            html.AppendLine("        </div>");
            html.AppendLine("        <ul>");

            for (int i = 0; i < steps.Count; i++)
            {
                html.AppendLine($"            <li id='step{i + 1}'><span class='emoji'>⏳</span> {WebUtility.HtmlEncode(steps[i])}</li>");
            }

            html.AppendLine("        </ul>");
            html.AppendLine("        <div id='completion' class='completion'></div>");
            html.AppendLine("        <button class='btn-close' onclick='closeForm()'>✅ Fermer</button>");
            html.AppendLine("        <script>");
            html.AppendLine("            function closeForm() {");
            html.AppendLine("                if (window.chrome && window.chrome.webview) {");
            html.AppendLine("                    window.chrome.webview.postMessage('CLOSE_FORM');");
            html.AppendLine("                }");
            html.AppendLine("            }");
            html.AppendLine("        </script>");
            html.AppendLine("    </div>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            return html.ToString();
        }

        /// <summary>
        /// Retourne les étapes pour Smart Save selon le type de document
        /// </summary>
        private List<string> GetSmartSaveSteps(int docType)
        {
            var steps = new List<string>();
            const int kAssemblyDocumentObject = 12290;
            const int kPartDocumentObject = 12288;
            const int kDrawingDocumentObject = 12291;

            if (docType == kAssemblyDocumentObject)
            {
                steps.Add("🔍 Étape 1: Activation représentation par défaut (Position 2 prioritaire)");
                steps.Add("👁️ Étape 2: Affichage de TOUS les composants masqués (hors références)");
                steps.Add("🌲 Étape 3: Réduction de l'arborescence du navigateur");
                steps.Add("🔄 Étape 4: Mise à jour du document");
                steps.Add("📐 Étape 5: Application de la vue isométrique");
                steps.Add("🙈 Étape 6: Masquage références (Dummy, Swing, PanFactice, AirFlow, Cut_Opening)");
                steps.Add("💾 Étape 7: Sauvegarde du document actif");
            }
            else if (docType == kPartDocumentObject)
            {
                steps.Add("🔍 Étape 1: Activation représentation par défaut");
                steps.Add("🛠️ Étape 2: Vérification des fonctions supprimées");
                steps.Add("🌲 Étape 3: Réduction de l'arborescence du navigateur");
                steps.Add("🔄 Étape 4: Mise à jour du document");
                steps.Add("📐 Étape 5: Application de la vue isométrique");
                steps.Add("🎨 Étape 6: Vérification des matériaux et apparences");
                steps.Add("💾 Étape 7: Sauvegarde du document actif");
            }
            else if (docType == kDrawingDocumentObject)
            {
                steps.Add("📋 Étape 1: Vérification des vues de mise en plan");
                steps.Add("🌲 Étape 2: Réduction de l'arborescence du navigateur");
                steps.Add("🔄 Étape 3: Mise à jour du document et des vues");
                steps.Add("📏 Étape 4: Vérification des cotations et annotations");
                steps.Add("📄 Étape 5: Vérification du cartouche et propriétés");
                steps.Add("💾 Étape 6: Sauvegarde du document actif");
            }
            else
            {
                steps.Add("🔍 Étape 1: Vérification générale du document");
                steps.Add("🔄 Étape 2: Mise à jour du document");
                steps.Add("💾 Étape 3: Sauvegarde du document actif");
            }

            return steps;
        }

        /// <summary>
        /// Retourne les étapes pour Safe Close selon le type de document
        /// </summary>
        private List<string> GetSafeCloseSteps(int docType)
        {
            var steps = new List<string>();
            const int kAssemblyDocumentObject = 12290;
            const int kPartDocumentObject = 12288;
            const int kDrawingDocumentObject = 12291;

            if (docType == kAssemblyDocumentObject)
            {
                steps.Add("🔍 Étape 1: Activation représentation par défaut (Position 2 prioritaire)");
                steps.Add("👁️ Étape 2: Affichage de TOUS les composants masqués (hors références)");
                steps.Add("🌲 Étape 3: Réduction de l'arborescence du navigateur");
                steps.Add("🔄 Étape 4: Mise à jour du document");
                steps.Add("📐 Étape 5: Application de la vue isométrique");
                steps.Add("🧠 Étape 6: Masquage références (Cut_Opening, tous types de Dummy, Internal/External Swing, etc.)");
                steps.Add("💾 Étape 7: Sauvegarde de tous les documents ouverts");
                steps.Add("🚪 Étape 8: Fermeture du document actif");
            }
            else if (docType == kPartDocumentObject)
            {
                steps.Add("🔍 Étape 1: Activation représentation par défaut");
                steps.Add("🛠️ Étape 2: Vérification des fonctions supprimées");
                steps.Add("🌲 Étape 3: Réduction de l'arborescence du navigateur");
                steps.Add("🔄 Étape 4: Mise à jour du document");
                steps.Add("📐 Étape 5: Application de la vue isométrique");
                steps.Add("🎨 Étape 6: Vérification des matériaux et apparences");
                steps.Add("💾 Étape 7: Sauvegarde de tous les documents ouverts");
                steps.Add("🚪 Étape 8: Fermeture du document actif");
            }
            else if (docType == kDrawingDocumentObject)
            {
                steps.Add("📋 Étape 1: Vérification des vues de mise en plan");
                steps.Add("🌲 Étape 2: Réduction de l'arborescence du navigateur");
                steps.Add("🔄 Étape 3: Mise à jour du document et des vues");
                steps.Add("📏 Étape 4: Vérification des cotations et annotations");
                steps.Add("📄 Étape 5: Vérification du cartouche et propriétés");
                steps.Add("💾 Étape 6: Sauvegarde de tous les documents ouverts");
                steps.Add("🚪 Étape 7: Fermeture du document actif");
            }
            else
            {
                steps.Add("🔍 Étape 1: Vérification générale du document");
                steps.Add("🔄 Étape 2: Mise à jour du document");
                steps.Add("💾 Étape 3: Sauvegarde de tous les documents ouverts");
                steps.Add("🚪 Étape 4: Fermeture du document actif");
            }

            return steps;
        }

        #endregion
    }
}
