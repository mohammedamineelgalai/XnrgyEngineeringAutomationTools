using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Services
{
    /// <summary>
    /// Service pour exécuter les outils Smart Tools directement via COM Inventor
    /// Code converti depuis VB (règles iLogic) vers C#
    /// </summary>
    public class SmartToolsService
    {
        private readonly InventorService _inventorService;
        private Action<string, string>? _logCallback;

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

        private void Log(string message, string level = "INFO")
        {
            _logCallback?.Invoke(message, level);
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
                    Log("Mise à jour du document...", "INFO");
                    doc.Update2(true);
                    Log("HideBox exécuté avec succès", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'exécution de HideBox: {ex.Message}", "ERROR");
                    throw new Exception($"Erreur lors de l'exécution de HideBox: {ex.Message}", ex);
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

        #region ToggleRefVisibility - Basculer visibilité des références

        /// <summary>
        /// Bascule la visibilité des éléments de référence (Plans, Axes, Points)
        /// Code converti depuis Toggle_RefVisibility.Vb
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

                    // Vérifier le type de document
                    int docType = doc.DocumentType;
                    const int kPartDocumentObject = 12288;
                    const int kAssemblyDocumentObject = 12290;
                    const int kDrawingDocumentObject = 12291;

                    if (docType != kPartDocumentObject && docType != kAssemblyDocumentObject)
                    {
                        string docTypeName = docType == kDrawingDocumentObject ? "Dessin" : "Autre";
                        Log($"Ce script fonctionne uniquement sur des fichiers de type Pièce ou Assemblage. Type actuel: {docTypeName}", "WARNING");
                        return; // Ne pas lever d'exception, juste logger
                    }

                    Log("Analyse de l'état actuel des références...", "INFO");
                    // Déterminer l'état actuel
                    bool currentState = AreAnyRefVisible(doc);
                    
                    Log($"État actuel: {(currentState ? "Visibles" : "Masquées")}. Basculement...", "INFO");
                    // Basculer la visibilité
                    ToggleRefVisibility(doc, !currentState);
                    Log("Visibilité des références basculée avec succès", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de l'exécution de ToggleRefVisibility: {ex.Message}", "ERROR");
                    throw new Exception($"Erreur lors de l'exécution de ToggleRefVisibility: {ex.Message}", ex);
                }
            });
        }

        private bool AreAnyRefVisible(dynamic doc)
        {
            try
            {
                int docType = doc.DocumentType;
                const int kPartDocumentObject = 12288;
                const int kAssemblyDocumentObject = 12290;

                if (docType == kPartDocumentObject)
                {
                    dynamic partDef = doc.ComponentDefinition;
                    
                    // Vérifier WorkPlanes
                    dynamic workPlanes = partDef.WorkPlanes;
                    foreach (dynamic wp in workPlanes)
                    {
                        if (wp.Visible) return true;
                    }
                    
                    // Vérifier WorkAxes
                    dynamic workAxes = partDef.WorkAxes;
                    foreach (dynamic wa in workAxes)
                    {
                        if (wa.Visible) return true;
                    }
                    
                    // Vérifier WorkPoints
                    dynamic workPoints = partDef.WorkPoints;
                    foreach (dynamic wpt in workPoints)
                    {
                        if (wpt.Visible) return true;
                    }
                }
                else if (docType == kAssemblyDocumentObject)
                {
                    dynamic asmDef = doc.ComponentDefinition;
                    dynamic occurrences = asmDef.Occurrences;
                    
                    foreach (dynamic occ in occurrences)
                    {
                        if (!occ.Suppressed && occ.Visible)
                        {
                            try
                            {
                                dynamic defDoc = occ.Definition.Document;
                                if (AreAnyRefVisible(defDoc))
                                {
                                    return true;
                                }
                            }
                            catch
                            {
                                // Continuer
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur dans la détection de visibilité : {ex.Message}", 
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return false;
        }

        private void ToggleRefVisibility(dynamic doc, bool visibleState)
        {
            try
            {
                int docType = doc.DocumentType;
                const int kPartDocumentObject = 12288;
                const int kAssemblyDocumentObject = 12290;

                if (docType == kPartDocumentObject)
                {
                    dynamic partDef = doc.ComponentDefinition;
                    
                    // WorkPlanes
                    dynamic workPlanes = partDef.WorkPlanes;
                    foreach (dynamic wp in workPlanes)
                    {
                        wp.Visible = visibleState;
                    }
                    
                    // WorkAxes
                    dynamic workAxes = partDef.WorkAxes;
                    foreach (dynamic wa in workAxes)
                    {
                        wa.Visible = visibleState;
                    }
                    
                    // WorkPoints
                    dynamic workPoints = partDef.WorkPoints;
                    foreach (dynamic wpt in workPoints)
                    {
                        wpt.Visible = visibleState;
                    }
                }
                else if (docType == kAssemblyDocumentObject)
                {
                    dynamic asmDef = doc.ComponentDefinition;
                    
                    // Éléments de référence du niveau d'assemblage
                    dynamic workPlanes = asmDef.WorkPlanes;
                    foreach (dynamic wp in workPlanes)
                    {
                        wp.Visible = visibleState;
                    }
                    
                    dynamic workAxes = asmDef.WorkAxes;
                    foreach (dynamic wa in workAxes)
                    {
                        wa.Visible = visibleState;
                    }
                    
                    dynamic workPoints = asmDef.WorkPoints;
                    foreach (dynamic wpt in workPoints)
                    {
                        wpt.Visible = visibleState;
                    }
                    
                    // Parcourir les occurrences
                    dynamic occurrences = asmDef.Occurrences;
                    foreach (dynamic occ in occurrences)
                    {
                        if (!occ.Suppressed && occ.Visible)
                        {
                            try
                            {
                                dynamic defDoc = occ.Definition.Document;
                                ToggleRefVisibility(defDoc, visibleState);
                            }
                            catch
                            {
                                // Continuer
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur dans le changement de visibilité : {ex.Message}", 
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region ToggleSketchVisibility - Basculer visibilité des esquisses

        /// <summary>
        /// Active ou désactive la visibilité de toutes les esquisses (2D et 3D)
        /// Code converti depuis Toggle_SketchVisibility.Vb
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

                    int docType = doc.DocumentType;
                    const int kPartDocumentObject = 12288;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType != kPartDocumentObject && docType != kAssemblyDocumentObject)
                    {
                        Log("Ce script fonctionne uniquement sur des fichiers .ipt ou .iam.", "WARNING");
                        return;
                    }

                    Log("Analyse des esquisses...", "INFO");
                    var processor = new SketchVisibilityProcessor();
                    processor.AnalyzeDocument(doc);
                    Log($"Esquisses trouvées: {processor.TotalCount} (Visibles: {processor.VisibleCount}, Masquées: {processor.HiddenCount})", "INFO");
                    processor.ToggleSketches(doc);
                    
                    // Mise à jour de la vue
                    dynamic activeView = inventorApp.ActiveView;
                    if (activeView != null)
                    {
                        activeView.Update();
                    }
                    Log("Visibilité des esquisses basculée avec succès", "SUCCESS");
                }
                catch (Exception ex)
                {
                    Log($"Erreur Toggle_SketchVisibility: {ex.Message}", "ERROR");
                }
            });
        }

        private class SketchVisibilityProcessor
        {
            public int VisibleCount { get; set; } = 0;
            public int HiddenCount { get; set; } = 0;
            public int TotalCount { get; set; } = 0;
            public bool TargetVisibility { get; set; } = true;

            public void AnalyzeDocument(dynamic doc)
            {
                try
                {
                    int docType = doc.DocumentType;
                    const int kPartDocumentObject = 12288;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType == kPartDocumentObject)
                    {
                        AnalyzePart(doc.ComponentDefinition);
                    }
                    else if (docType == kAssemblyDocumentObject)
                    {
                        AnalyzeAssembly(doc);
                    }

                    TotalCount = VisibleCount + HiddenCount;

                    if (TotalCount > 0)
                    {
                        TargetVisibility = (VisibleCount <= HiddenCount);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erreur analyse document: {ex.Message}");
                }
            }

            private void AnalyzePart(dynamic partDef)
            {
                try
                {
                    // Esquisses 2D
                    dynamic sketches = partDef.Sketches;
                    foreach (dynamic sketch in sketches)
                    {
                        if (sketch.Visible)
                        {
                            VisibleCount++;
                        }
                        else
                        {
                            HiddenCount++;
                        }
                    }

                    // Esquisses 3D
                    try
                    {
                        dynamic sketches3D = partDef.Sketches3D;
                        if (sketches3D != null)
                        {
                            foreach (dynamic sketch3d in sketches3D)
                            {
                                if (sketch3d.Visible)
                                {
                                    VisibleCount++;
                                }
                                else
                                {
                                    HiddenCount++;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Pas d'esquisses 3D
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erreur analyse pièce: {ex.Message}");
                }
            }

            private void AnalyzeAssembly(dynamic doc)
            {
                try
                {
                    AnalyzePart(doc.ComponentDefinition);

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType == kAssemblyDocumentObject)
                    {
                        dynamic asmDef = doc.ComponentDefinition;
                        dynamic occurrences = asmDef.Occurrences;

                        foreach (dynamic occ in occurrences)
                        {
                            try
                            {
                                AnalyzePart(occ.Definition);

                                // Traitement récursif pour les sous-assemblages
                                try
                                {
                                    int defDocType = occ.DefinitionDocumentType;
                                    if (defDocType == kAssemblyDocumentObject)
                                    {
                                        dynamic subOccurrences = occ.SubOccurrences;
                                        if (subOccurrences != null && subOccurrences.Count > 0)
                                        {
                                            AnalyzeSubOccurrences(subOccurrences);
                                        }
                                    }
                                }
                                catch
                                {
                                    // Pas de sous-occurrences
                                }
                            }
                            catch
                            {
                                // Continuer
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erreur analyse assemblage: {ex.Message}");
                }
            }

            private void AnalyzeSubOccurrences(dynamic occurrences)
            {
                try
                {
                    foreach (dynamic occ in occurrences)
                    {
                        try
                        {
                            if (occ != null && occ.IsActive)
                            {
                                AnalyzePart(occ.Definition);

                                const int kAssemblyDocumentObject = 12290;
                                try
                                {
                                    int defDocType = occ.DefinitionDocumentType;
                                    if (defDocType == kAssemblyDocumentObject)
                                    {
                                        dynamic subOccurrences = occ.SubOccurrences;
                                        if (subOccurrences != null && subOccurrences.Count > 0)
                                        {
                                            AnalyzeSubOccurrences(subOccurrences);
                                        }
                                    }
                                }
                                catch
                                {
                                    // Pas de sous-occurrences
                                }
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

            public void ToggleSketches(dynamic doc)
            {
                try
                {
                    int docType = doc.DocumentType;
                    const int kPartDocumentObject = 12288;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType == kPartDocumentObject)
                    {
                        ApplyVisibilityToPart(doc.ComponentDefinition, TargetVisibility);
                    }
                    else if (docType == kAssemblyDocumentObject)
                    {
                        ApplyVisibilityToAssembly(doc, TargetVisibility);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erreur toggle esquisses: {ex.Message}");
                }
            }

            private void ApplyVisibilityToPart(dynamic partDef, bool newState)
            {
                try
                {
                    // Esquisses 2D
                    dynamic sketches = partDef.Sketches;
                    foreach (dynamic sketch in sketches)
                    {
                        try
                        {
                            sketch.Visible = newState;
                        }
                        catch
                        {
                            // Continuer
                        }
                    }

                    // Esquisses 3D
                    try
                    {
                        dynamic sketches3D = partDef.Sketches3D;
                        if (sketches3D != null)
                        {
                            foreach (dynamic sketch3d in sketches3D)
                            {
                                try
                                {
                                    sketch3d.Visible = newState;
                                }
                                catch
                                {
                                    // Continuer
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Pas d'esquisses 3D
                    }
                }
                catch
                {
                    // Continuer
                }
            }

            private void ApplyVisibilityToAssembly(dynamic doc, bool newState)
            {
                try
                {
                    ApplyVisibilityToPart(doc.ComponentDefinition, newState);

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;

                    if (docType == kAssemblyDocumentObject)
                    {
                        dynamic asmDef = doc.ComponentDefinition;
                        dynamic occurrences = asmDef.Occurrences;

                        foreach (dynamic occ in occurrences)
                        {
                            try
                            {
                                ApplyVisibilityToPart(occ.Definition, newState);

                                // Traitement récursif
                                try
                                {
                                    int defDocType = occ.DefinitionDocumentType;
                                    if (defDocType == kAssemblyDocumentObject)
                                    {
                                        dynamic subOccurrences = occ.SubOccurrences;
                                        if (subOccurrences != null && subOccurrences.Count > 0)
                                        {
                                            ApplyVisibilityToSubOccurrences(subOccurrences, newState);
                                        }
                                    }
                                }
                                catch
                                {
                                    // Pas de sous-occurrences
                                }
                            }
                            catch
                            {
                                // Continuer
                            }
                        }
                    }
                }
                catch
                {
                    // Continuer
                }
            }

            private void ApplyVisibilityToSubOccurrences(dynamic occurrences, bool newState)
            {
                try
                {
                    foreach (dynamic occ in occurrences)
                    {
                        try
                        {
                            if (occ != null && occ.IsActive)
                            {
                                ApplyVisibilityToPart(occ.Definition, newState);

                                const int kAssemblyDocumentObject = 12290;
                                try
                                {
                                    int defDocType = occ.DefinitionDocumentType;
                                    if (defDocType == kAssemblyDocumentObject)
                                    {
                                        dynamic subOccurrences = occ.SubOccurrences;
                                        if (subOccurrences != null && subOccurrences.Count > 0)
                                        {
                                            ApplyVisibilityToSubOccurrences(subOccurrences, newState);
                                        }
                                    }
                                }
                                catch
                                {
                                    // Pas de sous-occurrences
                                }
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

        #endregion

        public async Task ExecuteConstraintReportAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir AssemblyConstraintStatusReport.Vb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteIPropertiesSummaryAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir iPropertiesSummary.Vb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        #region SmartSave - Sauvegarde intelligente

        /// <summary>
        /// Sauvegarde intelligente des documents Inventor sans les fermer
        /// Code converti depuis SmartSave.iLogicVb
        /// </summary>
        public async Task ExecuteSmartSaveAsync(Action<string, string>? logCallback = null)
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
                        MessageBox.Show("⚠️ Aucun document actif n'est ouvert.", "Erreur", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;
                    const int kPartDocumentObject = 12288;
                    const int kDrawingDocumentObject = 12291;

                    if (docType == kAssemblyDocumentObject)
                    {
                        ExecuteAssemblySteps(doc, inventorApp);
                    }
                    else if (docType == kPartDocumentObject)
                    {
                        ExecutePartSteps(doc, inventorApp);
                    }
                    else if (docType == kDrawingDocumentObject)
                    {
                        ExecuteDrawingSteps(doc, inventorApp);
                    }
                    else
                    {
                        ExecuteGenericSteps(doc, inventorApp);
                    }
                }
                catch (Exception ex)
                {
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
                    dynamic cam = activeView.Camera;
                    const int kIsoTopRightViewOrientation = 4;
                    cam.ViewOrientationType = kIsoTopRightViewOrientation;
                    cam.Perspective = false;
                    cam.Fit();
                    cam.Apply();
                    activeView.Update();
                }
            }
            catch
            {
                // Continuer
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
        /// </summary>
        public async Task ExecuteSafeCloseAsync(Action<string, string>? logCallback = null)
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
                        MessageBox.Show("Aucun document actif n'est ouvert.", "Erreur", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    int docType = doc.DocumentType;
                    const int kAssemblyDocumentObject = 12290;
                    const int kPartDocumentObject = 12288;
                    const int kDrawingDocumentObject = 12291;

                    if (docType == kAssemblyDocumentObject)
                    {
                        ExecuteAssemblyStepsForClose(doc, inventorApp);
                    }
                    else if (docType == kPartDocumentObject)
                    {
                        ExecutePartStepsForClose(doc, inventorApp);
                    }
                    else if (docType == kDrawingDocumentObject)
                    {
                        ExecuteDrawingStepsForClose(doc, inventorApp);
                    }
                    else
                    {
                        ExecuteGenericStepsForClose(doc, inventorApp);
                    }
                }
                catch (Exception ex)
                {
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

        public async Task ExecuteExportIAMToIPTAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir ExportIAMToIPT en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteExportIDWtoShopPDFAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir ExportIDWtoShopPDF.iLogicVb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteFormCenteringUtilityAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir FormCenteringUtility en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteILogicFormsCentredAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir iLogicFormsCentred_Rule.vb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteInfoCommandAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
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

        public async Task ExecuteAutoSaveAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir AutoSave.iLogicVb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteCheckSaveStatusAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir CheckSaveStatus.iLogicVb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

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

        public async Task ExecuteInsertSpecificScrewInHolesAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir InsertSpecificScrewInHoles.vb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteIPropertyCustomBatchAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir iPropertyCustomBatch.iLogicVb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteFixPromotedVariesAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir FixPromotedVaries.iLogicVb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteOpenSelectedComponentAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir OpenSelectedComponen.Vb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteUpdateCurrentSheetAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir UpdateCurrentSheet.Vb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }

        public async Task ExecuteHideBoxTemplateMultiOpeningAsync(Action<string, string>? logCallback = null)
        {
            if (logCallback != null)
                SetLogCallback(logCallback);

            await Task.Run(() =>
            {
                // TODO: Convertir HideBoxTemplateMultiOpening.Vb en C#
                Log("Fonctionnalité non encore implémentée", "WARNING");
                throw new NotImplementedException("À implémenter");
            });
        }
    }
}
