#nullable enable
using System;
using System.Runtime.InteropServices;
using Inventor;

using XnrgyEngineeringAutomationTools.Services;

namespace VaultAutomationTool.Services
{
    /// <summary>
    /// Service pour modifier les iProperties des fichiers Inventor AVANT upload vers Vault.
    /// Vault synchronise automatiquement les iProperties vers les UDP lors du check-in.
    /// </summary>
    public class InventorPropertyService : IDisposable
    {
        private Application? _inventorApp;
        private bool _wasAlreadyRunning = false;
        private bool _disposed = false;

        /// <summary>
        /// Initialise la connexion à Inventor (ou démarre une instance invisible).
        /// </summary>
        public bool Initialize()
        {
            try
            {
                Logger.Log("   [INVENTOR-API] Connexion à Inventor...", Logger.LogLevel.DEBUG);
                
                // Essayer de se connecter à une instance existante
                try
                {
                    _inventorApp = (Application)Marshal.GetActiveObject("Inventor.Application");
                    _wasAlreadyRunning = true;
                    Logger.Log("   [OK] Connecté à instance Inventor existante", Logger.LogLevel.DEBUG);
                }
                catch (COMException)
                {
                    // Pas d'instance existante, démarrer une nouvelle instance invisible
                    Logger.Log("   [INFO] Aucune instance Inventor trouvée, démarrage en mode invisible...", Logger.LogLevel.DEBUG);
                    
                    Type inventorType = Type.GetTypeFromProgID("Inventor.Application")!;
                    _inventorApp = (Application)Activator.CreateInstance(inventorType)!;
                    _inventorApp.Visible = false;
                    _wasAlreadyRunning = false;
                    
                    Logger.Log("   [OK] Instance Inventor démarrée (invisible)", Logger.LogLevel.DEBUG);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"   [ERROR] Impossible d'initialiser Inventor: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Définit les iProperties personnalisées (Project, Reference, Module) sur un fichier Inventor.
        /// </summary>
        /// <param name="filePath">Chemin complet du fichier .ipt/.iam/.idw/.ipn</param>
        /// <param name="projectNumber">Valeur pour la propriété "Project"</param>
        /// <param name="reference">Valeur pour la propriété "Reference"</param>
        /// <param name="module">Valeur pour la propriété "Module"</param>
        /// <returns>True si succès</returns>
        public bool SetIProperties(string filePath, string? projectNumber, string? reference, string? module)
        {
            if (_inventorApp == null)
            {
                Logger.Log("   [ERROR] Inventor non initialisé", Logger.LogLevel.ERROR);
                return false;
            }

            Document? doc = null;
            try
            {
                Logger.Log($"   [INVENTOR-API] Ouverture: {System.IO.Path.GetFileName(filePath)}", Logger.LogLevel.DEBUG);
                
                // Ouvrir le document en mode invisible (pas de fenêtre visible)
                doc = _inventorApp.Documents.Open(filePath, false);
                
                // Accéder aux PropertySets
                PropertySets propSets = doc.PropertySets;
                
                // Trouver ou créer le PropertySet "Inventor User Defined Properties"
                PropertySet customProps;
                try
                {
                    customProps = propSets["Inventor User Defined Properties"];
                }
                catch
                {
                    // Si n'existe pas, utiliser le premier disponible (rare)
                    customProps = propSets["Design Tracking Properties"];
                }

                // Définir les propriétés
                SetOrCreateProperty(customProps, "Project", projectNumber);
                SetOrCreateProperty(customProps, "Reference", reference);
                SetOrCreateProperty(customProps, "Module", module);

                // Sauvegarder le document
                doc.Save();
                Logger.Log($"   [OK] iProperties définies et fichier sauvegardé", Logger.LogLevel.INFO);
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"   [ERROR] Erreur iProperties: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
            finally
            {
                // Fermer le document sans sauvegarder à nouveau
                if (doc != null)
                {
                    try
                    {
                        doc.Close(true); // true = skip save (déjà fait)
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Définit toutes les iProperties personnalisées du module XNRGY.
        /// </summary>
        /// <param name="filePath">Chemin complet du fichier .ipt/.iam/.idw/.ipn</param>
        /// <param name="project">Numéro de projet (ex: "25001")</param>
        /// <param name="reference">Référence (ex: "REF1")</param>
        /// <param name="module">Module (ex: "M1")</param>
        /// <param name="initialeDessinateur">Initiales du dessinateur (ex: "MAE")</param>
        /// <param name="initialeCoDessinateur">Initiales du co-dessinateur (optionnel)</param>
        /// <param name="creationDate">Date de création</param>
        /// <returns>True si succès</returns>
        public bool SetAllModuleProperties(string filePath, string? project, string? reference, string? module,
            string? initialeDessinateur, string? initialeCoDessinateur, DateTime? creationDate)
        {
            if (_inventorApp == null)
            {
                Logger.Log("   [ERROR] Inventor non initialisé", Logger.LogLevel.ERROR);
                return false;
            }

            Document? doc = null;
            try
            {
                Logger.Log($"   [INVENTOR-API] Ouverture: {System.IO.Path.GetFileName(filePath)}", Logger.LogLevel.DEBUG);
                
                // Ouvrir le document en mode invisible
                doc = _inventorApp.Documents.Open(filePath, false);
                
                // Accéder aux PropertySets
                PropertySets propSets = doc.PropertySets;
                
                // Trouver le PropertySet "Inventor User Defined Properties"
                PropertySet customProps;
                try
                {
                    customProps = propSets["Inventor User Defined Properties"];
                }
                catch
                {
                    customProps = propSets["Design Tracking Properties"];
                }

                // Définir les propriétés de base
                SetOrCreateProperty(customProps, "Project", project);
                SetOrCreateProperty(customProps, "Reference", reference);
                SetOrCreateProperty(customProps, "Module", module);
                
                // Propriétés supplémentaires du module
                SetOrCreateProperty(customProps, "Initiale_du_Dessinateur", initialeDessinateur);
                SetOrCreateProperty(customProps, "Initiale_du_Co_Dessinateur", initialeCoDessinateur);
                
                if (creationDate.HasValue)
                {
                    SetOrCreateProperty(customProps, "Creation_Date", creationDate.Value.ToString("yyyy-MM-dd"));
                }

                // Construire le numéro de projet complet (format: 25001REF1M1)
                var fullProjectNumber = $"{project}{reference}{module}";
                SetOrCreateProperty(customProps, "Numero_de_Projet", fullProjectNumber);

                // Sauvegarder le document
                doc.Save();
                Logger.Log($"   [OK] Toutes les iProperties définies et fichier sauvegardé", Logger.LogLevel.INFO);
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"   [ERROR] Erreur iProperties: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(true); } catch { }
                }
            }
        }

        /// <summary>
        /// Définit ou crée une propriété personnalisée.
        /// </summary>
        private void SetOrCreateProperty(PropertySet propSet, string propName, string? value)
        {
            if (string.IsNullOrEmpty(value)) return;

            try
            {
                // Essayer de trouver la propriété existante
                Property prop = propSet[propName];
                prop.Value = value;
                Logger.Log($"      [SET] {propName} = '{value}'", Logger.LogLevel.TRACE);
            }
            catch
            {
                // La propriété n'existe pas, la créer
                try
                {
                    propSet.Add(value, propName);
                    Logger.Log($"      [NEW] {propName} = '{value}'", Logger.LogLevel.TRACE);
                }
                catch (Exception ex)
                {
                    Logger.Log($"      [WARN] Impossible de créer {propName}: {ex.Message}", Logger.LogLevel.WARNING);
                }
            }
        }

        /// <summary>
        /// Libère les ressources Inventor.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_inventorApp != null && !_wasAlreadyRunning)
                {
                    // Fermer l'instance qu'on a démarrée
                    Logger.Log("   [INVENTOR-API] Fermeture de l'instance Inventor...", Logger.LogLevel.DEBUG);
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
                Logger.Log($"   [WARN] Erreur fermeture Inventor: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }
    }
}

