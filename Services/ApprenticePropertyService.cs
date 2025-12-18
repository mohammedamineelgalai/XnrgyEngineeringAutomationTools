using System;
using System.IO;
using System.Runtime.InteropServices;

#nullable enable

using XnrgyEngineeringAutomationTools.Services;

namespace VaultAutomationTool.Services
{
    /// <summary>
    /// Service pour écrire les iProperties Inventor via Inventor Application.
    /// Utilise Inventor.Application en mode invisible.
    /// Temps estimé: ~1-3 secondes par fichier (après initialisation initiale ~5-10s)
    /// </summary>
    public class ApprenticePropertyService : IDisposable
    {
        private dynamic? _inventorApp = null;
        private bool _disposed = false;
        private static readonly object _lock = new object();

        /// <summary>
        /// Initialise Inventor Application (une seule fois, réutilisable pour plusieurs fichiers)
        /// </summary>
        public bool Initialize()
        {
            if (_inventorApp != null)
                return true;

            lock (_lock)
            {
                if (_inventorApp != null)
                    return true;

                try
                {
                    var startTime = DateTime.Now;

                    // Utiliser Inventor.Application
                    Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType == null)
                    {
                        Logger.Log("[INVENTOR] ❌ Inventor.Application non disponible", Logger.LogLevel.ERROR);
                        return false;
                    }

                    Logger.Log("[INVENTOR] Initialisation d'Inventor Application...", Logger.LogLevel.INFO);
                    
                    // Essayer de se connecter à une instance existante
                    try
                    {
                        _inventorApp = Marshal.GetActiveObject("Inventor.Application");
                        Logger.Log("[INVENTOR] ✅ Connecté à Inventor existant", Logger.LogLevel.INFO);
                    }
                    catch
                    {
                        // Créer une nouvelle instance
                        _inventorApp = Activator.CreateInstance(inventorType);
                        if (_inventorApp != null)
                        {
                            _inventorApp.Visible = false; // Mode invisible
                            Logger.Log("[INVENTOR] ✅ Nouvelle instance créée (invisible)", Logger.LogLevel.INFO);
                        }
                    }

                    if (_inventorApp != null)
                    {
                        var elapsed = (DateTime.Now - startTime).TotalMilliseconds;
                        Logger.Log($"[INVENTOR] ✅ Prêt en {elapsed:F0}ms", Logger.LogLevel.INFO);
                        return true;
                    }

                    Logger.Log("[INVENTOR] ❌ Impossible d'initialiser Inventor", Logger.LogLevel.ERROR);
                    return false;
                }
                catch (Exception ex)
                {
                    Logger.Log($"[INVENTOR] ❌ Erreur initialisation: {ex.Message}", Logger.LogLevel.ERROR);
                    return false;
                }
            }
        }

        /// <summary>
        /// Écrit les iProperties Custom (Project, Reference, Module) dans un fichier Inventor.
        /// </summary>
        /// <param name="filePath">Chemin complet du fichier .ipt, .iam, .idw ou .ipn</param>
        /// <param name="project">Valeur de Project</param>
        /// <param name="reference">Valeur de Reference</param>
        /// <param name="module">Valeur de Module</param>
        /// <returns>True si succès</returns>
        public bool SetIProperties(string filePath, string? project, string? reference, string? module)
        {
            if (_inventorApp == null)
            {
                if (!Initialize())
                    return false;
            }

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Logger.Log($"[INVENTOR] Fichier non trouvé: {filePath}", Logger.LogLevel.ERROR);
                return false;
            }

            // Vérifier l'extension
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".ipt" && ext != ".iam" && ext != ".idw" && ext != ".ipn")
            {
                Logger.Log($"[INVENTOR] Extension non supportée: {ext}", Logger.LogLevel.WARNING);
                return false;
            }

            dynamic? document = null;
            var startTime = DateTime.Now;

            try
            {
                // Enlever l'attribut ReadOnly si nécessaire
                var fileInfo = new FileInfo(filePath);
                bool wasReadOnly = fileInfo.IsReadOnly;
                if (wasReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                    Logger.Log($"[INVENTOR] ReadOnly désactivé temporairement", Logger.LogLevel.DEBUG);
                }

                // Ouvrir le document avec Inventor
                Logger.Log($"[INVENTOR] Ouverture: {Path.GetFileName(filePath)}", Logger.LogLevel.DEBUG);
                document = _inventorApp!.Documents.Open(filePath, false); // false = ne pas activer

                if (document == null)
                {
                    Logger.Log($"[INVENTOR] ❌ Impossible d'ouvrir le document", Logger.LogLevel.ERROR);
                    return false;
                }

                // Accéder aux PropertySets
                dynamic propertySets = document.PropertySets;
                dynamic? customProps = null;

                // Trouver le PropertySet "Inventor User Defined Properties"
                foreach (dynamic propSet in propertySets)
                {
                    if (propSet.Name == "Inventor User Defined Properties")
                    {
                        customProps = propSet;
                        break;
                    }
                }

                if (customProps == null)
                {
                    Logger.Log($"[INVENTOR] ❌ PropertySet 'Inventor User Defined Properties' non trouvé", Logger.LogLevel.ERROR);
                    document.Close(true); // true = skip save
                    return false;
                }

                int propsSet = 0;

                // Définir Project
                if (!string.IsNullOrEmpty(project))
                {
                    SetOrCreateProperty(customProps, "Project", project);
                    propsSet++;
                }

                // Définir Reference
                if (!string.IsNullOrEmpty(reference))
                {
                    SetOrCreateProperty(customProps, "Reference", reference);
                    propsSet++;
                }

                // Définir Module
                if (!string.IsNullOrEmpty(module))
                {
                    SetOrCreateProperty(customProps, "Module", module);
                    propsSet++;
                }

                // Sauvegarder et fermer
                document.Save();
                document.Close();
                document = null;

                // Restaurer ReadOnly si nécessaire
                if (wasReadOnly)
                {
                    fileInfo.IsReadOnly = true;
                }

                var elapsed = (DateTime.Now - startTime).TotalMilliseconds;
                Logger.Log($"[INVENTOR] ✅ {propsSet} iProperties écrites en {elapsed:F0}ms: {Path.GetFileName(filePath)}", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[INVENTOR] ❌ Erreur: {ex.Message}", Logger.LogLevel.ERROR);

                try
                {
                    if (document != null)
                    {
                        document.Close(true); // true = skip save
                    }
                }
                catch { }

                return false;
            }
        }

        /// <summary>
        /// Définit ou crée une propriété custom
        /// </summary>
        private void SetOrCreateProperty(dynamic customProps, string propName, string value)
        {
            try
            {
                // Essayer de trouver la propriété existante
                foreach (dynamic prop in customProps)
                {
                    if (prop.Name == propName)
                    {
                        prop.Value = value;
                        Logger.Log($"[INVENTOR] Propriété '{propName}' = {value}", Logger.LogLevel.DEBUG);
                        return;
                    }
                }

                // Si pas trouvée, la créer
                customProps.Add(value, propName);
                Logger.Log($"[INVENTOR] Propriété '{propName}' créée = {value}", Logger.LogLevel.DEBUG);
            }
            catch (Exception ex)
            {
                Logger.Log($"[INVENTOR] Erreur propriété '{propName}': {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;

            if (_inventorApp != null)
            {
                try
                {
                    // Ne pas quitter Inventor si c'était une instance existante
                    // _inventorApp.Quit(); // Commenté pour ne pas fermer Inventor
                    Marshal.ReleaseComObject(_inventorApp);
                }
                catch { }
                _inventorApp = null;
            }

            Logger.Log("[INVENTOR] Service libéré", Logger.LogLevel.DEBUG);
        }
    }
}

