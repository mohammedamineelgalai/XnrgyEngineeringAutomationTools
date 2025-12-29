using System;
using System.Runtime.InteropServices;

using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service pour la connexion et l'interaction avec Autodesk Inventor 2026.2
    /// Utilise COM Interop pour piloter Inventor
    /// </summary>
    public class InventorService
    {
        private dynamic? _inventorApp;
        private bool _isConnected = false;

        // P/Invoke pour COM
        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public bool IsConnected => _isConnected && _inventorApp != null;

        /// <summary>
        /// Tente de se connecter à une instance Inventor en cours d'exécution
        /// </summary>
        public bool TryConnect()
        {
            try
            {
                Logger.Log("[>] Tentative de connexion à Inventor...", Logger.LogLevel.DEBUG);

                // Méthode 1: GetActiveObject via P/Invoke
                Guid clsid;
                int hr = CLSIDFromProgID("Inventor.Application", out clsid);
                
                if (hr == 0)
                {
                    object inventorObj;
                    GetActiveObject(ref clsid, IntPtr.Zero, out inventorObj);
                    _inventorApp = inventorObj;
                    _isConnected = true;
                    
                    Logger.Log($"[+] Connecté à Inventor via COM", Logger.LogLevel.INFO);
                    return true;
                }
            }
            catch (COMException comEx)
            {
                Logger.Log($"[!] Inventor non disponible (COM): {comEx.Message}", Logger.LogLevel.DEBUG);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur connexion Inventor: {ex.Message}", Logger.LogLevel.DEBUG);
            }

            _isConnected = false;
            return false;
        }

        /// <summary>
        /// Obtient le nom du document actif dans Inventor
        /// </summary>
        public string? GetActiveDocumentName()
        {
            try
            {
                if (_inventorApp != null)
                {
                    dynamic activeDoc = _inventorApp.ActiveDocument;
                    if (activeDoc != null)
                    {
                        return activeDoc.DisplayName;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur récupération document actif: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return null;
        }

        /// <summary>
        /// Obtient le chemin complet du document actif
        /// </summary>
        public string? GetActiveDocumentPath()
        {
            try
            {
                if (_inventorApp != null)
                {
                    dynamic activeDoc = _inventorApp.ActiveDocument;
                    if (activeDoc != null)
                    {
                        return activeDoc.FullFileName;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur récupération chemin document: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return null;
        }

        /// <summary>
        /// Obtient la version d'Inventor
        /// </summary>
        public string? GetInventorVersion()
        {
            try
            {
                if (_inventorApp != null)
                {
                    return _inventorApp.SoftwareVersion.DisplayVersion;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur récupération version: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return null;
        }

        /// <summary>
        /// Vérifie si un document est ouvert
        /// </summary>
        public bool HasActiveDocument()
        {
            try
            {
                if (_inventorApp != null)
                {
                    return _inventorApp.ActiveDocument != null;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Obtient le type de document actif (Part, Assembly, Drawing, Presentation)
        /// </summary>
        public string? GetActiveDocumentType()
        {
            try
            {
                if (_inventorApp != null)
                {
                    dynamic activeDoc = _inventorApp.ActiveDocument;
                    if (activeDoc != null)
                    {
                        int docType = activeDoc.DocumentType;
                        return docType switch
                        {
                            12291 => "Part",        // kPartDocumentObject
                            12290 => "Assembly",    // kAssemblyDocumentObject
                            12292 => "Drawing",     // kDrawingDocumentObject
                            12293 => "Presentation", // kPresentationDocumentObject
                            _ => "Unknown"
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur récupération type document: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return null;
        }

        /// <summary>
        /// Déconnecte d'Inventor (libère la référence COM)
        /// </summary>
        public void Disconnect()
        {
            if (_inventorApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_inventorApp);
                }
                catch { }
                _inventorApp = null;
            }
            _isConnected = false;
            Logger.Log("[>] Déconnexion d'Inventor", Logger.LogLevel.DEBUG);
        }
    }
}

