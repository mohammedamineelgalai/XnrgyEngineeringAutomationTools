using System;
using System.Runtime.InteropServices;
using System.Threading;

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

        // P/Invoke alternatif via ole32
        [DllImport("ole32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, out IntPtr ppunk);

        [DllImport("ole32.dll")]
        private static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

        [DllImport("ole32.dll")]
        private static extern void CoUninitialize();

        private const uint COINIT_MULTITHREADED = 0x0;
        private const uint COINIT_APARTMENTTHREADED = 0x2;

        public bool IsConnected => _isConnected && _inventorApp != null;

        /// <summary>
        /// Tente de se connecter à une instance Inventor en cours d'exécution
        /// Utilise plusieurs méthodes et retries pour maximiser les chances de connexion
        /// </summary>
        public bool TryConnect()
        {
            Logger.Log("[>] Tentative de connexion a Inventor...", Logger.LogLevel.DEBUG);

            // Essayer plusieurs méthodes de connexion
            for (int retry = 0; retry < 3; retry++)
            {
                if (retry > 0)
                {
                    Logger.Log($"[>] Retry connexion COM ({retry + 1}/3)...", Logger.LogLevel.DEBUG);
                    Thread.Sleep(1000);
                }

                // Méthode 1: Marshal.GetActiveObject (classique .NET Framework)
                if (TryConnectViaMarshall())
                {
                    return true;
                }

                // Méthode 2: P/Invoke GetActiveObject via oleaut32
                if (TryConnectViaPInvoke())
                {
                    return true;
                }

                // Méthode 3: Type.GetTypeFromProgID + Activator (pour instance existante)
                if (TryConnectViaActivator())
                {
                    return true;
                }
            }

            _isConnected = false;
            return false;
        }

        /// <summary>
        /// Méthode 1: Connexion via Marshal.GetActiveObject
        /// </summary>
        private bool TryConnectViaMarshall()
        {
            try
            {
                _inventorApp = Marshal.GetActiveObject("Inventor.Application");
                if (_inventorApp != null)
                {
                    _isConnected = true;
                    Logger.Log($"[+] Connecte a Inventor via Marshal.GetActiveObject", Logger.LogLevel.INFO);
                    return true;
                }
            }
            catch (COMException comEx)
            {
                Logger.Log($"[!] Marshal.GetActiveObject echoue: 0x{comEx.ErrorCode:X8}", Logger.LogLevel.DEBUG);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Marshal.GetActiveObject erreur: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return false;
        }

        /// <summary>
        /// Méthode 2: Connexion via P/Invoke oleaut32.GetActiveObject
        /// </summary>
        private bool TryConnectViaPInvoke()
        {
            try
            {
                Guid clsid;
                int hr = CLSIDFromProgID("Inventor.Application", out clsid);
                
                if (hr == 0)
                {
                    object inventorObj;
                    GetActiveObject(ref clsid, IntPtr.Zero, out inventorObj);
                    
                    if (inventorObj != null)
                    {
                        _inventorApp = inventorObj;
                        _isConnected = true;
                        Logger.Log($"[+] Connecte a Inventor via P/Invoke", Logger.LogLevel.INFO);
                        return true;
                    }
                }
            }
            catch (COMException comEx)
            {
                Logger.Log($"[!] P/Invoke GetActiveObject echoue: 0x{comEx.ErrorCode:X8}", Logger.LogLevel.DEBUG);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] P/Invoke erreur: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return false;
        }

        /// <summary>
        /// Méthode 3: Connexion via Type.GetTypeFromProgID + test si instance existe
        /// </summary>
        private bool TryConnectViaActivator()
        {
            try
            {
                // GetTypeFromProgID avec le flag pour obtenir une instance existante
                Type? inventorType = Type.GetTypeFromProgID("Inventor.Application", false);
                
                if (inventorType != null)
                {
                    // Essayer d'obtenir l'instance existante via ROT (Running Object Table)
                    object? inventorObj = null;
                    
                    try
                    {
                        // Cette méthode peut fonctionner dans certains contextes COM
                        inventorObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                    }
                    catch
                    {
                        // Si échec, on ne crée PAS de nouvelle instance (on veut une existante)
                        return false;
                    }
                    
                    if (inventorObj != null)
                    {
                        _inventorApp = inventorObj;
                        _isConnected = true;
                        Logger.Log($"[+] Connecte a Inventor via Activator/ROT", Logger.LogLevel.INFO);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Activator erreur: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return false;
        }

        /// <summary>
        /// Force la réinitialisation de la connexion COM et réessaie
        /// Utile quand l'app est lancée depuis un script
        /// </summary>
        public bool ForceReconnect()
        {
            Logger.Log("[>] Force reconnexion COM...", Logger.LogLevel.INFO);
            
            // Libérer la connexion existante
            Disconnect();
            
            // Attendre un peu pour que COM se stabilise
            Thread.Sleep(500);
            
            // Réessayer avec plus de tentatives
            for (int i = 0; i < 5; i++)
            {
                if (TryConnect())
                {
                    return true;
                }
                Thread.Sleep(1000);
            }
            
            return false;
        }

        /// <summary>
        /// Déconnecte proprement de l'instance Inventor
        /// </summary>
        public void Disconnect()
        {
            try
            {
                if (_inventorApp != null)
                {
                    Marshal.ReleaseComObject(_inventorApp);
                    _inventorApp = null;
                }
            }
            catch { }
            
            _isConnected = false;
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
    }
}

