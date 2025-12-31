using System;
using System.Diagnostics;
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
        private DateTime _lastConnectionAttempt = DateTime.MinValue;
        private int _consecutiveFailures = 0;

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
        
        // Constantes pour le throttling
        private const int MIN_RETRY_INTERVAL_MS = 2000; // Minimum 2 sec entre tentatives
        private const int MAX_CONSECUTIVE_FAILURES = 10; // Apres 10 echecs, augmenter l'intervalle

        public bool IsConnected => _isConnected && _inventorApp != null;

        /// <summary>
        /// Tente de se connecter à une instance Inventor en cours d'exécution
        /// Utilise plusieurs méthodes et retries pour maximiser les chances de connexion
        /// Implemente un throttling pour eviter de spammer les logs
        /// </summary>
        public bool TryConnect()
        {
            // Throttling: eviter les tentatives trop rapprochees
            var now = DateTime.Now;
            var elapsed = (now - _lastConnectionAttempt).TotalMilliseconds;
            
            if (elapsed < MIN_RETRY_INTERVAL_MS && _lastConnectionAttempt != DateTime.MinValue)
            {
                return false; // Trop tot, skip silencieusement
            }
            
            _lastConnectionAttempt = now;
            
            // Verifier d'abord si le processus Inventor existe
            var inventorProcesses = Process.GetProcessesByName("Inventor");
            if (inventorProcesses.Length == 0)
            {
                return false; // Pas de processus, pas de log
            }
            
            // Verifier si le processus principal est pret (pas juste le splash screen)
            bool processReady = false;
            foreach (var proc in inventorProcesses)
            {
                try
                {
                    // MainWindowHandle > 0 indique que la fenetre principale est creee
                    if (proc.MainWindowHandle != IntPtr.Zero && !string.IsNullOrEmpty(proc.MainWindowTitle))
                    {
                        processReady = true;
                        break;
                    }
                }
                catch { }
            }
            
            if (!processReady)
            {
                // Inventor demarre mais fenetre pas encore prete
                if (_consecutiveFailures == 0)
                {
                    Logger.Log("[~] Inventor en cours de demarrage (fenetre pas prete)...", Logger.LogLevel.DEBUG);
                }
                return false;
            }
            
            // Log uniquement si c'est une nouvelle serie de tentatives
            if (_consecutiveFailures == 0)
            {
                Logger.Log("[>] Tentative de connexion a Inventor...", Logger.LogLevel.DEBUG);
            }

            // Essayer une seule fois chaque methode (pas de retry interne - le timer gere ca)
            // Methode 1: Marshal.GetActiveObject (classique .NET Framework)
            if (TryConnectViaMarshall())
            {
                _consecutiveFailures = 0;
                return true;
            }

            // Methode 2: P/Invoke GetActiveObject via oleaut32
            if (TryConnectViaPInvoke())
            {
                _consecutiveFailures = 0;
                return true;
            }

            _consecutiveFailures++;
            _isConnected = false;
            
            // Log periodique pour indiquer que ca continue
            if (_consecutiveFailures % 5 == 0)
            {
                Logger.Log($"[~] Connexion Inventor: {_consecutiveFailures} tentatives, en attente du ROT...", Logger.LogLevel.DEBUG);
            }
            
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
            catch (COMException)
            {
                // Silencieux - gere par le throttling dans TryConnect
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
            catch (COMException)
            {
                // Silencieux - gere par le throttling
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] P/Invoke erreur: {ex.Message}", Logger.LogLevel.DEBUG);
            }
            return false;
        }

        /// <summary>
        /// Force la réinitialisation de la connexion COM et réessaie
        /// Utilise la meme logique que TryConnect sans throttling
        /// </summary>
        public bool ForceReconnect()
        {
            // Libérer la connexion existante
            Disconnect();
            
            // Reset le compteur de failures pour que TryConnect log si besoin
            _consecutiveFailures = 0;
            _lastConnectionAttempt = DateTime.MinValue;
            
            // Une seule tentative - le timer rappellera si besoin
            return TryConnect();
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

