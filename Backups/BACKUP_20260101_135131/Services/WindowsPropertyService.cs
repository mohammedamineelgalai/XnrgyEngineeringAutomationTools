using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using XnrgyEngineeringAutomationTools.Services;

namespace VaultAutomationTool.Services
{
    /// <summary>
    /// Service pour modifier les propriétés OLE des fichiers via les API Windows natives.
    /// Fonctionne au niveau "Windows Explorer" sans besoin d'Inventor API.
    /// Utilise IPropertySetStorage/IPropertyStorage pour les fichiers OLE Structured Storage.
    /// </summary>
    public class WindowsPropertyService : IDisposable
    {
        private bool _disposed;

        #region Win32 Constants
        
        // STGM - Storage mode flags
        private const int STGM_READ = 0x00000000;
        private const int STGM_WRITE = 0x00000001;
        private const int STGM_READWRITE = 0x00000002;
        private const int STGM_SHARE_EXCLUSIVE = 0x00000010;
        private const int STGM_SHARE_DENY_WRITE = 0x00000020;
        private const int STGM_SHARE_DENY_NONE = 0x00000040;
        private const int STGM_DIRECT = 0x00000000;
        private const int STGM_TRANSACTED = 0x00010000;

        // STGFMT - Storage format
        private const int STGFMT_STORAGE = 0;
        private const int STGFMT_FILE = 3;
        private const int STGFMT_ANY = 4;
        private const int STGFMT_DOCFILE = 5;

        // Property IDs
        private const int PIDSI_TITLE = 0x00000002;
        private const int PIDSI_SUBJECT = 0x00000003;
        private const int PIDSI_AUTHOR = 0x00000004;
        private const int PIDSI_KEYWORDS = 0x00000005;
        private const int PIDSI_COMMENTS = 0x00000006;

        #endregion

        #region GUIDs

        // FMTID_UserDefinedProperties - {D5CDD505-2E9C-101B-9397-08002B2CF9AE}
        private static readonly Guid FMTID_UserDefinedProperties = new Guid(
            0xD5CDD505, 0x2E9C, 0x101B, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE);

        // FMTID_SummaryInformation - {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
        private static readonly Guid FMTID_SummaryInformation = new Guid(
            0xF29F85E0, 0x4FF9, 0x1068, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9);

        // FMTID_DocSummaryInformation - {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
        private static readonly Guid FMTID_DocSummaryInformation = new Guid(
            0xD5CDD502, 0x2E9C, 0x101B, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE);

        // IID_IPropertySetStorage
        private static readonly Guid IID_IPropertySetStorage = new Guid(
            0x0000013A, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);

        #endregion

        #region P/Invoke Declarations

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int StgOpenStorageEx(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
            int grfMode,
            int stgfmt,
            int grfAttrs,
            IntPtr pStgOptions,
            IntPtr reserved2,
            ref Guid riid,
            out IPropertySetStorage ppObjectOpen);

        [DllImport("ole32.dll")]
        private static extern int PropVariantClear(ref PROPVARIANT pvar);

        #endregion

        #region COM Interfaces

        [ComImport]
        [Guid("0000013A-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertySetStorage
        {
            int Create(
                ref Guid rfmtid,
                IntPtr pclsid,
                int grfFlags,
                int grfMode,
                out IPropertyStorage ppprstg);

            int Open(
                ref Guid rfmtid,
                int grfMode,
                out IPropertyStorage ppprstg);

            int Delete(ref Guid rfmtid);

            int Enum(out IntPtr ppenum);
        }

        [ComImport]
        [Guid("00000138-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyStorage
        {
            int ReadMultiple(
                uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
                [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar);

            int WriteMultiple(
                uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar,
                uint propidNameFirst);

            int DeleteMultiple(
                uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec);

            int ReadPropertyNames(
                uint cpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid,
                [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IntPtr[] rglpwstrName);

            int WritePropertyNames(
                uint cpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.LPWStr)] string[] rglpwstrName);

            int DeletePropertyNames(
                uint cpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid);

            int Commit(int grfCommitFlags);

            int Revert();

            int Enum(out IntPtr ppenum);

            int SetTimes(
                ref System.Runtime.InteropServices.ComTypes.FILETIME pctime,
                ref System.Runtime.InteropServices.ComTypes.FILETIME patime,
                ref System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

            int SetClass(ref Guid clsid);

            int Stat(out STATPROPSETSTG pstatpsstg);
        }

        #endregion

        #region Structures

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPSPEC
        {
            public uint ulKind;
            public PropSpecUnion union;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct PropSpecUnion
        {
            [FieldOffset(0)]
            public uint propid;
            [FieldOffset(0)]
            public IntPtr lpwstr;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPVARIANT
        {
            public ushort vt;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            public IntPtr p;
            public int p2;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct STATPROPSETSTG
        {
            public Guid fmtid;
            public Guid clsid;
            public uint grfFlags;
            public System.Runtime.InteropServices.ComTypes.FILETIME mtime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ctime;
            public System.Runtime.InteropServices.ComTypes.FILETIME atime;
            public uint dwOSVersion;
        }

        // PROPSPEC kinds
        private const uint PRSPEC_LPWSTR = 0;
        private const uint PRSPEC_PROPID = 1;

        // PROPVARIANT types
        private const ushort VT_EMPTY = 0;
        private const ushort VT_NULL = 1;
        private const ushort VT_LPSTR = 30;
        private const ushort VT_LPWSTR = 31;
        private const ushort VT_BSTR = 8;

        // PROPSETFLAG
        private const int PROPSETFLAG_DEFAULT = 0;
        private const int PROPSETFLAG_ANSI = 2;

        #endregion

        /// <summary>
        /// Définit les propriétés personnalisées (User Defined Properties) d'un fichier Inventor
        /// en utilisant les API Windows OLE Structured Storage.
        /// </summary>
        public bool SetCustomProperties(string filePath, string project, string reference, string module)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Logger.Log($"   [WIN-PROP] Fichier non trouvé: {filePath}", Logger.LogLevel.ERROR);
                return false;
            }

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".ipt" && ext != ".iam" && ext != ".idw" && ext != ".ipn")
            {
                Logger.Log($"   [WIN-PROP] Extension non supportée: {ext}", Logger.LogLevel.WARNING);
                return false;
            }

            Logger.Log($"   [WIN-PROP] Modification propriétés OLE: {Path.GetFileName(filePath)}", Logger.LogLevel.INFO);

            bool wasReadOnly = false;
            try
            {
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.IsReadOnly)
                {
                    wasReadOnly = true;
                    fileInfo.IsReadOnly = false;
                    Logger.Log("   [WIN-PROP] Attribut ReadOnly retiré temporairement", Logger.LogLevel.DEBUG);
                }
            }
            catch (Exception roEx)
            {
                Logger.Log($"   [WIN-PROP] Erreur gestion ReadOnly: {roEx.Message}", Logger.LogLevel.WARNING);
                return false;
            }

            IPropertySetStorage propSetStorage = null;
            IPropertyStorage propStorage = null;

            try
            {
                // Ouvrir le fichier en mode lecture/écriture via OLE Structured Storage
                Guid iid = IID_IPropertySetStorage;
                int hr = StgOpenStorageEx(
                    filePath,
                    STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DIRECT,
                    STGFMT_ANY,
                    0,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    ref iid,
                    out propSetStorage);

                if (hr != 0 || propSetStorage == null)
                {
                    Logger.Log($"   [WIN-PROP] StgOpenStorageEx échoué: 0x{hr:X8}", Logger.LogLevel.ERROR);
                    
                    // Essayer avec STGFMT_STORAGE
                    hr = StgOpenStorageEx(
                        filePath,
                        STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DIRECT,
                        STGFMT_STORAGE,
                        0,
                        IntPtr.Zero,
                        IntPtr.Zero,
                        ref iid,
                        out propSetStorage);

                    if (hr != 0 || propSetStorage == null)
                    {
                        Logger.Log($"   [WIN-PROP] Deuxième tentative échouée: 0x{hr:X8}", Logger.LogLevel.ERROR);
                        return false;
                    }
                }

                Logger.Log("   [WIN-PROP] Fichier OLE ouvert en écriture", Logger.LogLevel.DEBUG);

                // Ouvrir ou créer le property set User Defined
                Guid fmtid = FMTID_UserDefinedProperties;
                hr = propSetStorage.Open(ref fmtid, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, out propStorage);

                if (hr != 0 || propStorage == null)
                {
                    Logger.Log("   [WIN-PROP] Property set non existant, création...", Logger.LogLevel.DEBUG);
                    
                    // Créer le property set s'il n'existe pas
                    hr = propSetStorage.Create(
                        ref fmtid,
                        IntPtr.Zero,
                        PROPSETFLAG_DEFAULT,
                        STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE,
                        out propStorage);

                    if (hr != 0 || propStorage == null)
                    {
                        Logger.Log($"   [WIN-PROP] Impossible de créer property set: 0x{hr:X8}", Logger.LogLevel.ERROR);
                        return false;
                    }
                }

                Logger.Log("   [WIN-PROP] Property set User Defined ouvert", Logger.LogLevel.DEBUG);

                // Écrire les propriétés
                int propsWritten = 0;
                
                if (!string.IsNullOrEmpty(project))
                {
                    if (WriteStringProperty(propStorage, "Project", project))
                        propsWritten++;
                }
                
                if (!string.IsNullOrEmpty(reference))
                {
                    if (WriteStringProperty(propStorage, "Reference", reference))
                        propsWritten++;
                }
                
                if (!string.IsNullOrEmpty(module))
                {
                    if (WriteStringProperty(propStorage, "Module", module))
                        propsWritten++;
                }

                // Commit des changements
                hr = propStorage.Commit(0);
                if (hr != 0)
                {
                    Logger.Log($"   [WIN-PROP] Commit échoué: 0x{hr:X8}", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"   [WIN-PROP] {propsWritten} propriétés écrites et commitées", Logger.LogLevel.INFO);
                return propsWritten > 0;
            }
            catch (Exception ex)
            {
                Logger.Log($"   [WIN-PROP] Exception: {ex.Message}", Logger.LogLevel.ERROR);
                Logger.Log($"   [WIN-PROP] Stack: {ex.StackTrace}", Logger.LogLevel.DEBUG);
                return false;
            }
            finally
            {
                // Libérer les ressources COM
                if (propStorage != null)
                {
                    try { Marshal.ReleaseComObject(propStorage); } catch { }
                }
                if (propSetStorage != null)
                {
                    try { Marshal.ReleaseComObject(propSetStorage); } catch { }
                }

                // Restaurer l'attribut ReadOnly si nécessaire
                if (wasReadOnly)
                {
                    try 
                    { 
                        new FileInfo(filePath).IsReadOnly = true;
                        Logger.Log("   [WIN-PROP] Attribut ReadOnly restauré", Logger.LogLevel.DEBUG);
                    } 
                    catch { }
                }
            }
        }

        /// <summary>
        /// Écrit une propriété string nommée dans le property storage.
        /// </summary>
        private bool WriteStringProperty(IPropertyStorage propStorage, string propertyName, string value)
        {
            IntPtr namePtr = IntPtr.Zero;
            IntPtr valuePtr = IntPtr.Zero;

            try
            {
                Logger.Log($"   [WIN-PROP] Écriture propriété '{propertyName}' = '{value}'", Logger.LogLevel.DEBUG);

                // Allouer un Property ID pour cette propriété nommée
                // On utilise un ID basé sur le hash du nom pour la persistance
                uint propId = (uint)(Math.Abs(propertyName.GetHashCode()) % 0x7FFFFFFF);
                if (propId < 2) propId = 2; // Les IDs 0 et 1 sont réservés

                // D'abord, enregistrer le nom de la propriété avec son ID
                var propIds = new uint[] { propId };
                var propNames = new string[] { propertyName };
                
                int hr = propStorage.WritePropertyNames(1, propIds, propNames);
                if (hr != 0)
                {
                    Logger.Log($"   [WIN-PROP] WritePropertyNames warning: 0x{hr:X8}", Logger.LogLevel.DEBUG);
                    // Ce n'est pas fatal, on continue
                }

                // Créer le PROPSPEC avec l'ID numérique
                var propSpec = new PROPSPEC
                {
                    ulKind = PRSPEC_PROPID,
                    union = new PropSpecUnion { propid = propId }
                };

                // Créer la PROPVARIANT avec la valeur string
                valuePtr = Marshal.StringToCoTaskMemUni(value);
                var propVar = new PROPVARIANT
                {
                    vt = VT_LPWSTR,
                    p = valuePtr
                };

                // Écrire la propriété
                hr = propStorage.WriteMultiple(1, new[] { propSpec }, new[] { propVar }, 2);
                
                if (hr != 0)
                {
                    Logger.Log($"   [WIN-PROP] WriteMultiple échoué: 0x{hr:X8}", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"   [WIN-PROP] Propriété '{propertyName}' écrite avec succès", Logger.LogLevel.DEBUG);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"   [WIN-PROP] Erreur écriture '{propertyName}': {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
            finally
            {
                if (namePtr != IntPtr.Zero)
                    Marshal.FreeCoTaskMem(namePtr);
                // Note: valuePtr sera libéré par PropVariantClear si nécessaire
            }
        }

        /// <summary>
        /// Lit une propriété string depuis un fichier.
        /// </summary>
        public string ReadCustomProperty(string filePath, string propertyName)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return null;

            IPropertySetStorage propSetStorage = null;
            IPropertyStorage propStorage = null;

            try
            {
                Guid iid = IID_IPropertySetStorage;
                int hr = StgOpenStorageEx(
                    filePath,
                    STGM_READ | STGM_SHARE_DENY_WRITE,
                    STGFMT_ANY,
                    0,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    ref iid,
                    out propSetStorage);

                if (hr != 0 || propSetStorage == null)
                    return null;

                Guid fmtid = FMTID_UserDefinedProperties;
                hr = propSetStorage.Open(ref fmtid, STGM_READ | STGM_SHARE_EXCLUSIVE, out propStorage);

                if (hr != 0 || propStorage == null)
                    return null;

                // Calculer l'ID de la propriété
                uint propId = (uint)(Math.Abs(propertyName.GetHashCode()) % 0x7FFFFFFF);
                if (propId < 2) propId = 2;

                var propSpec = new PROPSPEC
                {
                    ulKind = PRSPEC_PROPID,
                    union = new PropSpecUnion { propid = propId }
                };

                var propVar = new PROPVARIANT[1];
                hr = propStorage.ReadMultiple(1, new[] { propSpec }, propVar);

                if (hr == 0 && propVar[0].vt == VT_LPWSTR && propVar[0].p != IntPtr.Zero)
                {
                    return Marshal.PtrToStringUni(propVar[0].p);
                }

                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (propStorage != null)
                    try { Marshal.ReleaseComObject(propStorage); } catch { }
                if (propSetStorage != null)
                    try { Marshal.ReleaseComObject(propSetStorage); } catch { }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }

        // Constante manquante pour Create
        private const int STGM_CREATE = 0x00001000;
    }
}

