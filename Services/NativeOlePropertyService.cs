using System;
using System.IO;
using System.Runtime.InteropServices;

using XnrgyEngineeringAutomationTools.Services;

namespace VaultAutomationTool.Services
{
    /// <summary>
    /// Service haute performance pour modifier les Custom iProperties des fichiers Inventor
    /// via les API Windows OLE Structured Storage natives.
    /// 
    /// Performance: ~50-100ms par fichier (vs 3-15s avec Inventor.Application)
    /// 
    /// IMPORTANT: Ce service écrit dans le PropertySet OLE standard User Defined Properties.
    /// Les fichiers Inventor 2026+ peuvent lire ces propriétés si elles existent dans ce format.
    /// </summary>
    public sealed class NativeOlePropertyService : IDisposable
    {
        #region Constants

        // Storage Mode Flags (STGM)
        private const int STGM_READ = 0x00000000;
        private const int STGM_WRITE = 0x00000001;
        private const int STGM_READWRITE = 0x00000002;
        private const int STGM_SHARE_EXCLUSIVE = 0x00000010;
        private const int STGM_SHARE_DENY_WRITE = 0x00000020;
        private const int STGM_DIRECT = 0x00000000;
        private const int STGM_TRANSACTED = 0x00010000;
        private const int STGM_CREATE = 0x00001000;

        // Storage Format (STGFMT)
        private const int STGFMT_STORAGE = 0;
        private const int STGFMT_FILE = 3;
        private const int STGFMT_ANY = 4;
        private const int STGFMT_DOCFILE = 5;

        // Property Set Flags
        private const int PROPSETFLAG_DEFAULT = 0;
        private const int PROPSETFLAG_NONSIMPLE = 1;
        private const int PROPSETFLAG_ANSI = 2;
        private const int PROPSETFLAG_UNBUFFERED = 4;
        private const int PROPSETFLAG_CASE_SENSITIVE = 8;

        // PROPSPEC Kind
        private const uint PRSPEC_LPWSTR = 0;
        private const uint PRSPEC_PROPID = 1;

        // PROPVARIANT Types (VT)
        private const ushort VT_EMPTY = 0;
        private const ushort VT_NULL = 1;
        private const ushort VT_I2 = 2;
        private const ushort VT_I4 = 3;
        private const ushort VT_BSTR = 8;
        private const ushort VT_BOOL = 11;
        private const ushort VT_VARIANT = 12;
        private const ushort VT_UI4 = 19;
        private const ushort VT_LPSTR = 30;
        private const ushort VT_LPWSTR = 31;
        private const ushort VT_FILETIME = 64;

        // Property IDs for our custom properties
        private const uint PROPID_PROJECT = 2;
        private const uint PROPID_REFERENCE = 3;
        private const uint PROPID_MODULE = 4;

        // HRESULT codes
        private const int S_OK = 0;
        private const int STG_E_FILENOTFOUND = unchecked((int)0x80030002);
        private const int STG_E_ACCESSDENIED = unchecked((int)0x80030005);
        private const int STG_E_FILEALREADYEXISTS = unchecked((int)0x80030050);

        #endregion

        #region GUIDs

        /// <summary>
        /// FMTID_UserDefinedProperties - Standard OLE GUID for custom/user-defined properties
        /// {D5CDD505-2E9C-101B-9397-08002B2CF9AE}
        /// </summary>
        private static readonly Guid FMTID_UserDefinedProperties = new Guid(
            0xD5CDD505, 0x2E9C, 0x101B, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE);

        /// <summary>
        /// FMTID_SummaryInformation - Standard document summary properties
        /// {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
        /// </summary>
        private static readonly Guid FMTID_SummaryInformation = new Guid(
            0xF29F85E0, 0x4FF9, 0x1068, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9);

        /// <summary>
        /// FMTID_DocSummaryInformation - Document summary information
        /// {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
        /// </summary>
        private static readonly Guid FMTID_DocSummaryInformation = new Guid(
            0xD5CDD502, 0x2E9C, 0x101B, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE);

        /// <summary>
        /// IID_IPropertySetStorage
        /// {0000013A-0000-0000-C000-000000000046}
        /// </summary>
        private static readonly Guid IID_IPropertySetStorage = new Guid(
            0x0000013A, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);

        #endregion

        #region P/Invoke Declarations

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int StgOpenStorageEx(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
            int grfMode,
            int stgfmt,
            int grfAttrs,
            IntPtr pStgOptions,
            IntPtr reserved2,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IPropertySetStorage ppObjectOpen);

        [DllImport("ole32.dll", PreserveSig = true)]
        private static extern int PropVariantClear(ref PROPVARIANT pvar);

        [DllImport("ole32.dll", PreserveSig = true)]
        private static extern int CoInitialize(IntPtr pvReserved);

        [DllImport("ole32.dll", PreserveSig = true)]
        private static extern void CoUninitialize();

        #endregion

        #region COM Interfaces

        [ComImport]
        [Guid("0000013A-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertySetStorage
        {
            [PreserveSig]
            int Create(
                ref Guid rfmtid,
                IntPtr pclsid,
                int grfFlags,
                int grfMode,
                [MarshalAs(UnmanagedType.Interface)] out IPropertyStorage ppprstg);

            [PreserveSig]
            int Open(
                ref Guid rfmtid,
                int grfMode,
                [MarshalAs(UnmanagedType.Interface)] out IPropertyStorage ppprstg);

            [PreserveSig]
            int Delete(ref Guid rfmtid);

            [PreserveSig]
            int Enum(out IntPtr ppenum);
        }

        [ComImport]
        [Guid("00000138-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyStorage
        {
            [PreserveSig]
            int ReadMultiple(
                uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
                [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar);

            [PreserveSig]
            int WriteMultiple(
                uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar,
                uint propidNameFirst);

            [PreserveSig]
            int DeleteMultiple(
                uint cpspec,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec);

            [PreserveSig]
            int ReadPropertyNames(
                uint cpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid,
                [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IntPtr[] rglpwstrName);

            [PreserveSig]
            int WritePropertyNames(
                uint cpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.LPWStr)] string[] rglpwstrName);

            [PreserveSig]
            int DeletePropertyNames(
                uint cpropid,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid);

            [PreserveSig]
            int Commit(int grfCommitFlags);

            [PreserveSig]
            int Revert();

            [PreserveSig]
            int Enum(out IntPtr ppenum);

            [PreserveSig]
            int SetTimes(
                ref System.Runtime.InteropServices.ComTypes.FILETIME pctime,
                ref System.Runtime.InteropServices.ComTypes.FILETIME patime,
                ref System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

            [PreserveSig]
            int SetClass(ref Guid clsid);

            [PreserveSig]
            int Stat(out STATPROPSETSTG pstatpsstg);
        }

        #endregion

        #region Structures

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPSPEC
        {
            public uint ulKind;
            public IntPtr unionmember;

            public static PROPSPEC FromPropertyId(uint propId)
            {
                return new PROPSPEC
                {
                    ulKind = PRSPEC_PROPID,
                    unionmember = (IntPtr)propId
                };
            }

            public static PROPSPEC FromPropertyName(string name)
            {
                return new PROPSPEC
                {
                    ulKind = PRSPEC_LPWSTR,
                    unionmember = Marshal.StringToCoTaskMemUni(name)
                };
            }
        }

        [StructLayout(LayoutKind.Explicit, Size = 16)]
        private struct PROPVARIANT
        {
            [FieldOffset(0)] public ushort vt;
            [FieldOffset(2)] public ushort wReserved1;
            [FieldOffset(4)] public ushort wReserved2;
            [FieldOffset(6)] public ushort wReserved3;
            [FieldOffset(8)] public IntPtr pointerValue;
            [FieldOffset(8)] public int intValue;
            [FieldOffset(8)] public long longValue;

            public static PROPVARIANT FromString(string value)
            {
                var pv = new PROPVARIANT
                {
                    vt = VT_LPWSTR,
                    pointerValue = Marshal.StringToCoTaskMemUni(value)
                };
                return pv;
            }

            public string GetStringValue()
            {
                if (vt == VT_LPWSTR && pointerValue != IntPtr.Zero)
                    return Marshal.PtrToStringUni(pointerValue);
                if (vt == VT_LPSTR && pointerValue != IntPtr.Zero)
                    return Marshal.PtrToStringAnsi(pointerValue);
                if (vt == VT_BSTR && pointerValue != IntPtr.Zero)
                    return Marshal.PtrToStringBSTR(pointerValue);
                return null;
            }

            public void Clear()
            {
                if ((vt == VT_LPWSTR || vt == VT_LPSTR) && pointerValue != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pointerValue);
                    pointerValue = IntPtr.Zero;
                }
                vt = VT_EMPTY;
            }
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

        #endregion

        #region Fields

        private bool _disposed;
        private readonly object _lock = new object();

        #endregion

        #region Public Methods

        /// <summary>
        /// Définit les iProperties Custom (Project, Reference, Module) dans un fichier Inventor.
        /// Crée automatiquement un backup et le restaure en cas d'erreur.
        /// </summary>
        /// <param name="filePath">Chemin complet du fichier .ipt, .iam, .idw ou .ipn</param>
        /// <param name="project">Valeur de la propriété Project (ex: "12345")</param>
        /// <param name="reference">Valeur de la propriété Reference (ex: "REF01")</param>
        /// <param name="module">Valeur de la propriété Module (ex: "M01")</param>
        /// <returns>True si succès, False sinon</returns>
        public bool SetIProperties(string filePath, string project, string reference, string module)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                Logger.Log("[OLE] Chemin fichier vide", Logger.LogLevel.ERROR);
                return false;
            }

            if (!File.Exists(filePath))
            {
                Logger.Log($"[OLE] Fichier non trouvé: {filePath}", Logger.LogLevel.ERROR);
                return false;
            }

            // Vérifier l'extension
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".ipt" && ext != ".iam" && ext != ".idw" && ext != ".ipn")
            {
                Logger.Log($"[OLE] Extension non supportée: {ext}", Logger.LogLevel.WARNING);
                return false;
            }

            // Vérifier qu'au moins une propriété est fournie
            if (string.IsNullOrEmpty(project) && string.IsNullOrEmpty(reference) && string.IsNullOrEmpty(module))
            {
                Logger.Log("[OLE] Aucune propriété à définir", Logger.LogLevel.WARNING);
                return true; // Pas d'erreur, juste rien à faire
            }

            string backupPath = null;
            bool wasReadOnly = false;
            var startTime = DateTime.Now;

            try
            {
                // Gérer l'attribut ReadOnly
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.IsReadOnly)
                {
                    wasReadOnly = true;
                    fileInfo.IsReadOnly = false;
                    Logger.Log($"[OLE] ReadOnly retiré temporairement: {Path.GetFileName(filePath)}", Logger.LogLevel.DEBUG);
                }

                // Créer un backup
                backupPath = filePath + ".bak_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                File.Copy(filePath, backupPath, true);
                Logger.Log($"[OLE] Backup créé: {Path.GetFileName(backupPath)}", Logger.LogLevel.DEBUG);

                // Écrire les propriétés
                bool success = WritePropertiesToFile(filePath, project, reference, module);

                if (success)
                {
                    // Supprimer le backup si succès
                    try { File.Delete(backupPath); } catch { }
                    backupPath = null;

                    var elapsed = (DateTime.Now - startTime).TotalMilliseconds;
                    Logger.Log($"[OLE] [+] iProperties définies en {elapsed:F0}ms: {Path.GetFileName(filePath)}", Logger.LogLevel.INFO);
                    return true;
                }
                else
                {
                    Logger.Log($"[OLE] [-] Échec écriture propriétés: {Path.GetFileName(filePath)}", Logger.LogLevel.ERROR);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[OLE] Exception: {ex.Message}", Logger.LogLevel.ERROR);
                Logger.Log($"[OLE] Stack: {ex.StackTrace}", Logger.LogLevel.DEBUG);
                return false;
            }
            finally
            {
                // Restaurer le backup si nécessaire
                if (!string.IsNullOrEmpty(backupPath) && File.Exists(backupPath))
                {
                    try
                    {
                        File.Copy(backupPath, filePath, true);
                        File.Delete(backupPath);
                        Logger.Log($"[OLE] Backup restauré suite à erreur", Logger.LogLevel.WARNING);
                    }
                    catch (Exception restoreEx)
                    {
                        Logger.Log($"[OLE] Erreur restauration backup: {restoreEx.Message}", Logger.LogLevel.ERROR);
                    }
                }

                // Restaurer ReadOnly si nécessaire
                if (wasReadOnly)
                {
                    try
                    {
                        new FileInfo(filePath).IsReadOnly = true;
                        Logger.Log($"[OLE] ReadOnly restauré", Logger.LogLevel.DEBUG);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Lit les iProperties Custom d'un fichier Inventor.
        /// </summary>
        public (string Project, string Reference, string Module) ReadIProperties(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return (null, null, null);

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

                if (hr != S_OK || propSetStorage == null)
                    return (null, null, null);

                Guid fmtid = FMTID_UserDefinedProperties;
                hr = propSetStorage.Open(ref fmtid, STGM_READ | STGM_SHARE_EXCLUSIVE, out propStorage);

                if (hr != S_OK || propStorage == null)
                    return (null, null, null);

                string project = ReadPropertyValue(propStorage, PROPID_PROJECT);
                string reference = ReadPropertyValue(propStorage, PROPID_REFERENCE);
                string module = ReadPropertyValue(propStorage, PROPID_MODULE);

                return (project, reference, module);
            }
            catch
            {
                return (null, null, null);
            }
            finally
            {
                SafeRelease(propStorage);
                SafeRelease(propSetStorage);
            }
        }

        /// <summary>
        /// Vérifie si un fichier est un OLE Compound Document valide.
        /// </summary>
        public bool IsValidOleFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            IPropertySetStorage propSetStorage = null;

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

                return hr == S_OK && propSetStorage != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                SafeRelease(propSetStorage);
            }
        }

        #endregion

        #region Private Methods

        private bool WritePropertiesToFile(string filePath, string project, string reference, string module)
        {
            IPropertySetStorage propSetStorage = null;
            IPropertyStorage propStorage = null;

            try
            {
                // Ouvrir le fichier en mode lecture/écriture
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

                if (hr != S_OK || propSetStorage == null)
                {
                    Logger.Log($"[OLE] StgOpenStorageEx échoué: 0x{hr:X8}", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log("[OLE] Storage ouvert en écriture", Logger.LogLevel.DEBUG);

                // Essayer d'ouvrir le property set User Defined existant
                Guid fmtid = FMTID_UserDefinedProperties;
                hr = propSetStorage.Open(ref fmtid, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, out propStorage);

                if (hr == STG_E_FILENOTFOUND || propStorage == null)
                {
                    // Le property set n'existe pas, le créer
                    Logger.Log("[OLE] Property set non trouvé, création...", Logger.LogLevel.DEBUG);
                    
                    hr = propSetStorage.Create(
                        ref fmtid,
                        IntPtr.Zero,
                        PROPSETFLAG_DEFAULT,
                        STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE,
                        out propStorage);

                    if (hr != S_OK || propStorage == null)
                    {
                        Logger.Log($"[OLE] Création property set échouée: 0x{hr:X8}", Logger.LogLevel.ERROR);
                        return false;
                    }

                    Logger.Log("[OLE] Property set créé", Logger.LogLevel.DEBUG);
                }
                else if (hr != S_OK)
                {
                    Logger.Log($"[OLE] Ouverture property set échouée: 0x{hr:X8}", Logger.LogLevel.ERROR);
                    return false;
                }

                // Écrire les propriétés
                int propsWritten = 0;

                if (!string.IsNullOrEmpty(project))
                {
                    if (WriteProperty(propStorage, "Project", project, PROPID_PROJECT))
                        propsWritten++;
                }

                if (!string.IsNullOrEmpty(reference))
                {
                    if (WriteProperty(propStorage, "Reference", reference, PROPID_REFERENCE))
                        propsWritten++;
                }

                if (!string.IsNullOrEmpty(module))
                {
                    if (WriteProperty(propStorage, "Module", module, PROPID_MODULE))
                        propsWritten++;
                }

                // Commit les changements
                hr = propStorage.Commit(0);
                if (hr != S_OK)
                {
                    Logger.Log($"[OLE] Commit échoué: 0x{hr:X8}", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"[OLE] {propsWritten} propriétés écrites et commitées", Logger.LogLevel.DEBUG);
                return propsWritten > 0;
            }
            finally
            {
                SafeRelease(propStorage);
                SafeRelease(propSetStorage);
            }
        }

        private bool WriteProperty(IPropertyStorage propStorage, string name, string value, uint propId)
        {
            try
            {
                // 1. Enregistrer le nom de la propriété avec son ID
                int hr = propStorage.WritePropertyNames(
                    1,
                    new uint[] { propId },
                    new string[] { name });

                // Ignorer l'erreur si le nom existe déjà
                if (hr != S_OK && hr != STG_E_FILEALREADYEXISTS)
                {
                    Logger.Log($"[OLE] WritePropertyNames '{name}' warning: 0x{hr:X8}", Logger.LogLevel.DEBUG);
                }

                // 2. Créer la PROPSPEC avec l'ID
                var propSpec = PROPSPEC.FromPropertyId(propId);

                // 3. Créer la PROPVARIANT avec la valeur string
                var propVar = PROPVARIANT.FromString(value);

                try
                {
                    // 4. Écrire la valeur
                    hr = propStorage.WriteMultiple(1, new[] { propSpec }, new[] { propVar }, 2);

                    if (hr != S_OK)
                    {
                        Logger.Log($"[OLE] WriteMultiple '{name}' échoué: 0x{hr:X8}", Logger.LogLevel.ERROR);
                        return false;
                    }

                    Logger.Log($"[OLE] Propriété '{name}' = '{value}' écrite", Logger.LogLevel.DEBUG);
                    return true;
                }
                finally
                {
                    // Libérer la mémoire allouée pour la string
                    propVar.Clear();
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[OLE] Exception WriteProperty '{name}': {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        private string ReadPropertyValue(IPropertyStorage propStorage, uint propId)
        {
            try
            {
                var propSpec = PROPSPEC.FromPropertyId(propId);
                var propVar = new PROPVARIANT[1];

                int hr = propStorage.ReadMultiple(1, new[] { propSpec }, propVar);

                if (hr == S_OK)
                {
                    string value = propVar[0].GetStringValue();
                    propVar[0].Clear();
                    return value;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static void SafeRelease(object comObject)
        {
            if (comObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch { }
            }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }

        #endregion
    }
}

