using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace VaultAutomationTool.Tools
{
    /// <summary>
    /// Outil de diagnostic pour lister tous les Property Sets d'un fichier OLE Structured Storage
    /// </summary>
    public static class DiagnoseOleProperties
    {
        [DllImport("ole32.dll")]
        private static extern int StgOpenStorageEx(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
            uint grfMode,
            uint stgfmt,
            uint grfAttrs,
            IntPtr pStgOptions,
            IntPtr reserved2,
            ref Guid riid,
            out IPropertySetStorage ppObjectOpen);

        private const uint STGM_DIRECT = 0x00000000;
        private const uint STGM_READ = 0x00000000;
        private const uint STGM_SHARE_DENY_WRITE = 0x00000020;
        private const uint STGFMT_ANY = 4;

        [ComImport]
        [Guid("0000013A-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertySetStorage
        {
            int Create(ref Guid rfmtid, IntPtr pclsid, uint grfFlags, uint grfMode, out IntPtr ppprstg);
            int Open(ref Guid rfmtid, uint grfMode, out IntPtr ppprstg);
            int Delete(ref Guid rfmtid);
            int Enum(out IEnumSTATPROPSETSTG ppenum);
        }

        [ComImport]
        [Guid("0000013B-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IEnumSTATPROPSETSTG
        {
            [PreserveSig]
            int Next(uint celt, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] STATPROPSETSTG[] rgelt, out uint pceltFetched);
            int Skip(uint celt);
            int Reset();
            int Clone(out IEnumSTATPROPSETSTG ppenum);
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

        /// <summary>
        /// Liste tous les Property Sets d'un fichier Inventor
        /// </summary>
        public static void ListPropertySets(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Fichier non trouvé: {filePath}");
                return;
            }

            Console.WriteLine($"\n═══════════════════════════════════════════════════════");
            Console.WriteLine($"DIAGNOSTIC OLE PROPERTIES: {Path.GetFileName(filePath)}");
            Console.WriteLine($"═══════════════════════════════════════════════════════\n");

            IPropertySetStorage propSetStorage = null;

            try
            {
                Guid iidPropSetStg = typeof(IPropertySetStorage).GUID;
                uint mode = STGM_DIRECT | STGM_READ | STGM_SHARE_DENY_WRITE;

                int hr = StgOpenStorageEx(
                    filePath,
                    mode,
                    STGFMT_ANY,
                    0,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    ref iidPropSetStg,
                    out propSetStorage);

                if (hr != 0)
                {
                    Console.WriteLine($"Erreur ouverture: 0x{hr:X8}");
                    return;
                }

                IEnumSTATPROPSETSTG enumPropSetStg;
                hr = propSetStorage.Enum(out enumPropSetStg);
                
                if (hr != 0)
                {
                    Console.WriteLine($"Erreur énumération: 0x{hr:X8}");
                    return;
                }

                STATPROPSETSTG[] stats = new STATPROPSETSTG[1];
                uint fetched;
                int index = 1;

                Console.WriteLine("Property Sets trouvés:");
                Console.WriteLine("─────────────────────────────────────────────────────────\n");

                while (enumPropSetStg.Next(1, stats, out fetched) == 0 && fetched > 0)
                {
                    var stat = stats[0];
                    string knownName = GetKnownPropertySetName(stat.fmtid);
                    
                    Console.WriteLine($"[{index}] GUID: {stat.fmtid}");
                    Console.WriteLine($"    Nom:  {knownName}");
                    Console.WriteLine($"    CLSID: {stat.clsid}");
                    Console.WriteLine($"    Flags: 0x{stat.grfFlags:X8}");
                    Console.WriteLine();
                    
                    index++;
                }

                Console.WriteLine("═══════════════════════════════════════════════════════\n");

                Marshal.ReleaseComObject(enumPropSetStg);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
            finally
            {
                if (propSetStorage != null)
                    Marshal.ReleaseComObject(propSetStorage);
            }
        }

        private static string GetKnownPropertySetName(Guid fmtid)
        {
            // GUIDs OLE standard
            if (fmtid == new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"))
                return "SummaryInformation (Standard OLE)";
            if (fmtid == new Guid("D5CDD502-2E9C-101B-9397-08002B2CF9AE"))
                return "DocumentSummaryInformation (Standard OLE)";
            if (fmtid == new Guid("D5CDD505-2E9C-101B-9397-08002B2CF9AE"))
                return "UserDefinedProperties (Standard OLE)";
            
            // GUIDs Inventor connus
            if (fmtid == new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"))
                return "Inventor SummaryInfo";
            
            // Design Tracking Properties (Inventor)
            if (fmtid.ToString().ToUpperInvariant().Contains("32853F0F"))
                return "Inventor Design Tracking Properties";
                
            return $"Unknown ({fmtid})";
        }
    }
}
