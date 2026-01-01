using System;
using System.IO;

using XnrgyEngineeringAutomationTools.Services;

namespace VaultAutomationTool.Services
{
    public class OlePropertyService : IDisposable
    {
        private static dynamic _apprentice;
        private static bool _apprenticeAvailable = true;
        private static readonly object _lock = new object();
        private bool _disposed;

        private static bool EnsureApprentice()
        {
            if (_apprentice != null) return true;
            if (!_apprenticeAvailable) return false;

            lock (_lock)
            {
                if (_apprentice != null) return true;

                try
                {
                    Logger.Log("   [APPRENTICE] Init ApprenticeServer...", Logger.LogLevel.INFO);
                    Type apprenticeType = Type.GetTypeFromProgID("Inventor.ApprenticeServer");
                    if (apprenticeType == null)
                    {
                        _apprenticeAvailable = false;
                        Logger.Log("   [APPRENTICE] ProgID non trouve", Logger.LogLevel.ERROR);
                        return false;
                    }
                    _apprentice = Activator.CreateInstance(apprenticeType);
                    Logger.Log("   [APPRENTICE] ApprenticeServer pret!", Logger.LogLevel.INFO);
                    return true;
                }
                catch (Exception ex)
                {
                    _apprenticeAvailable = false;
                    Logger.Log("   [APPRENTICE] Erreur init: " + ex.Message, Logger.LogLevel.ERROR);
                    return false;
                }
            }
        }

        public bool SetIProperties(string filePath, string project, string reference, string module)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Logger.Log("   [APPRENTICE] Fichier non trouve: " + filePath, Logger.LogLevel.ERROR);
                return false;
            }

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".ipt" && ext != ".iam" && ext != ".idw" && ext != ".ipn")
            {
                Logger.Log("   [APPRENTICE] Extension non supportee: " + ext, Logger.LogLevel.WARNING);
                return false;
            }

            Logger.Log("   [APPRENTICE] Modification iProperties...", Logger.LogLevel.INFO);

            bool wasReadOnly = false;
            try
            {
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.IsReadOnly)
                {
                    wasReadOnly = true;
                    fileInfo.IsReadOnly = false;
                }
            }
            catch (Exception roEx)
            {
                Logger.Log("   [APPRENTICE] Erreur ReadOnly: " + roEx.Message, Logger.LogLevel.WARNING);
                return false;
            }

            dynamic apprenticeDoc = null;
            try
            {
                if (!EnsureApprentice()) return false;

                apprenticeDoc = _apprentice.Open(filePath);
                if (apprenticeDoc == null)
                {
                    Logger.Log("   [APPRENTICE] Erreur ouverture", Logger.LogLevel.ERROR);
                    return false;
                }

                dynamic customProps = null;
                foreach (dynamic propSet in apprenticeDoc.PropertySets)
                {
                    string setName = propSet.Name;
                    if (setName.Contains("User Defined") || setName.Contains("Custom"))
                    {
                        customProps = propSet;
                        break;
                    }
                }

                if (customProps == null)
                {
                    Logger.Log("   [APPRENTICE] Custom props non trouve", Logger.LogLevel.ERROR);
                    return false;
                }

                if (!string.IsNullOrEmpty(project))
                    SetOrCreateProperty(customProps, "Project", project);
                if (!string.IsNullOrEmpty(reference))
                    SetOrCreateProperty(customProps, "Reference", reference);
                if (!string.IsNullOrEmpty(module))
                    SetOrCreateProperty(customProps, "Module", module);

                // NOTE: ApprenticeServer est READ-ONLY - ne peut pas sauvegarder
                Logger.Log("   [APPRENTICE] iProperties modifiees (ApprenticeServer=ReadOnly)", Logger.LogLevel.WARNING);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log("   [APPRENTICE] Exception: " + ex.Message, Logger.LogLevel.ERROR);
                return false;
            }
            finally
            {
                if (apprenticeDoc != null)
                {
                    try { apprenticeDoc.Close(); } catch { }
                }
                if (wasReadOnly)
                {
                    try { new FileInfo(filePath).IsReadOnly = true; } catch { }
                }
            }
        }

        private void SetOrCreateProperty(dynamic customProps, string name, string value)
        {
            try
            {
                bool found = false;
                foreach (dynamic prop in customProps)
                {
                    if (prop.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        prop.Value = value;
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    customProps.Add(value, name);
                }
            }
            catch { }
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

        public static void ReleaseApprentice()
        {
            lock (_lock)
            {
                if (_apprentice != null)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(_apprentice); } catch { }
                    _apprentice = null;
                }
            }
        }
    }
}

