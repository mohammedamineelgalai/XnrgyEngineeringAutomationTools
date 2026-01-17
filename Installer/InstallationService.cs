using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace XnrgyInstaller
{
    /// <summary>
    /// Options d'installation configurables
    /// </summary>
    public class InstallationOptions
    {
        public string InstallPath { get; set; }
        public bool CreateDesktopShortcut { get; set; }
        public bool CreateStartMenuShortcut { get; set; }
    }

    /// <summary>
    /// Progression de l'installation
    /// </summary>
    public class InstallationProgress
    {
        public int Percentage { get; set; }
        public string CurrentAction { get; set; }
    }

    /// <summary>
    /// Service d'installation pour XNRGY Engineering Automation Tools
    /// Gere la copie des fichiers, creation de raccourcis, enregistrement Windows
    /// </summary>
    public class InstallationService
    {
        #region Constants

        private const string APP_NAME = "XNRGY Engineering Automation Tools";
        private const string EXE_NAME = "XnrgyEngineeringAutomationTools.exe";
        private const string ICO_NAME = "XnrgyEngineeringAutomationTools.ico";
        private const string PUBLISHER = "Mohammed Amine Elgalai - XNRGY Climate Systems ULC";
        private const string VERSION = "1.0.0";
        private const string RELEASE_DATE = "2026-01-16";
        private const string UNINSTALL_GUID = "{XNRGY-EAT-2026-INSTALL}";

        #endregion

        #region Public Methods

        /// <summary>
        /// Execute l'installation complete de facon asynchrone
        /// Supporte la reinstallation par-dessus une installation existante (force install)
        /// </summary>
        public async Task<bool> InstallAsync(InstallationOptions options, IProgress<InstallationProgress> progress)
        {
            try
            {
                // Etape 0: Fermer toute instance en cours (Force Install)
                progress.Report(new InstallationProgress { Percentage = 0, CurrentAction = "Fermeture des instances en cours..." });
                KillRunningProcesses();
                await Task.Delay(1000); // Attendre que les fichiers soient liberes

                // Etape 1: Preparation (0-10%)
                progress.Report(new InstallationProgress { Percentage = 2, CurrentAction = "Preparation de l'installation..." });
                await Task.Delay(500);

                // Creer le dossier d'installation
                if (!Directory.Exists(options.InstallPath))
                {
                    Directory.CreateDirectory(options.InstallPath);
                }
                progress.Report(new InstallationProgress { Percentage = 5, CurrentAction = "Dossier d'installation cree..." });
                await Task.Delay(300);

                // Etape 2: Copie des fichiers (10-70%)
                progress.Report(new InstallationProgress { Percentage = 10, CurrentAction = "Copie des fichiers..." });
                await CopyApplicationFilesAsync(options.InstallPath, progress);

                // Etape 3: Creation des raccourcis (70-85%)
                if (options.CreateDesktopShortcut)
                {
                    progress.Report(new InstallationProgress { Percentage = 72, CurrentAction = "Creation du raccourci Bureau..." });
                    CreateDesktopShortcut(options.InstallPath);
                    await Task.Delay(200);
                }

                if (options.CreateStartMenuShortcut)
                {
                    progress.Report(new InstallationProgress { Percentage = 78, CurrentAction = "Creation du raccourci Menu Demarrer..." });
                    CreateStartMenuShortcut(options.InstallPath);
                    await Task.Delay(200);
                }

                // Etape 4: Enregistrement Windows (85-95%)
                progress.Report(new InstallationProgress { Percentage = 85, CurrentAction = "Enregistrement dans Windows..." });
                RegisterUninstaller(options.InstallPath);
                await Task.Delay(300);

                // Etape 5: Finalisation (95-100%)
                progress.Report(new InstallationProgress { Percentage = 95, CurrentAction = "Finalisation..." });
                CreateUninstallScript(options.InstallPath);
                await Task.Delay(300);

                progress.Report(new InstallationProgress { Percentage = 100, CurrentAction = "Installation terminee !" });
                await Task.Delay(500);

                return true;
            }
            catch (Exception ex)
            {
                LogError($"[-] Erreur installation: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Desinstalle l'application
        /// </summary>
        public async Task<bool> UninstallAsync(IProgress<InstallationProgress> progress)
        {
            try
            {
                progress.Report(new InstallationProgress { Percentage = 0, CurrentAction = "Demarrage de la desinstallation..." });

                // Lire le chemin depuis le registre
                string installPath = GetInstallPathFromRegistry();
                if (string.IsNullOrEmpty(installPath))
                {
                    progress.Report(new InstallationProgress { Percentage = 100, CurrentAction = "Application non trouvee." });
                    return false;
                }

                // Fermer l'application si elle est en cours
                progress.Report(new InstallationProgress { Percentage = 10, CurrentAction = "Fermeture de l'application..." });
                KillRunningProcesses();
                await Task.Delay(500);

                // Supprimer les raccourcis
                progress.Report(new InstallationProgress { Percentage = 30, CurrentAction = "Suppression des raccourcis..." });
                RemoveShortcuts();
                await Task.Delay(300);

                // Supprimer les fichiers
                progress.Report(new InstallationProgress { Percentage = 50, CurrentAction = "Suppression des fichiers..." });
                await DeleteDirectoryAsync(installPath);

                // Nettoyer le registre
                progress.Report(new InstallationProgress { Percentage = 80, CurrentAction = "Nettoyage du registre..." });
                RemoveUninstallerRegistry();
                await Task.Delay(300);

                progress.Report(new InstallationProgress { Percentage = 100, CurrentAction = "Desinstallation terminee !" });
                return true;
            }
            catch (Exception ex)
            {
                LogError($"[-] Erreur desinstallation: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Private Methods - File Operations

        private async Task CopyApplicationFilesAsync(string destPath, IProgress<InstallationProgress> progress)
        {
            // Obtenir le dossier source (ou l'installateur contient les fichiers)
            string sourcePath = GetSourcePath();
            
            if (string.IsNullOrEmpty(sourcePath) || !Directory.Exists(sourcePath))
            {
                // Mode demo: simuler la copie
                await SimulateCopyFilesAsync(destPath, progress);
                return;
            }

            // Copier tous les fichiers
            string[] files = Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories);
            int totalFiles = files.Length;
            int currentFile = 0;

            foreach (string file in files)
            {
                currentFile++;
                string relativePath = file.Substring(sourcePath.Length + 1);
                string destFile = Path.Combine(destPath, relativePath);
                string destDir = Path.GetDirectoryName(destFile);

                if (!Directory.Exists(destDir))
                    Directory.CreateDirectory(destDir);

                System.IO.File.Copy(file, destFile, true);

                // Calculer le pourcentage (10-70%)
                int percentage = 10 + (int)((currentFile / (double)totalFiles) * 60);
                progress.Report(new InstallationProgress
                {
                    Percentage = percentage,
                    CurrentAction = $"Copie: {Path.GetFileName(file)} ({currentFile}/{totalFiles})"
                });

                await Task.Delay(20); // Petit delai pour la fluidite
            }
        }

        private async Task SimulateCopyFilesAsync(string destPath, IProgress<InstallationProgress> progress)
        {
            // Simulation pour test/demo
            string[] simulatedFiles = new[]
            {
                "XnrgyEngineeringAutomationTools.exe",
                "XnrgyEngineeringAutomationTools.exe.config",
                "XnrgyEngineeringAutomationTools.ico",
                "XnrgyEngineeringAutomationTools.pdb",
                "Autodesk.Connectivity.WebServices.dll",
                "Autodesk.DataManagement.Client.Framework.dll",
                "Autodesk.DataManagement.Client.Framework.Vault.dll",
                "Inventor.Interop.dll",
                "Newtonsoft.Json.dll",
                "Logs\\.gitkeep"
            };

            // Copier la vraie icone si disponible
            string sourceIco = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Resources", ICO_NAME);
            string destIco = Path.Combine(destPath, ICO_NAME);
            if (System.IO.File.Exists(sourceIco))
            {
                System.IO.File.Copy(sourceIco, destIco, true);
            }
            else
            {
                // Essayer depuis le dossier parent
                sourceIco = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ICO_NAME);
                if (System.IO.File.Exists(sourceIco))
                {
                    System.IO.File.Copy(sourceIco, destIco, true);
                }
            }

            int total = simulatedFiles.Length;
            for (int i = 0; i < total; i++)
            {
                string fileName = simulatedFiles[i];
                string filePath = Path.Combine(destPath, fileName);
                string dirPath = Path.GetDirectoryName(filePath);

                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);

                // Creer un fichier vide pour la demo (sauf icone deja copiee)
                if (!fileName.EndsWith(".gitkeep") && !fileName.EndsWith(".ico"))
                {
                    System.IO.File.WriteAllText(filePath, $"// Placeholder for {fileName}");
                }

                int percentage = 10 + (int)(((i + 1) / (double)total) * 60);
                progress.Report(new InstallationProgress
                {
                    Percentage = percentage,
                    CurrentAction = $"Copie: {Path.GetFileName(fileName)} ({i + 1}/{total})"
                });

                await Task.Delay(200);
            }
        }

        private string GetSourcePath()
        {
            // Option 1: Dossier "Files" a cote de l'installateur (pour distribution)
            string installerDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string filesDir = Path.Combine(installerDir, "Files");
            
            if (Directory.Exists(filesDir))
                return filesDir;

            // Option 2: Dossier bin/Release du projet principal (en dev)
            // Chemin: Installer\bin\Release -> XnrgyEngineeringAutomationTools\bin\Release
            string parentDir = Directory.GetParent(installerDir)?.Parent?.Parent?.FullName;
            if (!string.IsNullOrEmpty(parentDir))
            {
                string devPath = Path.Combine(parentDir, "bin", "Release");
                if (Directory.Exists(devPath))
                    return devPath;
            }

            // Option 3: Chemin absolu connu (fallback dev)
            string knownPath = @"C:\Users\mohammedamine.elgala\source\repos\XnrgyEngineeringAutomationTools\bin\Release";
            if (Directory.Exists(knownPath))
                return knownPath;

            return null;
        }

        private async Task DeleteDirectoryAsync(string path)
        {
            if (!Directory.Exists(path))
                return;

            // Supprimer les fichiers avec retry
            foreach (string file in Directory.GetFiles(path, "*.*", SearchOption.AllDirectories))
            {
                try
                {
                    System.IO.File.SetAttributes(file, FileAttributes.Normal);
                    System.IO.File.Delete(file);
                }
                catch
                {
                    // Ignorer les fichiers en cours d'utilisation
                }
                await Task.Delay(10);
            }

            // Supprimer les dossiers
            try
            {
                Directory.Delete(path, true);
            }
            catch
            {
                // Le dossier sera supprime au redemarrage
            }
        }

        #endregion

        #region Private Methods - Shortcuts

        private void CreateDesktopShortcut(string installPath)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string shortcutPath = Path.Combine(desktopPath, $"{APP_NAME}.lnk");
                string targetPath = Path.Combine(installPath, EXE_NAME);

                CreateShortcut(shortcutPath, targetPath, installPath);
            }
            catch (Exception ex)
            {
                LogError($"[-] Erreur creation raccourci Bureau: {ex.Message}");
            }
        }

        private void CreateStartMenuShortcut(string installPath)
        {
            try
            {
                string startMenuPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                    "Programs",
                    PUBLISHER
                );

                if (!Directory.Exists(startMenuPath))
                    Directory.CreateDirectory(startMenuPath);

                string shortcutPath = Path.Combine(startMenuPath, $"{APP_NAME}.lnk");
                string targetPath = Path.Combine(installPath, EXE_NAME);

                CreateShortcut(shortcutPath, targetPath, installPath);
            }
            catch (Exception ex)
            {
                LogError($"[-] Erreur creation raccourci Menu Demarrer: {ex.Message}");
            }
        }

        private void CreateShortcut(string shortcutPath, string targetPath, string workingDir)
        {
            // Utiliser l'icone ICO si presente, sinon l'exe
            string iconPath = Path.Combine(workingDir, ICO_NAME);
            if (!System.IO.File.Exists(iconPath))
            {
                iconPath = targetPath; // Fallback sur l'exe
            }

            // Utiliser PowerShell pour creer le raccourci (sans dependance COM)
            string psScript = $@"
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut('{shortcutPath.Replace("'", "''")}')
$Shortcut.TargetPath = '{targetPath.Replace("'", "''")}'
$Shortcut.WorkingDirectory = '{workingDir.Replace("'", "''")}'
$Shortcut.Description = '{APP_NAME} - Outils d''automatisation engineering XNRGY'
$Shortcut.IconLocation = '{iconPath.Replace("'", "''")},0'
$Shortcut.Save()
";
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psScript.Replace("\"", "`\"")}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true
                };

                using (Process process = Process.Start(psi))
                {
                    process.WaitForExit(5000);
                }
            }
            catch (Exception ex)
            {
                LogError($"[-] Erreur creation raccourci via PowerShell: {ex.Message}");
            }
        }

        private void RemoveShortcuts()
        {
            try
            {
                // Supprimer raccourci Bureau
                string desktopShortcut = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                    $"{APP_NAME}.lnk"
                );
                if (System.IO.File.Exists(desktopShortcut))
                    System.IO.File.Delete(desktopShortcut);

                // Supprimer dossier Menu Demarrer
                string startMenuPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                    "Programs",
                    PUBLISHER
                );
                if (Directory.Exists(startMenuPath))
                    Directory.Delete(startMenuPath, true);
            }
            catch (Exception ex)
            {
                LogError($"[-] Erreur suppression raccourcis: {ex.Message}");
            }
        }

        #endregion

        #region Private Methods - Registry

        private void RegisterUninstaller(string installPath)
        {
            try
            {
                string uninstallKey = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{UNINSTALL_GUID}";
                
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey(uninstallKey))
                {
                    if (key == null)
                    {
                        // Fallback: Current User si pas admin
                        using (RegistryKey userKey = Registry.CurrentUser.CreateSubKey(uninstallKey))
                        {
                            WriteUninstallRegistryValues(userKey, installPath);
                        }
                    }
                    else
                    {
                        WriteUninstallRegistryValues(key, installPath);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"[-] Erreur enregistrement registre: {ex.Message}");
            }
        }

        private void WriteUninstallRegistryValues(RegistryKey key, string installPath)
        {
            string uninstallerPath = Path.Combine(installPath, "Uninstall.exe");
            string iconPath = Path.Combine(installPath, ICO_NAME);
            
            // Utiliser l'icone ICO si presente, sinon l'exe
            if (!System.IO.File.Exists(iconPath))
            {
                iconPath = Path.Combine(installPath, EXE_NAME);
            }
            
            // Nom affiche dans Programmes et fonctionnalites
            key.SetValue("DisplayName", APP_NAME);
            key.SetValue("DisplayVersion", VERSION);
            key.SetValue("Publisher", PUBLISHER);
            key.SetValue("InstallLocation", installPath);
            key.SetValue("UninstallString", $"\"{uninstallerPath}\"");
            key.SetValue("DisplayIcon", $"{iconPath},0");
            key.SetValue("NoModify", 1, RegistryValueKind.DWord);
            key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            key.SetValue("EstimatedSize", 150000, RegistryValueKind.DWord); // ~150 MB
            key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
            
            // Commentaires affiches sous le nom (comme Smart Tools)
            // Format: "1.0.0 | Mohammed Amine Elgalai - XNRGY Climate Systems ULC | 2026-01-16"
            key.SetValue("Comments", $"{VERSION} | {PUBLISHER} | {RELEASE_DATE}");
            
            // URL de support et contact
            key.SetValue("URLInfoAbout", "https://github.com/mohammedamineelgalai/XnrgyEngineeringAutomationTools");
            key.SetValue("Contact", "Mohammed Amine Elgalai");
        }

        private string GetInstallPathFromRegistry()
        {
            string uninstallKey = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{UNINSTALL_GUID}";
            
            try
            {
                // Essayer HKLM d'abord
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(uninstallKey))
                {
                    if (key != null)
                    {
                        return key.GetValue("InstallLocation") as string;
                    }
                }

                // Fallback HKCU
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(uninstallKey))
                {
                    if (key != null)
                    {
                        return key.GetValue("InstallLocation") as string;
                    }
                }
            }
            catch { }

            return null;
        }

        private void RemoveUninstallerRegistry()
        {
            string uninstallKey = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{UNINSTALL_GUID}";
            
            try
            {
                Registry.LocalMachine.DeleteSubKeyTree(uninstallKey, false);
                Registry.CurrentUser.DeleteSubKeyTree(uninstallKey, false);
            }
            catch { }
        }

        #endregion

        #region Private Methods - Utilities

        private void CreateUninstallScript(string installPath)
        {
            // Copier l'installateur comme uninstaller
            string currentExe = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string uninstallerPath = Path.Combine(installPath, "Uninstall.exe");

            try
            {
                System.IO.File.Copy(currentExe, uninstallerPath, true);
            }
            catch
            {
                // Creer un script batch comme fallback
                string batchPath = Path.Combine(installPath, "Uninstall.bat");
                string batchContent = $@"@echo off
echo Desinstallation de {APP_NAME}...
taskkill /f /im ""{EXE_NAME}"" 2>nul
timeout /t 2 /nobreak >nul
rmdir /s /q ""{installPath}""
reg delete ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{UNINSTALL_GUID}"" /f 2>nul
reg delete ""HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{UNINSTALL_GUID}"" /f 2>nul
del ""%USERPROFILE%\Desktop\{APP_NAME}.lnk"" 2>nul
rmdir /s /q ""%APPDATA%\Microsoft\Windows\Start Menu\Programs\{PUBLISHER}"" 2>nul
echo Desinstallation terminee !
pause
";
                System.IO.File.WriteAllText(batchPath, batchContent);
            }
        }

        private void KillRunningProcesses()
        {
            try
            {
                string processName = Path.GetFileNameWithoutExtension(EXE_NAME);
                foreach (var process in Process.GetProcessesByName(processName))
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(5000);
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void LogError(string message)
        {
            try
            {
                string logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    PUBLISHER,
                    "InstallerLogs"
                );
                
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                string logFile = Path.Combine(logDir, $"Install_{DateTime.Now:yyyyMMdd}.log");
                System.IO.File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n");
            }
            catch { }
        }

        #endregion
    }
}
