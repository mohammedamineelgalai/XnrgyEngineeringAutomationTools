using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using XnrgyEngineeringAutomationTools.Views;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service de mise a jour automatique professionnelle
    /// - Verification periodique en arriere-plan
    /// - Telechargement automatique
    /// - Installation silencieuse et redemarrage
    /// </summary>
    public class AutoUpdateService : IDisposable
    {
        // Intervalle de verification (5 minutes)
        private const int CHECK_INTERVAL_MS = 5 * 60 * 1000;
        
        // Dossier temporaire pour les telechargements
        private static readonly string UPDATE_FOLDER = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "XnrgyEngineeringAutomationTools", "Updates");

        private readonly System.Timers.Timer _checkTimer;
        private readonly HttpClient _httpClient;
        private bool _isChecking;
        private bool _disposed;
        private string _lastNotifiedVersion;

        // Singleton
        private static AutoUpdateService _instance;
        public static AutoUpdateService Instance => _instance;

        // Evenement pour notifier l'UI
        public event EventHandler<UpdateAvailableEventArgs> UpdateAvailable;
        public event EventHandler<UpdateProgressEventArgs> DownloadProgressChanged;
        public event EventHandler<UpdateCompletedEventArgs> UpdateCompleted;

        public AutoUpdateService()
        {
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromMinutes(10) // Timeout long pour gros fichiers
            };

            // Creer le dossier de telechargement
            Directory.CreateDirectory(UPDATE_FOLDER);

            // Timer de verification periodique
            _checkTimer = new System.Timers.Timer(CHECK_INTERVAL_MS);
            _checkTimer.Elapsed += async (s, e) => await CheckForUpdatesAsync(silent: true);
            _checkTimer.AutoReset = true;

            _instance = this;
        }

        /// <summary>
        /// Demarre le service de verification automatique
        /// </summary>
        public void Start()
        {
            _checkTimer.Start();
            Logger.Log("[+] Service de mise a jour automatique demarre");
        }

        /// <summary>
        /// Arrete le service
        /// </summary>
        public void Stop()
        {
            _checkTimer.Stop();
        }

        /// <summary>
        /// Verifie si une mise a jour est disponible
        /// </summary>
        /// <param name="silent">Si true, ne notifie que pour les mises a jour forcees</param>
        public async Task<bool> CheckForUpdatesAsync(bool silent = false)
        {
            if (_isChecking) return false;
            _isChecking = true;

            try
            {
                var result = await FirebaseRemoteConfigService.CheckConfigurationAsync();
                
                if (!result.Success) return false;

                // Verifier si mise a jour disponible
                if (result.UpdateAvailable)
                {
                    string currentVersion = GetCurrentVersion();
                    
                    // Ne pas re-notifier pour la meme version (sauf si force)
                    if (!result.ForceUpdate && _lastNotifiedVersion == result.LatestVersion && silent)
                    {
                        return false;
                    }

                    _lastNotifiedVersion = result.LatestVersion;

                    // Si mise a jour forcee OU mode non-silencieux, notifier
                    if (result.ForceUpdate || !silent)
                    {
                        Logger.Log($"[+] Mise a jour disponible: {currentVersion} -> {result.LatestVersion} (Force: {result.ForceUpdate})");
                        
                        // Notifier sur le thread UI
                        Application.Current?.Dispatcher?.Invoke(() =>
                        {
                            OnUpdateAvailable(new UpdateAvailableEventArgs
                            {
                                CurrentVersion = currentVersion,
                                NewVersion = result.LatestVersion,
                                Changelog = result.Changelog,
                                DownloadUrl = result.DownloadUrl,
                                IsForced = result.ForceUpdate
                            });
                        });

                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification mise a jour: {ex.Message}", Logger.LogLevel.WARNING);
                return false;
            }
            finally
            {
                _isChecking = false;
            }
        }

        /// <summary>
        /// Telecharge et installe la mise a jour automatiquement
        /// </summary>
        public async Task<bool> DownloadAndInstallUpdateAsync(string downloadUrl, string newVersion)
        {
            string installerPath = null;

            try
            {
                Logger.Log($"[>] Demarrage du telechargement: {downloadUrl}");

                // Nettoyer les anciens fichiers
                CleanupOldUpdates();

                // Determiner le nom du fichier
                string fileName = $"XnrgyEngineeringAutomationTools_Setup_{newVersion}.exe";
                installerPath = Path.Combine(UPDATE_FOLDER, fileName);

                // Telecharger avec progression
                using (var response = await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead))
                {
                    response.EnsureSuccessStatusCode();

                    long totalBytes = response.Content.Headers.ContentLength ?? -1;
                    long downloadedBytes = 0;

                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                    {
                        byte[] buffer = new byte[8192];
                        int bytesRead;

                        while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead);
                            downloadedBytes += bytesRead;

                            // Notifier la progression
                            if (totalBytes > 0)
                            {
                                int progressPercent = (int)((downloadedBytes * 100) / totalBytes);
                                OnDownloadProgressChanged(new UpdateProgressEventArgs
                                {
                                    BytesDownloaded = downloadedBytes,
                                    TotalBytes = totalBytes,
                                    ProgressPercent = progressPercent
                                });
                            }
                        }
                    }
                }

                Logger.Log($"[+] Telechargement termine: {installerPath}");

                // Verifier que le fichier existe et a une taille raisonnable
                var fileInfo = new FileInfo(installerPath);
                if (!fileInfo.Exists || fileInfo.Length < 100000) // Min 100 KB
                {
                    Logger.Log("[-] Fichier telecharge invalide", Logger.LogLevel.ERROR);
                    return false;
                }

                // Lancer l'installation
                return await LaunchInstallerAndExitAsync(installerPath);
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur telechargement/installation: {ex.Message}", Logger.LogLevel.ERROR);
                
                // Nettoyer le fichier partiel
                if (!string.IsNullOrEmpty(installerPath) && File.Exists(installerPath))
                {
                    try { File.Delete(installerPath); } catch { }
                }

                OnUpdateCompleted(new UpdateCompletedEventArgs
                {
                    Success = false,
                    ErrorMessage = ex.Message
                });

                return false;
            }
        }

        /// <summary>
        /// Lance l'installateur et ferme l'application actuelle
        /// </summary>
        private async Task<bool> LaunchInstallerAndExitAsync(string installerPath)
        {
            try
            {
                Logger.Log($"[>] Lancement de l'installateur: {installerPath}");

                // Creer un script batch pour:
                // 1. Attendre que l'app se ferme
                // 2. Lancer l'installateur en mode silencieux
                // 3. Relancer l'application
                string batchPath = Path.Combine(UPDATE_FOLDER, "update_launcher.bat");
                string appPath = Assembly.GetExecutingAssembly().Location;

                string batchContent = $@"@echo off
title XNRGY Update in Progress...
echo.
echo ============================================
echo   XNRGY Engineering Automation Tools
echo   Installation de la mise a jour...
echo ============================================
echo.
echo Fermeture de l'application en cours...
timeout /t 3 /nobreak > nul

:waitloop
tasklist /FI ""IMAGENAME eq XnrgyEngineeringAutomationTools.exe"" 2>NUL | find /I /N ""XnrgyEngineeringAutomationTools.exe"">NUL
if ""%ERRORLEVEL%""==""0"" (
    echo En attente de la fermeture...
    timeout /t 1 /nobreak > nul
    goto waitloop
)

echo.
echo Lancement de l'installateur...
echo.
start /wait """" ""{installerPath}"" /SILENT /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS

echo.
echo Installation terminee!
echo Redemarrage de l'application...
timeout /t 2 /nobreak > nul

start """" ""{appPath}""

echo.
echo Nettoyage...
del ""{installerPath}"" 2>nul
del ""%~f0"" 2>nul
";

                File.WriteAllText(batchPath, batchContent, System.Text.Encoding.Default);

                // Lancer le batch en arriere-plan
                var startInfo = new ProcessStartInfo
                {
                    FileName = batchPath,
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal
                };

                Process.Start(startInfo);

                Logger.Log("[+] Script de mise a jour lance, fermeture de l'application...");

                // Notifier la completion
                OnUpdateCompleted(new UpdateCompletedEventArgs { Success = true });

                // Fermer l'application apres un court delai
                await Task.Delay(1000);
                
                Application.Current?.Dispatcher?.Invoke(() =>
                {
                    Application.Current.Shutdown();
                });

                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur lancement installateur: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Nettoie les anciens fichiers de mise a jour
        /// </summary>
        private void CleanupOldUpdates()
        {
            try
            {
                if (Directory.Exists(UPDATE_FOLDER))
                {
                    foreach (var file in Directory.GetFiles(UPDATE_FOLDER, "*.exe"))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch { }
                    }
                    foreach (var file in Directory.GetFiles(UPDATE_FOLDER, "*.bat"))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Obtient la version actuelle de l'application
        /// </summary>
        private string GetCurrentVersion()
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                return version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "1.0.0";
            }
            catch
            {
                return "1.0.0";
            }
        }

        #region Events

        protected virtual void OnUpdateAvailable(UpdateAvailableEventArgs e)
        {
            UpdateAvailable?.Invoke(this, e);
        }

        protected virtual void OnDownloadProgressChanged(UpdateProgressEventArgs e)
        {
            Application.Current?.Dispatcher?.Invoke(() =>
            {
                DownloadProgressChanged?.Invoke(this, e);
            });
        }

        protected virtual void OnUpdateCompleted(UpdateCompletedEventArgs e)
        {
            Application.Current?.Dispatcher?.Invoke(() =>
            {
                UpdateCompleted?.Invoke(this, e);
            });
        }

        #endregion

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            _checkTimer?.Stop();
            _checkTimer?.Dispose();
            _httpClient?.Dispose();
        }
    }

    #region Event Args

    public class UpdateAvailableEventArgs : EventArgs
    {
        public string CurrentVersion { get; set; }
        public string NewVersion { get; set; }
        public string Changelog { get; set; }
        public string DownloadUrl { get; set; }
        public bool IsForced { get; set; }
    }

    public class UpdateProgressEventArgs : EventArgs
    {
        public long BytesDownloaded { get; set; }
        public long TotalBytes { get; set; }
        public int ProgressPercent { get; set; }
    }

    public class UpdateCompletedEventArgs : EventArgs
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }

    #endregion
}
