using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Views;

namespace XnrgyEngineeringAutomationTools
{
    public partial class App : Application
    {
        private DeviceTrackingService _deviceTracker;

        protected override async void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            base.OnStartup(e);
            AddInventorToPath();

            // Verification Firebase au demarrage
            bool canContinue = await CheckFirebaseConfigurationAsync();
            if (!canContinue)
            {
                Shutdown();
                return;
            }

            // Enregistrer l'appareil dans Firebase pour le tracking
            _deviceTracker = new DeviceTrackingService();
            await _deviceTracker.RegisterDeviceAsync();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Desenregistrer l'appareil a la fermeture
            _deviceTracker?.Dispose();
            base.OnExit(e);
        }

        /// <summary>
        /// Verifie la configuration Firebase (kill switch, maintenance, mises a jour)
        /// </summary>
        private async Task<bool> CheckFirebaseConfigurationAsync()
        {
            try
            {
                var result = await FirebaseRemoteConfigService.CheckConfigurationAsync();

                if (!result.Success)
                {
                    // Erreur de connexion - continuer en mode hors ligne
                    return true;
                }

                // 1. Kill Switch - Bloque completement l'application
                if (result.KillSwitchActive)
                {
                    FirebaseAlertWindow.ShowKillSwitch(result.KillSwitchMessage);
                    return false;
                }

                // 2. Utilisateur desactive - Bloque cet utilisateur specifique
                if (result.UserDisabled)
                {
                    FirebaseAlertWindow.ShowUserDisabled(result.UserDisabledMessage);
                    return false;
                }

                // 3. Mode Maintenance - Bloque temporairement
                if (result.MaintenanceMode)
                {
                    FirebaseAlertWindow.ShowMaintenance(result.MaintenanceMessage);
                    return false;
                }

                // 4. Mise a jour disponible
                if (result.UpdateAvailable)
                {
                    string currentVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
                    currentVersion = $"{currentVersion.Split('.')[0]}.{currentVersion.Split('.')[1]}.{currentVersion.Split('.')[2]}";

                    var (shouldContinue, shouldDownload) = FirebaseAlertWindow.ShowUpdateAvailable(
                        currentVersion,
                        result.LatestVersion,
                        result.Changelog,
                        result.DownloadUrl,
                        result.ForceUpdate);

                    // Si mise a jour forcee, bloquer l'application
                    if (result.ForceUpdate)
                    {
                        return false;
                    }

                    // Sinon, continuer selon le choix de l'utilisateur
                    if (!shouldContinue) return false;
                }

                // 5. Message broadcast - Afficher sans bloquer (sauf si type "error")
                if (result.HasBroadcastMessage)
                {
                    bool shouldBlock = FirebaseAlertWindow.ShowBroadcastMessage(
                        result.BroadcastTitle,
                        result.BroadcastMessage,
                        result.BroadcastType);
                    
                    if (shouldBlock) return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur verification Firebase: {ex.Message}", Logger.LogLevel.WARNING);
                // En cas d'erreur, continuer normalement
                return true;
            }
        }

        private void AddInventorToPath()
        {
            try
            {
                string inventorPath = @"C:\Program Files\Autodesk\Inventor 2026\Bin";
                if (Directory.Exists(inventorPath))
                {
                    string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                    if (!currentPath.Contains(inventorPath))
                    {
                        Environment.SetEnvironmentVariable("PATH", inventorPath + ";" + currentPath);
                    }
                }
            }
            catch { }
        }

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            string assemblyName = new AssemblyName(args.Name).Name ?? "";
            string[] searchPaths = new[]
            {
                @"C:\Program Files\Autodesk\Vault Client 2026\Explorer",
                @"C:\Program Files\Autodesk\Autodesk Vault 2026 SDK\bin\x64",
                @"C:\Program Files\Autodesk\Inventor 2026\Bin"
            };
            foreach (var path in searchPaths)
            {
                string dllPath = Path.Combine(path, assemblyName + ".dll");
                if (File.Exists(dllPath))
                {
                    try { return Assembly.LoadFrom(dllPath); }
                    catch { }
                }
            }
            return null;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = (Exception)e.ExceptionObject;
            MessageBox.Show("Erreur critique: " + ex.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show("Erreur: " + e.Exception.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }
    }
}
