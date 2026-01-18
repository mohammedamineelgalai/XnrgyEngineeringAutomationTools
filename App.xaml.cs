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
        private AutoUpdateService _autoUpdateService;

        protected override async void OnStartup(StartupEventArgs e)
        {
            // Configurer les handlers d'erreur AVANT tout
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            
            // NE PAS appeler base.OnStartup ici car on a supprime StartupUri
            AddInventorToPath();

            // VERIFICATION FIREBASE OBLIGATOIRE AVANT DE LANCER L'APPLICATION
            bool canContinue = await CheckFirebaseConfigurationAsync();
            if (!canContinue)
            {
                // Fermer l'application immediatement
                Environment.Exit(0);
                return;
            }

            // Si Firebase OK, enregistrer l'appareil
            _deviceTracker = new DeviceTrackingService();
            await _deviceTracker.RegisterDeviceAsync();

            // Initialiser le service Firebase Audit (session + heartbeat)
            await FirebaseAuditService.Instance.InitializeAsync();

            // Demarrer le service de mise a jour automatique
            _autoUpdateService = new AutoUpdateService();
            _autoUpdateService.UpdateAvailable += OnUpdateAvailable;
            _autoUpdateService.Start();

            // MAINTENANT on peut lancer la fenetre principale
            var mainWindow = new MainWindow();
            MainWindow = mainWindow;
            mainWindow.Show();
            
            // Changer le mode de fermeture maintenant que la fenetre est ouverte
            ShutdownMode = ShutdownMode.OnMainWindowClose;
        }

        /// <summary>
        /// Gere la notification de mise a jour disponible (verification periodique)
        /// </summary>
        private void OnUpdateAvailable(object sender, UpdateAvailableEventArgs e)
        {
            // Afficher la notification de mise a jour
            var (shouldContinue, shouldDownload) = FirebaseAlertWindow.ShowUpdateAvailable(
                e.CurrentVersion,
                e.NewVersion,
                e.Changelog,
                e.DownloadUrl,
                e.IsForced);

            // Si mise a jour forcee et l'utilisateur refuse, forcer la fermeture
            if (e.IsForced && !shouldDownload)
            {
                Environment.Exit(0);
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Enregistrer la fin de session Firebase - SYNCHRONE pour garantir l'execution
            try
            {
                // Utiliser Wait() avec timeout pour garantir que la requete part avant fermeture
                var task = FirebaseAuditService.Instance.RegisterSessionEndAsync();
                task.Wait(TimeSpan.FromSeconds(5));
            }
            catch
            {
                // Silencieux - ne pas bloquer la fermeture
            }

            // Arreter les services
            _autoUpdateService?.Stop();
            _autoUpdateService?.Dispose();
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

                // 2a. Device suspendu - Bloque ce poste de travail entier
                if (result.DeviceDisabled)
                {
                    FirebaseAlertWindow.ShowDeviceDisabled(result.DeviceDisabledMessage, result.DeviceDisabledReason);
                    return false;
                }

                // 2b. Utilisateur suspendu SUR CE DEVICE - Bloque cet utilisateur sur ce poste
                if (result.DeviceUserDisabled)
                {
                    FirebaseAlertWindow.ShowDeviceUserDisabled(result.DeviceUserDisabledMessage, result.DeviceUserDisabledReason);
                    return false;
                }

                // 3. Utilisateur desactive globalement (optionnel - pour compatibilite)
                if (result.UserDisabled)
                {
                    FirebaseAlertWindow.ShowUserDisabled(result.UserDisabledMessage);
                    return false;
                }

                // 4. Mode Maintenance - Bloque temporairement
                if (result.MaintenanceMode)
                {
                    FirebaseAlertWindow.ShowMaintenance(result.MaintenanceMessage);
                    return false;
                }

                // 5. Mise a jour disponible
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

                // 6. Message broadcast - Afficher sans bloquer (sauf si type "error")
                if (result.HasBroadcastMessage)
                {
                    bool shouldBlock = FirebaseAlertWindow.ShowBroadcastMessage(
                        result.BroadcastTitle,
                        result.BroadcastMessage,
                        result.BroadcastType);
                    
                    if (shouldBlock) return false;
                }

                // 7. Message de bienvenue - Afficher au demarrage (informatif seulement)
                if (result.HasWelcomeMessage)
                {
                    FirebaseAlertWindow.ShowWelcomeMessage(
                        result.WelcomeTitle,
                        result.WelcomeMessage,
                        result.WelcomeType);
                    // Ne bloque jamais - continue toujours
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
