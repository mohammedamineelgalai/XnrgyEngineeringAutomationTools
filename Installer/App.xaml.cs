using System;
using System.Windows;

namespace XnrgyInstaller
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Verifier si l'application est deja installee (pour mode uninstall)
            if (e.Args.Length > 0 && e.Args[0] == "/uninstall")
            {
                // Mode desinstallation
                Environment.SetEnvironmentVariable("XNRGY_INSTALLER_MODE", "uninstall");
            }
        }
    }
}
