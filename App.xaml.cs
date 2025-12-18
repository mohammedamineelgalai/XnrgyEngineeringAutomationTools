using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Threading;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            base.OnStartup(e);
            AddInventorToPath();
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
