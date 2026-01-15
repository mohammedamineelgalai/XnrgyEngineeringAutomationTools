// ============================================================================
// SmartToolsInfoWindow.xaml.cs
// Smart Tools Info Window - Modern WPF Info Dialog
// Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
// ============================================================================

using System;
using System.Security.Principal;
using System.Windows;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenetre d'information moderne pour Smart Tools
    /// Migration depuis VB.NET InfoCommand.vb
    /// </summary>
    public partial class SmartToolsInfoWindow : Window
    {
        public SmartToolsInfoWindow()
        {
            InitializeComponent();
            LoadSystemInfo();
        }

        private void LoadSystemInfo()
        {
            try
            {
                // Utilisateur
                string currentUser = WindowsIdentity.GetCurrent().Name;
                if (currentUser.Contains("\\"))
                {
                    var parts = currentUser.Split('\\');
                    TxtUser.Text = parts[parts.Length - 1];
                }
                else
                {
                    TxtUser.Text = currentUser;
                }

                // Machine
                TxtMachine.Text = Environment.MachineName;

                // Date
                TxtDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

                // Inventor Version
                try
                {
                    var inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType != null)
                    {
                        dynamic inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                        if (inventorApp != null)
                        {
                            TxtInventor.Text = inventorApp.SoftwareVersion.DisplayVersion;
                        }
                    }
                }
                catch
                {
                    TxtInventor.Text = "Non connecte";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur LoadSystemInfo: {ex.Message}");
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
