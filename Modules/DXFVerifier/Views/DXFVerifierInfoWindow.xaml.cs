// ============================================================================
// DXFVerifierInfoWindow.xaml.cs
// DXF Verifier Info Window - Modern Info Dialog
// Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
// ============================================================================

using System.Windows;

namespace XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Views
{
    /// <summary>
    /// Fenetre d'information moderne pour DXF Verifier
    /// Design unifie avec SmartToolsInfoWindow
    /// </summary>
    public partial class DXFVerifierInfoWindow : Window
    {
        public DXFVerifierInfoWindow()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
