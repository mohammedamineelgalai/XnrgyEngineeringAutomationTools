using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using XnrgyEngineeringAutomationTools.Models;

namespace XnrgyEngineeringAutomationTools.Views
{
    /// <summary>
    /// Fenêtre de sélection de module XNRGY
    /// </summary>
    public partial class ModuleSelectionWindow : Window
    {
        public ModuleInfo SelectedModule { get; private set; }

        public ModuleSelectionWindow(List<ModuleInfo> modules)
        {
            InitializeComponent();

            // Pattern matching (IDE0019)
            if (this.FindName("ModulesListBox") is not System.Windows.Controls.ListBox modulesListBox) return;

            // Remplir la liste
            foreach (var module in modules)
            {
                modulesListBox.Items.Add(module);
            }

            // Sélectionner le premier élément par défaut
            if (modulesListBox.Items.Count > 0)
            {
                modulesListBox.SelectedIndex = 0;
            }
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            // Pattern matching (IDE0019)
            if (this.FindName("ModulesListBox") is System.Windows.Controls.ListBox modulesListBox 
                && modulesListBox.SelectedItem is ModuleInfo selected)
            {
                SelectedModule = selected;
                DialogResult = true;
                Close();
            }
            // Ne rien faire si pas de sélection - l'utilisateur doit sélectionner un module ou annuler
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void ModulesListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Double-clic pour sélectionner directement - Pattern matching (IDE0019)
            if (this.FindName("ModulesListBox") is System.Windows.Controls.ListBox modulesListBox 
                && modulesListBox.SelectedItem is ModuleInfo selected)
            {
                SelectedModule = selected;
                DialogResult = true;
                Close();
            }
        }
    }
}
