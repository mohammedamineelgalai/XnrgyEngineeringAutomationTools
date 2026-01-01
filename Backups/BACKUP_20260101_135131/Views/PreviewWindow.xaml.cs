using System;
using System.Collections.Generic;
using System.Windows;
using XnrgyEngineeringAutomationTools.Models;

namespace XnrgyEngineeringAutomationTools.Views
{
    /// <summary>
    /// Fenêtre de prévisualisation moderne pour la création de module
    /// </summary>
    public partial class PreviewWindow : Window
    {
        /// <summary>
        /// Résultat de la fenêtre - true si l'utilisateur a confirmé
        /// </summary>
        public bool IsConfirmed { get; private set; }

        public PreviewWindow()
        {
            InitializeComponent();
            IsConfirmed = false;
        }

        /// <summary>
        /// Configure la prévisualisation avec les données du module
        /// </summary>
        public void SetPreviewData(
            string project,
            string reference,
            string module,
            string fullProjectNumber,
            string destinationPath,
            string initialeDessinateur,
            string initialeCoDessinateur,
            DateTime creationDate,
            string jobTitle,
            IEnumerable<FileRenameItem> files,
            int fileCount)
        {
            // Infos projet
            TxtPreviewProject.Text = string.IsNullOrEmpty(project) ? "-" : project;
            TxtPreviewRefModule.Text = $"REF{reference} / M{module}";
            TxtPreviewFullNumber.Text = fullProjectNumber;
            TxtPreviewDestination.Text = destinationPath;

            // Propriétés
            TxtPreviewDessinateur.Text = string.IsNullOrEmpty(initialeDessinateur) ? "-" : initialeDessinateur;
            TxtPreviewCoDessinateur.Text = string.IsNullOrEmpty(initialeCoDessinateur) ? "(Non défini)" : initialeCoDessinateur;
            TxtPreviewDate.Text = creationDate.ToString("dd/MM/yyyy");
            TxtPreviewJobTitle.Text = string.IsNullOrEmpty(jobTitle) ? "(Non défini)" : jobTitle;

            // Fichiers
            TxtPreviewFileCount.Text = $"{fileCount} fichiers sélectionnés";
            LstPreviewFiles.ItemsSource = files;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            IsConfirmed = false;
            DialogResult = false;
            Close();
        }

        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            IsConfirmed = true;
            DialogResult = true;
            Close();
        }
    }
}
