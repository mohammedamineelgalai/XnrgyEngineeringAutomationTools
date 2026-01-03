using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenêtre d'options pour l'export IAM vers IPT/STEP
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public partial class ExportOptionsWindow : Window
    {
        /// <summary>
        /// Résultat de la fenêtre (OK ou Cancel)
        /// </summary>
        public bool IsConfirmed { get; private set; } = false;

        /// <summary>
        /// Format d'export sélectionné
        /// </summary>
        public ExportFormat SelectedFormat => RbExportIPT.IsChecked == true ? ExportFormat.IPT : ExportFormat.STEP;

        /// <summary>
        /// Chemin de destination
        /// </summary>
        public string DestinationPath => TxtDestinationPath.Text;

        /// <summary>
        /// Nom du fichier de sortie (sans extension)
        /// </summary>
        public string OutputFileName => TxtOutputFileName.Text;

        /// <summary>
        /// Masquer les éléments de référence avant export
        /// </summary>
        public bool HideReferences => ChkHideReferences.IsChecked == true;

        /// <summary>
        /// Activer la représentation par défaut
        /// </summary>
        public bool ActivateDefaultRepresentation => ChkActivateDefaultRep.IsChecked == true;

        /// <summary>
        /// Ouvrir le fichier après export
        /// </summary>
        public bool OpenAfterExport => ChkOpenAfterExport.IsChecked == true;

        /// <summary>
        /// Chemin complet du fichier de sortie
        /// </summary>
        public string FullOutputPath
        {
            get
            {
                string ext = SelectedFormat == ExportFormat.IPT ? ".ipt" : ".stp";
                return Path.Combine(DestinationPath, OutputFileName + ext);
            }
        }

        public ExportOptionsWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Initialise la fenêtre avec les valeurs par défaut basées sur le document source
        /// </summary>
        /// <param name="sourceFileName">Nom du fichier source (ex: Module_M06.iam)</param>
        /// <param name="sourcePath">Chemin complet du fichier source</param>
        public void Initialize(string sourceFileName, string sourcePath)
        {
            TxtSourceFile.Text = sourceFileName;

            // Extraire le chemin du projet si possible
            // Format attendu: C:\Vault\Engineering\Projects\XXXXX\REFXX\MXX\...
            string destinationPath = ExtractProjectPath(sourcePath);
            TxtDestinationPath.Text = destinationPath;

            // Générer un nom de fichier suggéré
            string suggestedName = GenerateSuggestedFileName(sourcePath, sourceFileName);
            TxtOutputFileName.Text = suggestedName;
        }

        /// <summary>
        /// Extrait le chemin du projet depuis le chemin source
        /// </summary>
        private string ExtractProjectPath(string sourcePath)
        {
            try
            {
                // Chercher le pattern C:\Vault\Engineering\Projects\XXXXX\REFXX\MXX
                string[] parts = sourcePath.Split(Path.DirectorySeparatorChar);
                
                for (int i = 0; i < parts.Length - 2; i++)
                {
                    if (parts[i].Equals("Projects", StringComparison.OrdinalIgnoreCase))
                    {
                        // Reconstruire le chemin jusqu'au module
                        int endIndex = Math.Min(i + 4, parts.Length);
                        return string.Join(Path.DirectorySeparatorChar.ToString(), parts, 0, endIndex);
                    }
                }

                // Fallback: dossier du fichier source
                return Path.GetDirectoryName(sourcePath) ?? @"C:\Vault\Engineering\Projects";
            }
            catch
            {
                return @"C:\Vault\Engineering\Projects";
            }
        }

        /// <summary>
        /// Génère un nom de fichier suggéré au format PROJET-REF-MODULE
        /// </summary>
        private string GenerateSuggestedFileName(string sourcePath, string sourceFileName)
        {
            try
            {
                // Extraire les infos depuis le chemin
                // Format: C:\Vault\Engineering\Projects\10516\REF01\M06\...
                string[] parts = sourcePath.Split(Path.DirectorySeparatorChar);
                
                string project = "";
                string reference = "";
                string module = "";

                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i].Equals("Projects", StringComparison.OrdinalIgnoreCase) && i + 3 < parts.Length)
                    {
                        project = parts[i + 1];
                        reference = parts[i + 2].Replace("REF", "").Replace("ref", "");
                        module = parts[i + 3];
                        break;
                    }
                }

                if (!string.IsNullOrEmpty(project) && !string.IsNullOrEmpty(reference))
                {
                    return $"{project}-{reference}-{module}";
                }

                // Fallback: utiliser le nom du fichier source sans extension
                return Path.GetFileNameWithoutExtension(sourceFileName);
            }
            catch
            {
                return Path.GetFileNameWithoutExtension(sourceFileName);
            }
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Sélectionner le dossier de destination",
                SelectedPath = TxtDestinationPath.Text,
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TxtDestinationPath.Text = dialog.SelectedPath;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            IsConfirmed = false;
            DialogResult = false;
            Close();
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            if (string.IsNullOrWhiteSpace(TxtDestinationPath.Text))
            {
                Shared.Views.XnrgyMessageBox.ShowError(
                    "Veuillez spécifier un dossier de destination.",
                    "Validation", this);
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtOutputFileName.Text))
            {
                Shared.Views.XnrgyMessageBox.ShowError(
                    "Veuillez spécifier un nom de fichier de sortie.",
                    "Validation", this);
                return;
            }

            // Vérifier si le dossier existe
            if (!Directory.Exists(TxtDestinationPath.Text))
            {
                var result = Shared.Views.XnrgyMessageBox.Confirm(
                    $"Le dossier '{TxtDestinationPath.Text}' n'existe pas.\n\nVoulez-vous le créer?",
                    "Créer le dossier?", this);

                if (result)
                {
                    try
                    {
                        Directory.CreateDirectory(TxtDestinationPath.Text);
                    }
                    catch (Exception ex)
                    {
                        Shared.Views.XnrgyMessageBox.ShowError(
                            $"Impossible de créer le dossier:\n{ex.Message}",
                            "Erreur", this);
                        return;
                    }
                }
                else
                {
                    return;
                }
            }

            // Vérifier si le fichier existe déjà
            if (File.Exists(FullOutputPath))
            {
                var result = Shared.Views.XnrgyMessageBox.Confirm(
                    $"Le fichier '{Path.GetFileName(FullOutputPath)}' existe déjà.\n\nVoulez-vous le remplacer?",
                    "Fichier existant", this);

                if (!result)
                {
                    return;
                }
            }

            IsConfirmed = true;
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// Affiche la fenêtre d'options et retourne le résultat
        /// </summary>
        /// <param name="sourceFileName">Nom du fichier source</param>
        /// <param name="sourcePath">Chemin complet du fichier source</param>
        /// <param name="owner">Fenêtre parente</param>
        /// <returns>Instance configurée ou null si annulé</returns>
        public static ExportOptionsWindow? ShowOptions(string sourceFileName, string sourcePath, Window? owner = null)
        {
            var window = new ExportOptionsWindow();
            window.Initialize(sourceFileName, sourcePath);

            if (owner != null)
            {
                window.Owner = owner;
            }

            if (window.ShowDialog() == true && window.IsConfirmed)
            {
                return window;
            }

            return null;
        }
    }

    /// <summary>
    /// Format d'export disponibles
    /// </summary>
    public enum ExportFormat
    {
        /// <summary>
        /// Fichier Inventor Part (.ipt) avec solides multiples
        /// </summary>
        IPT,

        /// <summary>
        /// Fichier STEP (.stp) - Format d'échange standard
        /// </summary>
        STEP
    }
}
