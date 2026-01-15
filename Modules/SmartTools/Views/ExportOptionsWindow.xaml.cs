using System;
using System.IO;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using XnrgyEngineeringAutomationTools.Shared.Views;

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
        /// Activer la représentation par défaut
        /// </summary>
        public bool ActivateDefaultRepresentation => ChkActivateDefaultRep.IsChecked == true;

        /// <summary>
        /// Afficher tous les composants masqués (hors références)
        /// </summary>
        public bool ShowHiddenComponents => ChkShowHiddenComponents.IsChecked == true;

        /// <summary>
        /// Réduire l'arborescence du navigateur
        /// </summary>
        public bool CollapseBrowserTree => ChkCollapseBrowserTree.IsChecked == true;

        /// <summary>
        /// Appliquer la vue isométrique
        /// </summary>
        public bool ApplyIsometricView => ChkApplyIsometricView.IsChecked == true;

        /// <summary>
        /// Masquer les éléments de référence avant export
        /// </summary>
        public bool HideReferences => ChkHideReferences.IsChecked == true;

        /// <summary>
        /// Ouvrir le fichier après export
        /// </summary>
        public bool OpenAfterExport => ChkOpenAfterExport.IsChecked == true;

        /// <summary>
        /// Chemin complet du fichier de sortie
        /// Si destination locale: chemin direct
        /// Si destination Vault: fichier temporaire local
        /// </summary>
        public string FullOutputPath
        {
            get
            {
                string ext = SelectedFormat == ExportFormat.IPT ? ".ipt" : ".stp";
                
                if (IsDestinationVault)
                {
                    // Fichier temporaire local pour upload vers Vault
                    string tempDir = Path.Combine(Path.GetTempPath(), "XnrgyExport");
                    Directory.CreateDirectory(tempDir);
                    return Path.Combine(tempDir, OutputFileName + ext);
                }
                else
                {
                    // Chemin direct local
                    return Path.Combine(LocalDestinationPath, OutputFileName + ext);
                }
            }
        }

        public ExportOptionsWindow()
        {
            InitializeComponent();
            Loaded += ExportOptionsWindow_Loaded;
        }

        private void ExportOptionsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            CenterWindowOnInventor();
        }

        /// <summary>
        /// Centre la fenêtre sur la fenêtre principale d'Inventor
        /// </summary>
        private void CenterWindowOnInventor()
        {
            try
            {
                var inventorProcesses = System.Diagnostics.Process.GetProcessesByName("Inventor");
                if (inventorProcesses.Length > 0)
                {
                    foreach (var proc in inventorProcesses)
                    {
                        try
                        {
                            if (proc.MainWindowHandle != IntPtr.Zero)
                            {
                                // Obtenir la position et taille de la fenêtre Inventor
                                RECT rect;
                                if (GetWindowRect(proc.MainWindowHandle, out rect))
                                {
                                    int inventorWidth = rect.Right - rect.Left;
                                    int inventorHeight = rect.Bottom - rect.Top;
                                    int inventorLeft = rect.Left;
                                    int inventorTop = rect.Top;

                                    // Calculer la position pour centrer
                                    this.Left = inventorLeft + (inventorWidth - (int)this.Width) / 2;
                                    this.Top = inventorTop + (inventorHeight - (int)this.Height) / 2;
                                    return;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            // Fallback: centrer sur l'écran
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// Project Number extrait depuis le chemin
        /// </summary>
        public string ProjectNumber { get; private set; } = "";

        /// <summary>
        /// Reference extraite depuis le chemin
        /// </summary>
        public string Reference { get; private set; } = "";

        /// <summary>
        /// Module extrait depuis le chemin (format M03)
        /// </summary>
        public string Module { get; private set; } = "";

        /// <summary>
        /// Chemin Vault de destination (si destination Vault)
        /// </summary>
        public string VaultDestinationPath { get; private set; } = "";

        /// <summary>
        /// Chemin local de destination (si destination locale)
        /// </summary>
        public string LocalDestinationPath { get; private set; } = "";

        /// <summary>
        /// Indique si la destination est Vault (true) ou locale (false)
        /// </summary>
        public bool IsDestinationVault { get; private set; } = false;

        /// <summary>
        /// Indique si c'est le Top Assembly (détecté par le nom du fichier)
        /// </summary>
        public bool IsTopAssembly { get; private set; } = true;

        /// <summary>
        /// Initialise la fenêtre avec les valeurs par défaut basées sur le document source
        /// </summary>
        /// <param name="sourceFileName">Nom du fichier source (ex: 123450101.iam ou LEFT_WALL_01.iam)</param>
        /// <param name="sourcePath">Chemin complet du fichier source</param>
        public void Initialize(string sourceFileName, string sourcePath)
        {
            TxtSourceFile.Text = sourceFileName;

            // Extraire Project Number, Reference et Module depuis le chemin
            // Format: C:\Vault\Engineering\Projects\12345\REF01\M03\...
            ExtractProjectAndReference(sourcePath);

            // Générer les chemins de destination (local et Vault)
            // Format local: C:\Vault\Engineering\Projects\12345\REF01\M03
            // Format Vault: $/Engineering/Projects/12345/REF01/M03
            if (!string.IsNullOrEmpty(ProjectNumber) && !string.IsNullOrEmpty(Reference))
            {
                if (!string.IsNullOrEmpty(Module))
                {
                    LocalDestinationPath = $"C:\\Vault\\Engineering\\Projects\\{ProjectNumber}\\{Reference}\\{Module}";
                    VaultDestinationPath = $"$/Engineering/Projects/{ProjectNumber}/{Reference}/{Module}";
                }
                else
                {
                    LocalDestinationPath = $"C:\\Vault\\Engineering\\Projects\\{ProjectNumber}\\{Reference}";
                    VaultDestinationPath = $"$/Engineering/Projects/{ProjectNumber}/{Reference}";
                }
            }
            else
            {
                LocalDestinationPath = "C:\\Vault\\Engineering\\Projects";
                VaultDestinationPath = "$/Engineering/Projects";
            }

            // Par défaut: destination locale (comme avant)
            // Définir IsChecked sans déclencher l'événement Checked
            RbDestinationLocal.Checked -= RbDestination_Checked;
            RbDestinationLocal.IsChecked = true;
            RbDestinationLocal.Checked += RbDestination_Checked;
            
            TxtDestinationPath.Text = LocalDestinationPath;
            TxtDestinationLabel.Text = "Chemin de destination (local):";
            BtnBrowseLocal.Visibility = Visibility.Visible;
            BtnBrowseVault.Visibility = Visibility.Collapsed;
            IsDestinationVault = false;

            // Générer un nom de fichier suggéré selon le nouveau format
            string suggestedName = GenerateSuggestedFileName(sourcePath, sourceFileName);
            TxtOutputFileName.Text = suggestedName;
            
            // Écouter les changements pour mettre à jour le chemin complet
            TxtOutputFileName.TextChanged += (s, e) => UpdateFullDestinationPath();
            RbDestinationLocal.Checked += (s, e) => UpdateFullDestinationPath();
            RbDestinationVault.Checked += (s, e) => UpdateFullDestinationPath();
            TxtDestinationPath.TextChanged += (s, e) => UpdateFullDestinationPath();
            
            // Mettre à jour le chemin complet de destination après l'initialisation complète
            // Utiliser Dispatcher pour s'assurer que tous les contrôles sont chargés
            Dispatcher.BeginInvoke(new Action(() => UpdateFullDestinationPath()), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        /// <summary>
        /// Met à jour l'affichage du chemin complet de destination
        /// </summary>
        private void UpdateFullDestinationPath()
        {
            try
            {
                // Vérifier que tous les contrôles sont initialisés
                if (RbExportIPT == null || RbExportSTEP == null || TxtOutputFileName == null || 
                    TxtDestinationPath == null || TxtFullDestinationPath == null ||
                    RbDestinationLocal == null || RbDestinationVault == null)
                {
                    return; // Contrôles pas encore initialisés
                }

                string ext = RbExportIPT.IsChecked == true ? ".ipt" : ".stp";
                string fileName = string.IsNullOrWhiteSpace(TxtOutputFileName.Text) ? "nom_fichier" : TxtOutputFileName.Text;
                string fullFileName = fileName + ext;
                
                if (RbDestinationVault.IsChecked == true)
                {
                    string vaultPath = string.IsNullOrWhiteSpace(TxtDestinationPath.Text) ? VaultDestinationPath : TxtDestinationPath.Text;
                    TxtFullDestinationPath.Text = $"{vaultPath}/{fullFileName}";
                }
                else
                {
                    string localPath = string.IsNullOrWhiteSpace(TxtDestinationPath.Text) ? LocalDestinationPath : TxtDestinationPath.Text;
                    TxtFullDestinationPath.Text = System.IO.Path.Combine(localPath, fullFileName);
                }
            }
            catch (Exception ex)
            {
                // Log l'erreur si possible, sinon ignorer
                try
                {
                    if (TxtFullDestinationPath != null)
                        TxtFullDestinationPath.Text = "--";
                }
                catch { }
            }
        }

        /// <summary>
        /// Extrait Project Number, Reference et Module depuis le chemin source
        /// Format: C:\Vault\Engineering\Projects\12345\REF01\M03\...
        /// </summary>
        private void ExtractProjectAndReference(string sourcePath)
        {
            try
            {
                string[] parts = sourcePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                
                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i].Equals("Projects", StringComparison.OrdinalIgnoreCase))
                    {
                        // Extraire Project Number (après "Projects")
                        if (i + 1 < parts.Length)
                        {
                            ProjectNumber = parts[i + 1];
                        }
                        
                        // Extraire Reference (format REF01 ou REF1)
                        if (i + 2 < parts.Length)
                        {
                            string refPart = parts[i + 2];
                            if (refPart.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
                            {
                                Reference = refPart;
                            }
                        }
                        
                        // Extraire Module (format M03 ou M3)
                        if (i + 3 < parts.Length)
                        {
                            string modulePart = parts[i + 3];
                            if (modulePart.StartsWith("M", StringComparison.OrdinalIgnoreCase) && 
                                System.Text.RegularExpressions.Regex.IsMatch(modulePart, @"^M\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                            {
                                Module = modulePart;
                            }
                        }
                        
                        break; // Sortir de la boucle une fois trouvé
                    }
                }
            }
            catch
            {
                // Valeurs par défaut vides
            }
        }

        /// <summary>
        /// Génère un nom de fichier suggéré selon le nouveau format
        /// Format Top Assembly: 123450101 (sans REF et M)
        /// Format Sous-Assembly: 123450101_LEFT_WALL_01 (avec préfixe 000000000_ si nécessaire)
        /// </summary>
        private string GenerateSuggestedFileName(string sourcePath, string sourceFileName)
        {
            try
            {
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourceFileName);
                
                // Si Project Number et Reference sont disponibles, construire le format Top Assembly
                if (!string.IsNullOrEmpty(ProjectNumber) && !string.IsNullOrEmpty(Reference))
                {
                    // Extraire les chiffres de Reference (REF01 -> 01)
                    string refNumber = Reference.Replace("REF", "").Replace("ref", "").Trim();
                    if (refNumber.Length == 1) refNumber = "0" + refNumber; // Ajouter 0 si un seul chiffre
                    
                    // Format Top Assembly: ProjectNumber + Reference (ex: 1234501)
                    string topAssemblyName = $"{ProjectNumber}{refNumber}01";
                    
                    // Vérifier si le nom du fichier correspond au format Top Assembly
                    // Si le nom correspond (ex: 123450101.iam), c'est le Top Assembly
                    if (fileNameWithoutExt == topAssemblyName || 
                        fileNameWithoutExt.StartsWith(topAssemblyName + "_", StringComparison.OrdinalIgnoreCase))
                    {
                        IsTopAssembly = true;
                        return topAssemblyName;
                    }
                    else
                    {
                        // C'est un sous-assembly, utiliser le nom du fichier avec préfixe si nécessaire
                        IsTopAssembly = false;
                        
                        // Si le nom ne commence pas par le format projet, ajouter préfixe 000000000_
                        if (!fileNameWithoutExt.StartsWith(ProjectNumber, StringComparison.OrdinalIgnoreCase))
                        {
                            return $"000000000_{fileNameWithoutExt}";
                        }
                        
                        // Sinon, utiliser le nom tel quel (ex: 123450101_LEFT_WALL_01)
                        return fileNameWithoutExt;
                    }
                }

                // Fallback: utiliser le nom du fichier source sans extension
                return fileNameWithoutExt;
            }
            catch
            {
                return Path.GetFileNameWithoutExtension(sourceFileName);
            }
        }

        private void RbDestination_Checked(object sender, RoutedEventArgs e)
        {
            // Mettre à jour le chemin de destination selon le choix
            if (RbDestinationLocal.IsChecked == true)
            {
                TxtDestinationPath.Text = LocalDestinationPath;
                TxtDestinationLabel.Text = "Chemin de destination (local):";
                BtnBrowseLocal.Visibility = Visibility.Visible;
                BtnBrowseVault.Visibility = Visibility.Collapsed;
                IsDestinationVault = false;
            }
            else if (RbDestinationVault.IsChecked == true)
            {
                TxtDestinationPath.Text = VaultDestinationPath;
                TxtDestinationLabel.Text = "Chemin de destination (Vault):";
                BtnBrowseLocal.Visibility = Visibility.Collapsed;
                BtnBrowseVault.Visibility = Visibility.Visible;
                IsDestinationVault = true;
            }
            
            // Mettre à jour le chemin complet
            UpdateFullDestinationPath();
        }

        private void RbExportFormat_Checked(object sender, RoutedEventArgs e)
        {
            // Mettre à jour le chemin complet quand le format change
            UpdateFullDestinationPath();
        }

        private void BtnBrowseLocal_Click(object sender, RoutedEventArgs e)
        {
            // Utiliser la fenêtre personnalisée pour sélectionner un dossier
            string selectedPath = FolderBrowserWindow.ShowDialog(this, TxtDestinationPath.Text);
            
            if (!string.IsNullOrEmpty(selectedPath) && Directory.Exists(selectedPath))
            {
                TxtDestinationPath.Text = selectedPath;
                LocalDestinationPath = selectedPath;
                UpdateFullDestinationPath();
            }
        }

        private void BtnBrowseVault_Click(object sender, RoutedEventArgs e)
        {
            // Afficher info pour Vault
            Shared.Views.XnrgyMessageBox.ShowInfo(
                "Destination Vault.\n\n" +
                $"Chemin Vault: {VaultDestinationPath}\n\n" +
                "Vous pouvez modifier le chemin Vault manuellement dans le champ ci-dessus.",
                "Destination Vault", this);
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
                    "Veuillez spécifier un chemin de destination.",
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

            // Validation selon le type de destination
            if (RbDestinationVault.IsChecked == true)
            {
                // Validation chemin Vault
                if (!TxtDestinationPath.Text.TrimStart().StartsWith("$/", StringComparison.OrdinalIgnoreCase))
                {
                    Shared.Views.XnrgyMessageBox.ShowError(
                        "Le chemin de destination Vault doit commencer par $/ (format Vault).\n\n" +
                        "Exemple: $/Engineering/Projects/12345/REF01",
                        "Chemin invalide", this);
                    return;
                }
                VaultDestinationPath = TxtDestinationPath.Text.Trim();
                IsDestinationVault = true;
            }
            else
            {
                // Validation chemin local
                if (!System.IO.Path.IsPathRooted(TxtDestinationPath.Text))
                {
                    Shared.Views.XnrgyMessageBox.ShowError(
                        "Le chemin de destination local doit être un chemin absolu.\n\n" +
                        "Exemple: C:\\Vault\\Engineering\\Projects\\12345\\REF01",
                        "Chemin invalide", this);
                    return;
                }
                LocalDestinationPath = TxtDestinationPath.Text.Trim();
                IsDestinationVault = false;
            }

            // DestinationPath est calculé depuis TxtDestinationPath.Text (pas besoin de l'assigner)

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
            try
            {
                var window = new ExportOptionsWindow();
                
                if (owner != null)
                {
                    window.Owner = owner;
                }

                // Initialiser après avoir défini le owner
                window.Initialize(sourceFileName, sourcePath);

                if (window.ShowDialog() == true && window.IsConfirmed)
                {
                    return window;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur ShowOptions: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                XnrgyMessageBox.ShowError($"Erreur lors de l'ouverture de la fenetre d'export:\n{ex.Message}", "Erreur");
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
