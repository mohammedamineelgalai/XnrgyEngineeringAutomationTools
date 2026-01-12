using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using XnrgyEngineeringAutomationTools.Models;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models;
using XnrgyEngineeringAutomationTools.Shared.Views;
using WinForms = System.Windows.Forms;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Views
{
    /// <summary>
    /// Fenetre de reglages pour le module "Place Equipment" - Acces Admin uniquement
    /// Les parametres sont sauvegardes de maniere chiffree et synchronises via Vault
    /// vers $/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/
    /// </summary>
    public partial class PlaceEquipmentSettingsWindow : Window
    {
        private PlaceEquipmentSettings _settings;
        private readonly VaultSettingsService _vaultSettingsService;
        private bool _isDirty = false;

        public PlaceEquipmentSettingsWindow(VaultSdkService? vaultService = null)
        {
            InitializeComponent();
            
            // Initialiser le service de parametres avec Vault
            _vaultSettingsService = new VaultSettingsService(vaultService);
            
            LoadSettings();
        }

        /// <summary>
        /// Charge les parametres actuels dans les controles
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                // Charger depuis le service Vault (synchronise depuis Vault si connecte)
                _vaultSettingsService.Reload();
                _settings = _vaultSettingsService.Current.PlaceEquipment ?? new PlaceEquipmentSettings();

                // Initiales dessinateurs
                TxtDesignerInitials.Text = _settings.DesignerInitials != null 
                    ? string.Join(", ", _settings.DesignerInitials) 
                    : string.Empty;

                // Chemins
                TxtEquipmentLocalPath.Text = _settings.EquipmentLocalPath ?? string.Empty;
                TxtEquipmentVaultPath.Text = _settings.EquipmentVaultPath ?? string.Empty;
                TxtProjectsLocalPath.Text = _settings.ProjectsLocalBasePath ?? string.Empty;
                TxtProjectsVaultPath.Text = _settings.ProjectsVaultBasePath ?? string.Empty;
                TxtEquipmentSubfolder.Text = _settings.EquipmentSubfolderName ?? "1-Equipment";

                // Prefixes
                TxtReferenceFolderPrefix.Text = _settings.ReferenceFolderPrefix ?? "REF";
                TxtModuleFolderPrefix.Text = _settings.ModuleFolderPrefix ?? "M";

                // Options
                ChkAutoUploadToVault.IsChecked = _settings.AutoUploadToVault;
                ChkAutoOpenTopAssembly.IsChecked = _settings.AutoOpenTopAssembly;
                ChkAutoCreateEquipmentFolder.IsChecked = _settings.AutoCreateEquipmentFolder;

                // Extensions
                TxtInventorExtensions.Text = _settings.InventorExtensions != null
                    ? string.Join(", ", _settings.InventorExtensions)
                    : string.Empty;
                TxtExcludedExtensions.Text = _settings.ExcludedExtensions != null
                    ? string.Join(", ", _settings.ExcludedExtensions)
                    : string.Empty;

                // Dossiers exclus et library
                TxtExcludedFolders.Text = _settings.ExcludedFolders != null
                    ? string.Join(", ", _settings.ExcludedFolders)
                    : string.Empty;
                TxtLibraryPatterns.Text = _settings.LibraryPathPatterns != null
                    ? string.Join("\n", _settings.LibraryPathPatterns)
                    : string.Empty;

                // iProperties
                if (_settings.IPropertyNames != null)
                {
                    TxtIPropDesigner.Text = _settings.IPropertyNames.DesignerInitial ?? string.Empty;
                    TxtIPropCoDesigner.Text = _settings.IPropertyNames.CoDesignerInitial ?? string.Empty;
                    TxtIPropCreationDate.Text = _settings.IPropertyNames.CreationDate ?? string.Empty;
                    TxtIPropJobTitle.Text = _settings.IPropertyNames.JobTitle ?? string.Empty;
                    TxtIPropProject.Text = _settings.IPropertyNames.Project ?? string.Empty;
                    TxtIPropReference.Text = _settings.IPropertyNames.Reference ?? string.Empty;
                    TxtIPropModule.Text = _settings.IPropertyNames.Module ?? string.Empty;
                    TxtIPropEquipmentType.Text = _settings.IPropertyNames.EquipmentType ?? string.Empty;
                    TxtIPropEquipmentInstance.Text = _settings.IPropertyNames.EquipmentInstance ?? string.Empty;
                }

                _isDirty = false;
            }
            catch (Exception ex)
            {
                XnrgyMessageBox.ShowError(
                    $"Erreur lors du chargement des parametres:\n\n{ex.Message}",
                    "Erreur",
                    this);
            }
        }

        /// <summary>
        /// Sauvegarde les parametres depuis les controles
        /// </summary>
        private bool SaveSettings()
        {
            try
            {
                // Parser les initiales (virgule ou retour a la ligne)
                var initialsText = TxtDesignerInitials.Text;
                var initials = initialsText
                    .Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct()
                    .ToList();

                _settings.DesignerInitials = initials;

                // Chemins
                _settings.EquipmentLocalPath = TxtEquipmentLocalPath.Text.Trim();
                _settings.EquipmentVaultPath = TxtEquipmentVaultPath.Text.Trim();
                _settings.ProjectsLocalBasePath = TxtProjectsLocalPath.Text.Trim();
                _settings.ProjectsVaultBasePath = TxtProjectsVaultPath.Text.Trim();
                _settings.EquipmentSubfolderName = TxtEquipmentSubfolder.Text.Trim();

                // Prefixes
                _settings.ReferenceFolderPrefix = TxtReferenceFolderPrefix.Text.Trim();
                _settings.ModuleFolderPrefix = TxtModuleFolderPrefix.Text.Trim();

                // Options
                _settings.AutoUploadToVault = ChkAutoUploadToVault.IsChecked ?? true;
                _settings.AutoOpenTopAssembly = ChkAutoOpenTopAssembly.IsChecked ?? true;
                _settings.AutoCreateEquipmentFolder = ChkAutoCreateEquipmentFolder.IsChecked ?? true;

                // Extensions Inventor
                var inventorExt = TxtInventorExtensions.Text
                    .Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim().ToLowerInvariant())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct()
                    .ToList();
                _settings.InventorExtensions = inventorExt;

                // Extensions exclues
                var excludedExt = TxtExcludedExtensions.Text
                    .Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim().ToLowerInvariant())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct()
                    .ToList();
                _settings.ExcludedExtensions = excludedExt;

                // Dossiers exclus
                var excludedFolders = TxtExcludedFolders.Text
                    .Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct()
                    .ToList();
                _settings.ExcludedFolders = excludedFolders;

                // Library patterns
                var libraryPatterns = TxtLibraryPatterns.Text
                    .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct()
                    .ToList();
                _settings.LibraryPathPatterns = libraryPatterns;

                // iProperties
                _settings.IPropertyNames = new PlaceEquipmentIPropertySettings
                {
                    DesignerInitial = TxtIPropDesigner.Text.Trim(),
                    CoDesignerInitial = TxtIPropCoDesigner.Text.Trim(),
                    CreationDate = TxtIPropCreationDate.Text.Trim(),
                    JobTitle = TxtIPropJobTitle.Text.Trim(),
                    Project = TxtIPropProject.Text.Trim(),
                    Reference = TxtIPropReference.Text.Trim(),
                    Module = TxtIPropModule.Text.Trim(),
                    EquipmentType = TxtIPropEquipmentType.Text.Trim(),
                    EquipmentInstance = TxtIPropEquipmentInstance.Text.Trim()
                };

                // Sauvegarder via le service Vault (chiffre + upload vers Vault)
                var currentSettings = _vaultSettingsService.Current;
                currentSettings.PlaceEquipment = _settings;
                bool success = _vaultSettingsService.Save(currentSettings);

                if (success)
                {
                    _isDirty = false;
                    XnrgyMessageBox.ShowSuccess(
                        "Parametres sauvegardes et synchronises avec Vault.\n\nTous les utilisateurs recevront ces parametres au prochain lancement de l'application.",
                        "Parametres Sauvegardes",
                        this);
                    return true;
                }
                else
                {
                    XnrgyMessageBox.ShowError(
                        "Erreur lors de la sauvegarde des parametres.\n\nVerifiez vos droits administrateur Vault.",
                        "Erreur de Sauvegarde",
                        this);
                    return false;
                }
            }
            catch (Exception ex)
            {
                XnrgyMessageBox.ShowError(
                    $"Erreur lors de la sauvegarde:\n\n{ex.Message}",
                    "Erreur",
                    this);
                return false;
            }
        }

        #region Event Handlers

        private void BtnBrowseEquipment_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "Selectionner le dossier Equipements (Library)";
                dialog.SelectedPath = TxtEquipmentLocalPath.Text;
                dialog.ShowNewFolderButton = false;

                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    TxtEquipmentLocalPath.Text = dialog.SelectedPath;
                    _isDirty = true;
                }
            }
        }

        private void BtnBrowseProjects_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "Selectionner le dossier Projets";
                dialog.SelectedPath = TxtProjectsLocalPath.Text;
                dialog.ShowNewFolderButton = false;

                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    TxtProjectsLocalPath.Text = dialog.SelectedPath;
                    _isDirty = true;
                }
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveSettings())
            {
                // Message deja affiche dans SaveSettings()
                DialogResult = true;
                Close();
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (_isDirty)
            {
                bool confirm = XnrgyMessageBox.Confirm(
                    "Des modifications non sauvegardees seront perdues.\n\nVoulez-vous quitter sans sauvegarder?",
                    "Modifications Non Sauvegardees",
                    this);

                if (!confirm)
                    return;
            }

            DialogResult = false;
            Close();
        }

        private void BtnResetDefaults_Click(object sender, RoutedEventArgs e)
        {
            bool confirm = XnrgyMessageBox.Confirm(
                "Voulez-vous reinitialiser tous les parametres aux valeurs par defaut?\n\nCette action ne peut pas etre annulee.",
                "Reinitialisation des Parametres",
                this);

            if (confirm)
            {
                _settings = new PlaceEquipmentSettings();
                LoadSettingsToControls();
                _isDirty = true;
                XnrgyMessageBox.ShowInfo(
                    "Parametres reinitialises.\n\nCliquez sur 'Sauvegarder' pour appliquer les changements.",
                    "Parametres Reinitialises",
                    this);
            }
        }

        /// <summary>
        /// Charge les parametres dans les controles (utilise apres reinitialisation)
        /// </summary>
        private void LoadSettingsToControls()
        {
            // Initiales dessinateurs
            TxtDesignerInitials.Text = _settings.DesignerInitials != null 
                ? string.Join(", ", _settings.DesignerInitials) 
                : string.Empty;

            // Chemins
            TxtEquipmentLocalPath.Text = _settings.EquipmentLocalPath ?? string.Empty;
            TxtEquipmentVaultPath.Text = _settings.EquipmentVaultPath ?? string.Empty;
            TxtProjectsLocalPath.Text = _settings.ProjectsLocalBasePath ?? string.Empty;
            TxtProjectsVaultPath.Text = _settings.ProjectsVaultBasePath ?? string.Empty;
            TxtEquipmentSubfolder.Text = _settings.EquipmentSubfolderName ?? "1-Equipment";

            // Prefixes
            TxtReferenceFolderPrefix.Text = _settings.ReferenceFolderPrefix ?? "REF";
            TxtModuleFolderPrefix.Text = _settings.ModuleFolderPrefix ?? "M";

            // Options
            ChkAutoUploadToVault.IsChecked = _settings.AutoUploadToVault;
            ChkAutoOpenTopAssembly.IsChecked = _settings.AutoOpenTopAssembly;
            ChkAutoCreateEquipmentFolder.IsChecked = _settings.AutoCreateEquipmentFolder;

            // Extensions
            TxtInventorExtensions.Text = _settings.InventorExtensions != null
                ? string.Join(", ", _settings.InventorExtensions)
                : string.Empty;
            TxtExcludedExtensions.Text = _settings.ExcludedExtensions != null
                ? string.Join(", ", _settings.ExcludedExtensions)
                : string.Empty;

            // Dossiers exclus et library
            TxtExcludedFolders.Text = _settings.ExcludedFolders != null
                ? string.Join(", ", _settings.ExcludedFolders)
                : string.Empty;
            TxtLibraryPatterns.Text = _settings.LibraryPathPatterns != null
                ? string.Join("\n", _settings.LibraryPathPatterns)
                : string.Empty;

            // iProperties
            if (_settings.IPropertyNames != null)
            {
                TxtIPropDesigner.Text = _settings.IPropertyNames.DesignerInitial ?? string.Empty;
                TxtIPropCoDesigner.Text = _settings.IPropertyNames.CoDesignerInitial ?? string.Empty;
                TxtIPropCreationDate.Text = _settings.IPropertyNames.CreationDate ?? string.Empty;
                TxtIPropJobTitle.Text = _settings.IPropertyNames.JobTitle ?? string.Empty;
                TxtIPropProject.Text = _settings.IPropertyNames.Project ?? string.Empty;
                TxtIPropReference.Text = _settings.IPropertyNames.Reference ?? string.Empty;
                TxtIPropModule.Text = _settings.IPropertyNames.Module ?? string.Empty;
                TxtIPropEquipmentType.Text = _settings.IPropertyNames.EquipmentType ?? string.Empty;
                TxtIPropEquipmentInstance.Text = _settings.IPropertyNames.EquipmentInstance ?? string.Empty;
            }
        }

        #endregion
    }
}
