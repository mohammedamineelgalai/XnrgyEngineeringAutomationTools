using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models
{
    /// <summary>
    /// Type de fichier principal pour l'equipement
    /// </summary>
    public enum PrimaryFileType
    {
        /// <summary>Assemblage (.iam)</summary>
        Assembly,
        /// <summary>Piece (.ipt)</summary>
        Part,
        /// <summary>Dessin (.idw)</summary>
        Drawing
    }

    /// <summary>
    /// Variante d'un equipement (pour equipements avec plusieurs options comme Infinitum)
    /// </summary>
    public class EquipmentVariant
    {
        /// <summary>Nom de la variante (ex: "7.5Hp_1800RPM")</summary>
        public string Name { get; set; } = string.Empty;
        
        /// <summary>Nom d'affichage de la variante</summary>
        public string DisplayName { get; set; } = string.Empty;
        
        /// <summary>Nom du fichier IPJ de la variante</summary>
        public string ProjectFileName { get; set; } = string.Empty;
        
        /// <summary>Nom du fichier principal (.iam ou .ipt)</summary>
        public string PrimaryFileName { get; set; } = string.Empty;
        
        /// <summary>Type de fichier principal</summary>
        public PrimaryFileType FileType { get; set; } = PrimaryFileType.Assembly;
        
        /// <summary>Sous-chemin relatif dans le dossier equipement (ex: "7.5Hp_1800RPM")</summary>
        public string SubFolder { get; set; } = string.Empty;
    }

    /// <summary>
    /// Modele representant un equipement avec ses fichiers .ipj et .iam/.ipt principaux
    /// Supporte les equipements avec variantes et fichiers IPT comme principal
    /// </summary>
    public class EquipmentItem : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        /// <summary>
        /// Nom interne de l'equipement (ex: "Angular_Filter", "Infinitum")
        /// Utilise pour le nom de dossier Vault
        /// </summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        private string _displayName = string.Empty;
        /// <summary>
        /// Nom d'affichage de l'equipement (ex: "Angular Filter", "Infinitum Motor")
        /// </summary>
        public string DisplayName
        {
            get => _displayName;
            set { _displayName = value; OnPropertyChanged(); }
        }

        private string _projectFileName = string.Empty;
        /// <summary>
        /// Nom du fichier projet (.ipj) - ESSENTIEL pour Copy Design
        /// </summary>
        public string ProjectFileName
        {
            get => _projectFileName;
            set { _projectFileName = value; OnPropertyChanged(); }
        }

        private string _assemblyFileName = string.Empty;
        /// <summary>
        /// Nom du fichier principal (.iam ou .ipt) a inserer dans le top assembly
        /// DEPRECATED: Utiliser PrimaryFileName a la place
        /// </summary>
        public string AssemblyFileName
        {
            get => _assemblyFileName;
            set { _assemblyFileName = value; OnPropertyChanged(); }
        }

        private string _primaryFileName = string.Empty;
        /// <summary>
        /// Nom du fichier principal (.iam ou .ipt) a copier/placer
        /// </summary>
        public string PrimaryFileName
        {
            get => string.IsNullOrEmpty(_primaryFileName) ? _assemblyFileName : _primaryFileName;
            set { _primaryFileName = value; OnPropertyChanged(); }
        }

        private PrimaryFileType _primaryFileType = PrimaryFileType.Assembly;
        /// <summary>
        /// Type du fichier principal (Assembly=.iam, Part=.ipt)
        /// </summary>
        public PrimaryFileType PrimaryFileType
        {
            get => _primaryFileType;
            set { _primaryFileType = value; OnPropertyChanged(); }
        }

        private string _vaultPath = string.Empty;
        /// <summary>
        /// Chemin Vault de l'equipement (ex: $/Engineering/Library/Equipment/Angular_Filter)
        /// </summary>
        public string VaultPath
        {
            get => _vaultPath;
            set { _vaultPath = value; OnPropertyChanged(); }
        }

        private string _localTempPath = string.Empty;
        /// <summary>
        /// Chemin local temporaire apres telechargement (ex: C:\Vault\Engineering\Library\Equipment\Angular_Filter)
        /// </summary>
        public string LocalTempPath
        {
            get => _localTempPath;
            set { _localTempPath = value; OnPropertyChanged(); }
        }

        private List<EquipmentVariant>? _variants;
        /// <summary>
        /// Liste des variantes disponibles pour cet equipement (ex: Infinitum avec 7.5Hp, 10Hp, etc.)
        /// Si null ou vide, l'equipement n'a qu'une seule version
        /// </summary>
        public List<EquipmentVariant>? Variants
        {
            get => _variants;
            set { _variants = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasVariants)); }
        }

        /// <summary>
        /// Indique si cet equipement a plusieurs variantes
        /// </summary>
        public bool HasVariants => Variants != null && Variants.Count > 0;

        private List<string>? _alternateDrawings;
        /// <summary>
        /// Liste des dessins alternatifs disponibles (ex: Damper avec Floor_Damper.idw et Wall_and_Roof_Damper.idw)
        /// </summary>
        public List<string>? AlternateDrawings
        {
            get => _alternateDrawings;
            set { _alternateDrawings = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasAlternateDrawings)); }
        }

        /// <summary>
        /// Indique si cet equipement a plusieurs dessins alternatifs
        /// </summary>
        public bool HasAlternateDrawings => AlternateDrawings != null && AlternateDrawings.Count > 1;

        private EquipmentVariant? _selectedVariant;
        /// <summary>
        /// Variante selectionnee par l'utilisateur (si HasVariants)
        /// </summary>
        public EquipmentVariant? SelectedVariant
        {
            get => _selectedVariant;
            set { _selectedVariant = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Retourne le nom du fichier IPJ effectif (tenant compte de la variante selectionnee)
        /// </summary>
        public string EffectiveProjectFileName => SelectedVariant?.ProjectFileName ?? ProjectFileName;

        /// <summary>
        /// Retourne le nom du fichier principal effectif (tenant compte de la variante selectionnee)
        /// </summary>
        public string EffectivePrimaryFileName => SelectedVariant?.PrimaryFileName ?? PrimaryFileName;

        /// <summary>
        /// Retourne le type de fichier effectif (tenant compte de la variante selectionnee)
        /// </summary>
        public PrimaryFileType EffectiveFileType => SelectedVariant?.FileType ?? PrimaryFileType;

        /// <summary>
        /// Retourne le sous-dossier effectif pour la variante (vide si pas de variante)
        /// </summary>
        public string EffectiveSubFolder => SelectedVariant?.SubFolder ?? string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}


