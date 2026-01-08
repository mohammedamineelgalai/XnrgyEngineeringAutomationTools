using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenêtre WPF moderne pour afficher et éditer les iProperties
    /// Remplace la version HTML/WebView2 pour plus de fiabilité
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public partial class IPropertiesWindow : Window
    {
        private ObservableCollection<PropertyItem> _properties;
        private ObservableCollection<PropertyGroup> _propertyGroups;
        private Dictionary<string, string> _originalValues;
        private Func<string, string, bool>? _applyPropertyChange;
        private Func<string, string, bool>? _addNewProperty;

        // Win32 API pour centrer sur Inventor
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        
        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
        
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        public IPropertiesWindow()
        {
            InitializeComponent();
            _properties = new ObservableCollection<PropertyItem>();
            _propertyGroups = new ObservableCollection<PropertyGroup>();
            _originalValues = new Dictionary<string, string>();
            PropertiesGroupList.ItemsSource = _propertyGroups;
            Loaded += IPropertiesWindow_Loaded;
        }

        /// <summary>
        /// Constructeur avec les données de propriétés
        /// </summary>
        public IPropertiesWindow(string fileName, string fullPath, Dictionary<string, string> properties,
            Func<string, string, bool>? applyPropertyChange = null,
            Func<string, string, bool>? addNewProperty = null) : this()
        {
            TxtFileName.Text = fileName;
            TxtFilePath.Text = fullPath;
            Title = $"iProperties - {fileName}";
            
            _applyPropertyChange = applyPropertyChange;
            _addNewProperty = addNewProperty;

            // Icônes et couleurs par propriété (groupes comme Inventor)
            var iconData = new Dictionary<string, (string symbol, string color, string category)>
            {
                // === CUSTOM (Propriétés personnalisées XNRGY) - Rouge ===
                {"Prefix", ("Aa", "#e74c3c", "Custom")},
                {"Tag", ("◆", "#e74c3c", "Custom")},
                {"Tag_Assy", ("◈", "#e74c3c", "Custom")},
                {"Structural Style", ("◎", "#e74c3c", "Custom")},
                {"DesignViewRepresentation", ("▢", "#e74c3c", "Custom")},
                {"Destination", ("➤", "#e74c3c", "Custom")},
                {"Length", ("⊥", "#e74c3c", "Custom")},
                {"Width", ("↔", "#e74c3c", "Custom")},
                {"Depth", ("↕", "#e74c3c", "Custom")},
                {"Thickness", ("═", "#e74c3c", "Custom")},
                {"Engraving", ("✒", "#e74c3c", "Custom")},
                {"Finish_Paint_Face", ("◐", "#e74c3c", "Custom")},
                {"MachineNo", ("⚙", "#e74c3c", "Custom")},
                {"FlatPatternLength", ("⊥", "#e74c3c", "Custom")},
                {"FlatPatternWidth", ("↔", "#e74c3c", "Custom")},
                {"Liner Bend Side", ("⟳", "#e74c3c", "Custom")},
                
                // === PROJECT (Design Tracking) - Bleu ===
                {"Part Number", ("＃", "#3498db", "Design")},
                {"Stock Number", ("▤", "#3498db", "Design")},
                {"Revision Number", ("◉", "#3498db", "Design")},
                {"Project", ("▧", "#3498db", "Design")},
                {"Designer", ("●", "#3498db", "Design")},
                {"Engineer", ("◈", "#3498db", "Design")},
                {"Cost Center", ("$", "#3498db", "Design")},
                {"Cost", ("€", "#3498db", "Design")},
                {"Vendor", ("▨", "#3498db", "Design")},
                {"Description", ("✎", "#3498db", "Design")},
                
                // === STATUS - Orange ===
                {"Design Status", ("◉", "#f39c12", "Status")},
                {"Checked By", ("✓", "#f39c12", "Status")},
                {"Date Checked", ("▪", "#f39c12", "Status")},
                {"Engr Approved By", ("✓", "#f39c12", "Status")},
                {"Engr Date Approved", ("▪", "#f39c12", "Status")},
                
                // === SUMMARY - Vert ===
                {"Auteur", ("●", "#27ae60", "Summary")},
                {"Title", ("T", "#27ae60", "Summary")},
                {"Subject", ("S", "#27ae60", "Summary")},
                {"Keywords", ("K", "#27ae60", "Summary")},
                {"Comments", ("✎", "#27ae60", "Summary")},
                {"Entreprise", ("▦", "#27ae60", "Summary")},
                {"Manager", ("M", "#27ae60", "Summary")},
                {"Category", ("C", "#27ae60", "Summary")},
                
                // === PHYSICAL - Violet ===
                {"Material", ("▣", "#9b59b6", "Physical")},
                {"Mass", ("M", "#9b59b6", "Physical")},
                {"Area", ("A", "#9b59b6", "Physical")},
                {"Volume", ("V", "#9b59b6", "Physical")},
                {"Density", ("D", "#9b59b6", "Physical")},
                
                // === SAVE - Gris ===
                {"Derniere sauvegarde", ("▪", "#7f8c8d", "System")},
                {"Creation Date", ("▪", "#7f8c8d", "System")},
                {"Last Modified", ("▪", "#7f8c8d", "System")}
            };

            // Créer les groupes (ordre Inventor : Custom, General, Summary, Project, Status, Save, Physical)
            var customGroup = new PropertyGroup { Name = "Custom", Icon = "★", Color = "#f39c12" };    // Orange/doré
            var generalGroup = new PropertyGroup { Name = "General", Icon = "◉", Color = "#34495e" };  // Gris foncé
            var summaryGroup = new PropertyGroup { Name = "Summary", Icon = "S", Color = "#27ae60" };  // Vert
            var projectGroup = new PropertyGroup { Name = "Project", Icon = "◆", Color = "#3498db" };  // Bleu
            var statusGroup = new PropertyGroup { Name = "Status", Icon = "◎", Color = "#9b59b6" };    // Violet
            var saveGroup = new PropertyGroup { Name = "Save", Icon = "▪", Color = "#7f8c8d" };        // Gris
            var physicalGroup = new PropertyGroup { Name = "Physical", Icon = "▣", Color = "#1abc9c" }; // Cyan

            // Fonction helper pour obtenir la couleur par catégorie
            string GetColorForCategory(string cat) => cat switch
            {
                "Custom" => "#f39c12",
                "General" => "#34495e",
                "Summary" => "#27ae60",
                "Project" => "#3498db",
                "Status" => "#9b59b6",
                "Save" => "#7f8c8d",
                "Physical" => "#1abc9c",
                _ => "#667eea"
            };

            foreach (var prop in properties)
            {
                // Extraire la catégorie du préfixe (ex: "Custom:Tag" -> category="Custom", name="Tag")
                string category = "Custom";
                string propName = prop.Key;
                
                if (prop.Key.Contains(":"))
                {
                    var parts = prop.Key.Split(new[] { ':' }, 2);
                    category = parts[0];
                    propName = parts[1];
                }
                
                // Obtenir l'icône et la couleur
                var (symbol, color, _) = iconData.ContainsKey(propName) 
                    ? iconData[propName] 
                    : ("◇", GetColorForCategory(category), category);
                    
                var item = new PropertyItem 
                { 
                    Name = propName, 
                    Value = prop.Value ?? "", 
                    Icon = symbol,
                    IconColor = color,
                    Category = category,
                    OriginalValue = prop.Value ?? ""
                };
                item.PropertyChanged += PropertyItem_PropertyChanged;
                _properties.Add(item);
                _originalValues[prop.Key] = prop.Value ?? "";

                // Ajouter au groupe approprié
                switch (category)
                {
                    case "Custom": customGroup.Properties.Add(item); break;
                    case "General": generalGroup.Properties.Add(item); break;
                    case "Summary": summaryGroup.Properties.Add(item); break;
                    case "Project": projectGroup.Properties.Add(item); break;
                    case "Status": statusGroup.Properties.Add(item); break;
                    case "Save": saveGroup.Properties.Add(item); break;
                    case "Physical": physicalGroup.Properties.Add(item); break;
                    default: customGroup.Properties.Add(item); break;
                }
            }

            // Ajouter les groupes dans l'ordre demandé: Custom, General, Summary, Project, Status, Save, Physical
            if (customGroup.Properties.Count > 0) _propertyGroups.Add(customGroup);
            if (generalGroup.Properties.Count > 0) _propertyGroups.Add(generalGroup);
            if (summaryGroup.Properties.Count > 0) _propertyGroups.Add(summaryGroup);
            if (projectGroup.Properties.Count > 0) _propertyGroups.Add(projectGroup);
            if (statusGroup.Properties.Count > 0) _propertyGroups.Add(statusGroup);
            if (saveGroup.Properties.Count > 0) _propertyGroups.Add(saveGroup);
            if (physicalGroup.Properties.Count > 0) _propertyGroups.Add(physicalGroup);
        }

        private void IPropertiesWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Ajuster la hauteur pour ne pas dépasser l'écran (avec marge pour la barre des tâches)
            double screenHeight = SystemParameters.WorkArea.Height; // Hauteur disponible (sans barre des tâches)
            double screenWidth = SystemParameters.WorkArea.Width;
            
            // Prendre 90% de la hauteur disponible, max 1000px
            double targetHeight = Math.Min(screenHeight * 0.92, 1100);
            this.Height = Math.Max(targetHeight, 700); // Minimum 700px
            
            // Centrer sur Inventor
            CenterOnInventorWindow();
            
            // S'assurer que la fenêtre reste dans l'écran
            if (this.Top + this.Height > screenHeight)
            {
                this.Top = Math.Max(0, screenHeight - this.Height - 10);
            }
            if (this.Left + this.Width > screenWidth)
            {
                this.Left = Math.Max(0, screenWidth - this.Width - 10);
            }
        }

        /// <summary>
        /// Centre la fenêtre sur Inventor
        /// </summary>
        private void CenterOnInventorWindow()
        {
            try
            {
                IntPtr inventorHandle = IntPtr.Zero;
                RECT inventorRect = new RECT();
                
                EnumWindows((hWnd, lParam) =>
                {
                    if (!IsWindowVisible(hWnd)) return true;
                    
                    int length = GetWindowTextLength(hWnd);
                    if (length == 0) return true;
                    
                    var sb = new System.Text.StringBuilder(length + 1);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    string title = sb.ToString();
                    
                    if (title.Contains("Autodesk Inventor") || 
                        title.EndsWith(".iam") || title.EndsWith(".ipt") || 
                        title.EndsWith(".idw") || title.EndsWith(".ipn") ||
                        title.Contains("Inventor Professional") ||
                        title.Contains("Inventor 20"))
                    {
                        if (GetWindowRect(hWnd, out RECT rect))
                        {
                            int width = rect.Right - rect.Left;
                            int height = rect.Bottom - rect.Top;
                            if (width > 800 && height > 600)
                            {
                                inventorHandle = hWnd;
                                inventorRect = rect;
                                return false;
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);

                if (inventorHandle != IntPtr.Zero)
                {
                    int inventorWidth = inventorRect.Right - inventorRect.Left;
                    int inventorHeight = inventorRect.Bottom - inventorRect.Top;
                    int inventorCenterX = inventorRect.Left + (inventorWidth / 2);
                    int inventorCenterY = inventorRect.Top + (inventorHeight / 2);

                    this.Left = inventorCenterX - (this.Width / 2);
                    this.Top = inventorCenterY - (this.Height / 2);

                    // S'assurer que la fenêtre reste visible
                    var screenWidth = SystemParameters.PrimaryScreenWidth;
                    var screenHeight = SystemParameters.PrimaryScreenHeight;
                    
                    if (this.Left < 0) this.Left = 20;
                    if (this.Top < 0) this.Top = 20;
                    if (this.Left + this.Width > screenWidth) this.Left = screenWidth - this.Width - 20;
                    if (this.Top + this.Height > screenHeight) this.Top = screenHeight - this.Height - 50;
                }
                else
                {
                    this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                    this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;
                }
            }
            catch
            {
                this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;
            }
        }

        /// <summary>
        /// Détecte les modifications de propriétés
        /// </summary>
        private void PropertyItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Value")
            {
                UpdateApplyButtonState();
            }
        }

        /// <summary>
        /// Met à jour l'état du bouton Appliquer
        /// </summary>
        private void UpdateApplyButtonState()
        {
            bool hasChanges = _properties.Any(p => p.Value != p.OriginalValue);
            
            if (hasChanges)
            {
                BtnApply.Background = new LinearGradientBrush(
                    Color.FromRgb(39, 174, 96),
                    Color.FromRgb(46, 204, 113),
                    45);
            }
            else
            {
                BtnApply.Background = new LinearGradientBrush(
                    Color.FromRgb(52, 152, 219),
                    Color.FromRgb(41, 128, 185),
                    45);
            }
        }

        /// <summary>
        /// Affiche un message de statut
        /// </summary>
        private void ShowStatus(string message, string type)
        {
            TxtStatus.Text = message;
            StatusBorder.Visibility = Visibility.Visible;
            
            switch (type)
            {
                case "success":
                    StatusBorder.Background = new SolidColorBrush(Color.FromRgb(212, 237, 218));
                    TxtStatus.Foreground = new SolidColorBrush(Color.FromRgb(21, 87, 36));
                    break;
                case "error":
                    StatusBorder.Background = new SolidColorBrush(Color.FromRgb(248, 215, 218));
                    TxtStatus.Foreground = new SolidColorBrush(Color.FromRgb(114, 28, 36));
                    break;
                case "info":
                    StatusBorder.Background = new SolidColorBrush(Color.FromRgb(209, 236, 241));
                    TxtStatus.Foreground = new SolidColorBrush(Color.FromRgb(12, 84, 96));
                    break;
            }

            // Masquer après 4 secondes pour success/error
            if (type != "info")
            {
                var timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromSeconds(4);
                timer.Tick += (s, e) =>
                {
                    StatusBorder.Visibility = Visibility.Collapsed;
                    timer.Stop();
                };
                timer.Start();
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
            }
            else
            {
                DragMove();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            var modifiedProperties = _properties.Where(p => p.Value != p.OriginalValue).ToList();
            
            if (modifiedProperties.Count == 0)
            {
                ShowStatus("ℹ️ Aucune modification à appliquer", "info");
                return;
            }

            ShowStatus($"⏳ Application de {modifiedProperties.Count} modification(s)...", "info");
            
            int successCount = 0;
            int errorCount = 0;

            foreach (var prop in modifiedProperties)
            {
                bool success = true;
                if (_applyPropertyChange != null)
                {
                    success = _applyPropertyChange(prop.Name, prop.Value);
                }

                if (success)
                {
                    prop.OriginalValue = prop.Value;
                    successCount++;
                }
                else
                {
                    errorCount++;
                }
            }

            if (errorCount == 0)
            {
                ShowStatus($"✅ {successCount} propriété(s) mise(s) à jour avec succès!", "success");
            }
            else
            {
                ShowStatus($"⚠️ {successCount} réussie(s), {errorCount} erreur(s)", "error");
            }

            UpdateApplyButtonState();
        }

        private void TxtNewPropertyName_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            BtnAddProperty.IsEnabled = !string.IsNullOrWhiteSpace(TxtNewPropertyName.Text);
        }

        private void BtnAddProperty_Click(object sender, RoutedEventArgs e)
        {
            string name = TxtNewPropertyName.Text.Trim();
            string value = TxtNewPropertyValue.Text;

            if (string.IsNullOrWhiteSpace(name))
            {
                ShowStatus("❌ Le nom de la propriété est requis!", "error");
                return;
            }

            if (_properties.Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                ShowStatus("❌ Cette propriété existe déjà!", "error");
                return;
            }

            ShowStatus("⏳ Ajout de la propriété...", "info");

            bool success = true;
            if (_addNewProperty != null)
            {
                success = _addNewProperty(name, value);
            }

            if (success)
            {
                var newItem = new PropertyItem
                {
                    Name = name,
                    Value = value,
                    Icon = "★",
                    IconColor = "#e74c3c", // Rouge pour nouvelle propriété
                    OriginalValue = value
                };
                newItem.PropertyChanged += PropertyItem_PropertyChanged;
                _properties.Add(newItem);
                _originalValues[name] = value;

                TxtNewPropertyName.Text = "";
                TxtNewPropertyValue.Text = "";
                BtnAddProperty.IsEnabled = false;

                ShowStatus($"✅ Propriété '{name}' ajoutée avec succès!", "success");
            }
            else
            {
                ShowStatus("❌ Erreur lors de l'ajout de la propriété", "error");
            }
        }

        /// <summary>
        /// Affiche la fenêtre iProperties
        /// </summary>
        public static void ShowProperties(string fileName, string fullPath, 
            Dictionary<string, string> properties,
            Func<string, string, bool>? applyPropertyChange = null,
            Func<string, string, bool>? addNewProperty = null)
        {
            var window = new IPropertiesWindow(fileName, fullPath, properties, applyPropertyChange, addNewProperty);
            window.ShowDialog();
        }
    }

    /// <summary>
    /// Modèle de données pour une propriété
    /// </summary>
    public class PropertyItem : INotifyPropertyChanged
    {
        private string _name = "";
        private string _value = "";
        private string _icon = "◇";
        private string _iconColor = "#667eea";
        private string _originalValue = "";
        private string _category = "Custom";

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(nameof(Value)); }
        }

        public string Icon
        {
            get => _icon;
            set { _icon = value; OnPropertyChanged(nameof(Icon)); }
        }

        public string IconColor
        {
            get => _iconColor;
            set { _iconColor = value; OnPropertyChanged(nameof(IconColor)); OnPropertyChanged(nameof(IconBrush)); }
        }

        /// <summary>
        /// Catégorie de la propriété (Custom, Design, Summary, Document)
        /// </summary>
        public string Category
        {
            get => _category;
            set { _category = value; OnPropertyChanged(nameof(Category)); }
        }

        /// <summary>
        /// Brush pour la couleur de l'icône (pour le binding XAML)
        /// </summary>
        public SolidColorBrush IconBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(_iconColor));

        public string OriginalValue
        {
            get => _originalValue;
            set { _originalValue = value; OnPropertyChanged(nameof(OriginalValue)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Groupe de propriétés pour l'affichage
    /// </summary>
    public class PropertyGroup
    {
        public string Name { get; set; } = "";
        public string Icon { get; set; } = "◆";
        public string Color { get; set; } = "#667eea";
        public SolidColorBrush ColorBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(Color));
        public ObservableCollection<PropertyItem> Properties { get; set; } = new ObservableCollection<PropertyItem>();
    }
}

