using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenêtre de gestion batch des propriétés personnalisées
    /// Permet d'éditer, ajouter et gérer les propriétés personnalisées d'un document Inventor
    /// </summary>
    public partial class CustomPropertyBatchWindow : Window, INotifyPropertyChanged
    {
        private ObservableCollection<CustomPropertyItem> _allProperties;
        private ObservableCollection<CustomPropertyItem> _existingProperties;
        private ObservableCollection<CustomPropertyItem> _boxModuleProperties;
        private string _documentPath;
        private string _documentName;
        
        // Propriétés Box Module prédéfinies
        private readonly Dictionary<string, string> _boxModuleDefaults = new Dictionary<string, string>
        {
            { "Base_Thermal_Break", "NTM J-Plastic ou Polyblock (See Note Importante)/ NTM  J-Plastic ou Polyblock (Voir Note Importante)" },
            { "Coating", "No Paint/ Non Peint" },
            { "Floor_Insulation", "Injected Polyurethane/Mousse Injecté" },
            { "Floor_Mount_Type", "Steel_Dunnage/Égale au Périmètre" },
            { "Floor_Note", "Installer des J Plastic au Plancher" },
            { "Hardware_Material", "Standard According to Material / Matériel Selon Standard" },
            { "Panel_Construction", "NTM (J-Plastic)" },
            { "Panel_Insulation", "Injected Polyurethane/ Mousse Injecté" },
            { "Unit_Certification", "Yes/Oui (Voir Notes)" },
            { "Unit_Configuration", "Stacked/ Empilé" },
            { "Unit_Option", "Washable/Lavable (aucune vis au plancher)" },
            { "Unit_Type", "Outdoor/ Extérieur" }
        };

        public CustomPropertyBatchWindow()
        {
            InitializeComponent();
            DataContext = this;
            
            _allProperties = new ObservableCollection<CustomPropertyItem>();
            _existingProperties = new ObservableCollection<CustomPropertyItem>();
            _boxModuleProperties = new ObservableCollection<CustomPropertyItem>();
            
            // Configurer le DataGrid principal
            DgProperties.ItemsSource = _existingProperties;
            
            // Désactiver le tri pendant l'édition pour éviter l'erreur "Sorting is not allowed during an AddNew or EditItem transaction"
            DgProperties.BeginningEdit += (s, e) =>
            {
                DgProperties.CanUserSortColumns = false;
            };
            DgProperties.CellEditEnding += (s, e) =>
            {
                DgProperties.CanUserSortColumns = true;
            };
            DgProperties.RowEditEnding += (s, e) =>
            {
                DgProperties.CanUserSortColumns = true;
            };
            
            // Gestion du placeholder pour la recherche
            TxtSearch.GotFocus += (s, e) =>
            {
                if (TxtSearch.Text == "Rechercher une propriété...")
                    TxtSearch.Text = "";
            };
            
            TxtSearch.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(TxtSearch.Text))
                    TxtSearch.Text = "Rechercher une propriété...";
            };
        }

        /// <summary>
        /// Initialise la fenêtre avec les propriétés du document
        /// </summary>
        public void InitializeProperties(Dictionary<string, string> existingProps, string documentPath, string documentName)
        {
            _documentPath = documentPath;
            _documentName = documentName;
            
            TxtDocumentInfo.Text = $"Document: {documentName}";
            TxtDocumentPath.Text = documentPath;
            
            _allProperties.Clear();
            _existingProperties.Clear();
            _boxModuleProperties.Clear();
            
            // Ajouter les propriétés existantes
            var existingPropNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var prop in existingProps)
            {
                var item = new CustomPropertyItem
                {
                    Name = prop.Key,
                    Value = prop.Value ?? "",
                    Type = PropertyType.Existing,
                    IsBoxModule = _boxModuleDefaults.ContainsKey(prop.Key)
                };
                
                _allProperties.Add(item);
                _existingProperties.Add(item);
                existingPropNames.Add(prop.Key);
            }
            
            // Ajouter les propriétés Box Module qui n'existent pas encore
            foreach (var boxProp in _boxModuleDefaults)
            {
                if (!existingPropNames.Contains(boxProp.Key))
                {
                    var item = new CustomPropertyItem
                    {
                        Name = boxProp.Key,
                        Value = boxProp.Value,
                        Type = PropertyType.BoxModule,
                        IsBoxModule = true
                    };
                    
                    _allProperties.Add(item);
                    _boxModuleProperties.Add(item);
                }
                else
                {
                    // Si elle existe déjà, l'ajouter aussi à Box Module pour l'affichage
                    var existingItem = _existingProperties.FirstOrDefault(p => 
                        string.Equals(p.Name, boxProp.Key, StringComparison.OrdinalIgnoreCase));
                    if (existingItem != null)
                    {
                        existingItem.IsBoxModule = true;
                        _boxModuleProperties.Add(existingItem);
                    }
                }
            }
            
            UpdateStats();
        }

        /// <summary>
        /// Retourne toutes les propriétés modifiées
        /// </summary>
        public Dictionary<string, string> GetProperties()
        {
            var result = new Dictionary<string, string>();
            
            foreach (var prop in _allProperties)
            {
                if (!string.IsNullOrWhiteSpace(prop.Name) && 
                    !prop.Name.StartsWith("---") &&
                    !string.IsNullOrWhiteSpace(prop.Value))
                {
                    result[prop.Name.Trim()] = prop.Value.Trim();
                }
            }
            
            return result;
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = TxtSearch.Text?.ToLowerInvariant() ?? "";
            
            if (string.IsNullOrWhiteSpace(searchText) || searchText == "rechercher une propriété...")
            {
                // Afficher toutes les propriétés selon l'onglet actif
                if (TabSections.SelectedItem == TabExisting)
                {
                    DgProperties.ItemsSource = _existingProperties;
                }
                else
                {
                    DgProperties.ItemsSource = _boxModuleProperties;
                }
                return;
            }
            
            // Filtrer les propriétés
            var filtered = _allProperties.Where(p => 
                p.Name.ToLowerInvariant().Contains(searchText) ||
                p.Value.ToLowerInvariant().Contains(searchText)).ToList();
            
            var filteredCollection = new ObservableCollection<CustomPropertyItem>(filtered);
            DgProperties.ItemsSource = filteredCollection;
        }

        private void BtnAddProperty_Click(object sender, RoutedEventArgs e)
        {
            var newItem = new CustomPropertyItem
            {
                Name = "",
                Value = "",
                Type = PropertyType.Custom,
                IsBoxModule = false
            };
            
            _allProperties.Add(newItem);
            _existingProperties.Add(newItem);
            
            // Sélectionner la nouvelle ligne
            DgProperties.SelectedItem = newItem;
            DgProperties.ScrollIntoView(newItem);
            DgProperties.Focus();
            
            // Commencer l'édition du nom
            DgProperties.CurrentCell = new DataGridCellInfo(newItem, DgProperties.Columns[0]);
            DgProperties.BeginEdit();
        }

        private void BtnClearBoxModule_Click(object sender, RoutedEventArgs e)
        {
            // Utiliser le message box custom de l'application
            var result = XnrgyEngineeringAutomationTools.Shared.Views.XnrgyMessageBox.Show(
                "Voulez-vous vraiment vider toutes les valeurs des propriétés Box Module ?",
                "Confirmation",
                XnrgyEngineeringAutomationTools.Shared.Views.XnrgyMessageBoxType.Question,
                XnrgyEngineeringAutomationTools.Shared.Views.XnrgyMessageBoxButtons.YesNo,
                this);
            
            if (result == XnrgyEngineeringAutomationTools.Shared.Views.XnrgyMessageBoxResult.Yes)
            {
                foreach (var prop in _boxModuleProperties)
                {
                    if (prop.IsBoxModule && _boxModuleDefaults.ContainsKey(prop.Name))
                    {
                        // Réinitialiser à la valeur par défaut
                        prop.Value = _boxModuleDefaults[prop.Name];
                    }
                    else
                    {
                        prop.Value = "";
                    }
                }
                
                TxtStatus.Text = "✅ Valeurs Box Module réinitialisées";
                TxtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27ae60"));
                AddOperationLog("Valeurs Box Module réinitialisées", "SUCCESS");
                UpdateStats();
            }
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            // Valider les propriétés
            var errors = new List<string>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            foreach (var prop in _allProperties)
            {
                if (string.IsNullOrWhiteSpace(prop.Name))
                    continue;
                
                var name = prop.Name.Trim();
                if (names.Contains(name))
                {
                    errors.Add($"La propriété '{name}' est dupliquée.");
                }
                else
                {
                    names.Add(name);
                }
            }
            
            if (errors.Count > 0)
            {
                // Utiliser le message box custom de l'application
                XnrgyEngineeringAutomationTools.Shared.Views.XnrgyMessageBox.ShowError(
                    "Erreurs de validation:\n" + string.Join("\n", errors),
                    "Erreur de validation",
                    this);
                return;
            }
            
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// Permet l'édition avec un simple clic sur une cellule
        /// </summary>
        private void DgProperties_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var dataGrid = sender as DataGrid;
            if (dataGrid == null) return;

            // Trouver la cellule sous le curseur
            var hit = VisualTreeHelper.HitTest(dataGrid, e.GetPosition(dataGrid));
            if (hit == null) return;

            var cell = FindParent<DataGridCell>(hit.VisualHit);
            if (cell == null || cell.IsEditing || cell.IsReadOnly) return;

            // Si on clique sur une colonne éditable (Nom ou Valeur), commencer l'édition
            var column = cell.Column;
            if (column != null && (column.DisplayIndex == 0 || column.DisplayIndex == 1))
            {
                dataGrid.CurrentCell = new DataGridCellInfo(cell);
                dataGrid.SelectedItem = cell.DataContext;
                dataGrid.BeginEdit();
                e.Handled = true;
            }
        }

        /// <summary>
        /// Trouve le parent d'un type donné dans l'arbre visuel
        /// </summary>
        private static T? FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            var parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            
            if (parentObject is T parent)
                return parent;
            else
                return FindParent<T>(parentObject);
        }

        private void UpdateStats()
        {
            var total = _allProperties.Count(p => !string.IsNullOrWhiteSpace(p.Name));
            var existing = _existingProperties.Count(p => !string.IsNullOrWhiteSpace(p.Name));
            var boxModule = _boxModuleProperties.Count(p => !string.IsNullOrWhiteSpace(p.Name));
            
            TxtStats.Text = $"{total} propriétés ({existing} existantes, {boxModule} Box Module)";
        }

        /// <summary>
        /// Affiche la barre de progression
        /// </summary>
        public void ShowProgress(string label = "⏳ Traitement en cours...")
        {
            Dispatcher.Invoke(() =>
            {
                TxtProgressLabel.Text = label;
                TxtProgressLabel.Visibility = Visibility.Visible;
                PbProgress.Visibility = Visibility.Visible;
                TxtProgressPercent.Visibility = Visibility.Visible;
                PbProgress.Value = 0;
            });
        }

        /// <summary>
        /// Met à jour la barre de progression
        /// </summary>
        public void UpdateProgress(int current, int total, string? label = null)
        {
            Dispatcher.Invoke(() =>
            {
                if (label != null)
                    TxtProgressLabel.Text = label;
                
                double percent = total > 0 ? (current * 100.0 / total) : 0;
                PbProgress.Value = percent;
                TxtProgressPercent.Text = $"{percent:F0}%";
            });
        }

        /// <summary>
        /// Cache la barre de progression
        /// </summary>
        public void HideProgress()
        {
            Dispatcher.Invoke(() =>
            {
                TxtProgressLabel.Visibility = Visibility.Collapsed;
                PbProgress.Visibility = Visibility.Collapsed;
                TxtProgressPercent.Visibility = Visibility.Collapsed;
            });
        }

        /// <summary>
        /// Ajoute une entrée au journal d'opérations
        /// </summary>
        public void AddOperationLog(string message, string type = "INFO")
        {
            Dispatcher.Invoke(() =>
            {
                BorderOperationLog.Visibility = Visibility.Visible;
                
                string icon = type switch
                {
                    "SUCCESS" => "✅",
                    "ERROR" => "❌",
                    "WARNING" => "⚠️",
                    "INFO" => "ℹ️",
                    _ => "•"
                };
                
                string color = type switch
                {
                    "SUCCESS" => "#27ae60",
                    "ERROR" => "#e74c3c",
                    "WARNING" => "#f39c12",
                    "INFO" => "#3498db",
                    _ => "#2c3e50"
                };
                
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                
                // Créer un Run avec la couleur appropriée
                var run = new System.Windows.Documents.Run($"[{timestamp}] {icon} {message}\n");
                run.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));
                
                // Ajouter au TextBlock (qui doit être un RichTextBox ou utiliser Inlines)
                // Pour simplifier, on utilise un TextBlock avec formatage simple
                if (TxtOperationLog.Inlines.Count > 0)
                {
                    TxtOperationLog.Inlines.Add(new System.Windows.Documents.LineBreak());
                }
                TxtOperationLog.Inlines.Add(run);
                
                // Scroll vers le bas
                SvOperationLog.ScrollToEnd();
            });
        }

        /// <summary>
        /// Affiche le résultat final dans la barre de statut
        /// </summary>
        public void ShowResult(int addedCount, int updatedCount, int errorCount)
        {
            Dispatcher.Invoke(() =>
            {
                HideProgress();
                
                if (errorCount > 0)
                {
                    TxtStatus.Text = $"❌ Terminé avec {errorCount} erreur(s)";
                    TxtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e74c3c"));
                }
                else if (addedCount > 0 || updatedCount > 0)
                {
                    TxtStatus.Text = $"✅ Succès: {addedCount} ajoutée(s), {updatedCount} mise(s) à jour";
                    TxtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27ae60"));
                }
                else
                {
                    TxtStatus.Text = "ℹ️ Aucune modification";
                    TxtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3498db"));
                }
                
                // Ajouter au journal
                if (addedCount > 0 || updatedCount > 0 || errorCount > 0)
                {
                    AddOperationLog($"Traitement terminé: {addedCount} ajoutée(s), {updatedCount} mise(s) à jour, {errorCount} erreur(s)", 
                        errorCount > 0 ? "ERROR" : "SUCCESS");
                }
            });
        }

        /// <summary>
        /// Réinitialise le journal et le statut
        /// </summary>
        public void ResetStatus()
        {
            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = "Prêt";
                TxtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7f8c8d"));
                TxtOperationLog.Inlines.Clear();
                BorderOperationLog.Visibility = Visibility.Collapsed;
                HideProgress();
            });
        }

        private void TabSections_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TabSections.SelectedItem == TabExisting)
            {
                DgProperties.ItemsSource = _existingProperties;
            }
            else if (TabSections.SelectedItem == TabBoxModule)
            {
                DgProperties.ItemsSource = _boxModuleProperties;
            }
            
            // Réappliquer le filtre de recherche si actif
            if (TxtSearch != null)
                TxtSearch_TextChanged(TxtSearch, null);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Item représentant une propriété personnalisée
    /// </summary>
    public class CustomPropertyItem : INotifyPropertyChanged
    {
        private string _name;
        private string _value;
        private PropertyType _type;
        private bool _isBoxModule;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged();
            }
        }

        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                OnPropertyChanged();
            }
        }

        public PropertyType Type
        {
            get => _type;
            set
            {
                _type = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TypeLabel));
                OnPropertyChanged(nameof(TypeColor));
            }
        }

        public bool IsBoxModule
        {
            get => _isBoxModule;
            set
            {
                _isBoxModule = value;
                OnPropertyChanged();
            }
        }

        public string TypeLabel
        {
            get
            {
                if (IsBoxModule)
                    return "Box Module";
                
                return Type switch
                {
                    PropertyType.Existing => "Existante",
                    PropertyType.BoxModule => "Box Module",
                    PropertyType.Custom => "Personnalisée",
                    _ => "Autre"
                };
            }
        }

        public Brush TypeColor
        {
            get
            {
                if (IsBoxModule)
                    return new SolidColorBrush(Color.FromRgb(230, 126, 34)); // Orange
                
                return Type switch
                {
                    PropertyType.Existing => new SolidColorBrush(Color.FromRgb(52, 152, 219)), // Bleu
                    PropertyType.BoxModule => new SolidColorBrush(Color.FromRgb(230, 126, 34)), // Orange
                    PropertyType.Custom => new SolidColorBrush(Color.FromRgb(46, 204, 113)), // Vert
                    _ => new SolidColorBrush(Color.FromRgb(149, 165, 166)) // Gris
                };
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Type de propriété
    /// </summary>
    public enum PropertyType
    {
        Existing,   // Propriété existante dans le document
        BoxModule,  // Propriété Box Module prédéfinie
        Custom      // Propriété personnalisée ajoutée par l'utilisateur
    }
}

