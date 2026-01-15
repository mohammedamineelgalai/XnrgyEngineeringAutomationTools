using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenêtre personnalisée pour sélectionner un dossier avec vue arborescente
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public partial class FolderBrowserWindow : System.Windows.Window
    {
        public string SelectedPath { get; private set; } = "";

        public FolderBrowserWindow()
        {
            InitializeComponent();
            Loaded += FolderBrowserWindow_Loaded;
        }

        private void FolderBrowserWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Initialiser avec les disques locaux
            LoadDrives();
        }

        /// <summary>
        /// Affiche la fenêtre et retourne le chemin sélectionné
        /// </summary>
        public static string ShowDialog(Window owner, string initialPath = "")
        {
            try
            {
                var window = new FolderBrowserWindow();
                
                if (owner != null)
                {
                    window.Owner = owner;
                }
                
                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath))
                {
                    window.TxtCurrentPath.Text = initialPath;
                    window.LoadPath(initialPath);
                }

                if (window.ShowDialog() == true && !string.IsNullOrEmpty(window.SelectedPath))
                {
                    return window.SelectedPath;
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur FolderBrowserWindow: {ex.Message}");
                MessageBox.Show($"Erreur lors de l'ouverture de la fenêtre de sélection: {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private void LoadDrives()
        {
            FolderTreeView.Items.Clear();
            
            try
            {
                foreach (DriveInfo drive in DriveInfo.GetDrives())
                {
                    if (drive.IsReady)
                    {
                        var item = CreateTreeViewItem(drive.Name, drive.Name, true);
                        item.Tag = drive.RootDirectory.FullName;
                        
                        // Ajouter un nœud placeholder pour permettre l'expansion
                        var placeholder = new TreeViewItem { Header = "Chargement..." };
                        item.Items.Add(placeholder);
                        
                        FolderTreeView.Items.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des disques: {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private TreeViewItem CreateTreeViewItem(string displayName, string fullPath, bool isDrive = false)
        {
            var item = new TreeViewItem
            {
                Header = displayName,
                Tag = fullPath,
                FontSize = 13,
                Foreground = new SolidColorBrush(Colors.White)
            };

                // Icône selon le type
                if (isDrive)
                {
                    item.Header = $"💾 {displayName}";
                }
                else if (Directory.Exists(fullPath))
                {
                    item.Header = $"📁 {displayName}";
                }

            // Gestion de l'expansion
            item.Expanded += Item_Expanded;
            
            // Ajouter placeholder pour permettre l'expansion
            if (isDrive || Directory.Exists(fullPath))
            {
                var placeholder = new TreeViewItem { Header = "Chargement..." };
                item.Items.Add(placeholder);
            }

            return item;
        }

        private void Item_Expanded(object sender, RoutedEventArgs e)
        {
            var item = sender as TreeViewItem;
            if (item == null) return;

            string? path = item.Tag as string;
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) return;

            // Supprimer le placeholder
            item.Items.Clear();

            try
            {
                // Charger les sous-dossiers
                var directories = Directory.GetDirectories(path)
                    .OrderBy(d => Path.GetFileName(d))
                    .ToList();

                foreach (var dir in directories)
                {
                    try
                    {
                        string dirName = Path.GetFileName(dir);
                        if (string.IsNullOrEmpty(dirName)) continue;

                        var subItem = CreateTreeViewItem(dirName, dir);
                        item.Items.Add(subItem);
                    }
                    catch
                    {
                        // Ignorer les dossiers inaccessibles
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Dossier non accessible - ne rien ajouter
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur chargement sous-dossiers: {ex.Message}");
            }
        }

        private void FolderTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (FolderTreeView.SelectedItem is TreeViewItem selectedItem)
            {
                string? path = selectedItem.Tag as string;
                if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                {
                    TxtSelectedPath.Text = path;
                    SelectedPath = path;
                }
            }
        }

        private void FolderTreeView_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Double-clic sur un dossier = sélectionner et fermer
            if (FolderTreeView.SelectedItem is TreeViewItem selectedItem)
            {
                string? path = selectedItem.Tag as string;
                if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                {
                    SelectedPath = path;
                    DialogResult = true;
                    Close();
                }
            }
        }

        private void BtnGoToPath_Click(object sender, RoutedEventArgs e)
        {
            string path = TxtCurrentPath.Text.Trim();
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
            {
                LoadPath(path);
            }
            else
            {
                MessageBox.Show("Le chemin spécifié n'existe pas ou n'est pas un dossier valide.",
                    "Chemin invalide", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadPath(string path)
        {
            try
            {
                FolderTreeView.Items.Clear();
                
                // Créer l'arborescence jusqu'au chemin spécifié
                string currentPath = path;
                var pathParts = new Stack<string>();
                
                while (!string.IsNullOrEmpty(currentPath) && Directory.Exists(currentPath))
                {
                    pathParts.Push(currentPath);
                    currentPath = Path.GetDirectoryName(currentPath);
                    if (currentPath == pathParts.Peek()) break; // Éviter boucle infinie
                }

                // Charger depuis la racine
                if (pathParts.Count > 0)
                {
                    string rootPath = pathParts.Pop();
                    var rootItem = CreateTreeViewItem(Path.GetFileName(rootPath) ?? rootPath, rootPath);
                    
                    // Construire l'arborescence
                    TreeViewItem currentItem = rootItem;
                    while (pathParts.Count > 0)
                    {
                        string nextPath = pathParts.Pop();
                        var nextItem = CreateTreeViewItem(Path.GetFileName(nextPath) ?? nextPath, nextPath);
                        currentItem.Items.Add(nextItem);
                        currentItem.IsExpanded = true;
                        currentItem = nextItem;
                    }
                    
                    // Sélectionner le dernier élément
                    currentItem.IsSelected = true;
                    currentItem.BringIntoView();
                    
                    // Charger les sous-dossiers du dernier élément
                    Item_Expanded(currentItem, null);
                    
                    FolderTreeView.Items.Add(rootItem);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement du chemin: {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TxtCurrentPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Mettre à jour le chemin sélectionné si valide
            string path = TxtCurrentPath.Text.Trim();
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
            {
                TxtSelectedPath.Text = path;
                SelectedPath = path;
            }
        }

        private void BtnCreateFolder_Click(object sender, RoutedEventArgs e)
        {
            string basePath = TxtSelectedPath.Text;
            if (string.IsNullOrEmpty(basePath))
            {
                basePath = TxtCurrentPath.Text;
            }
            
            if (string.IsNullOrEmpty(basePath) || !Directory.Exists(basePath))
            {
                MessageBox.Show("Veuillez d'abord sélectionner un dossier valide.",
                    "Aucun dossier sélectionné", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Demander le nom du nouveau dossier avec une fenêtre WPF
            string folderName = ShowInputDialog(
                "Créer un dossier",
                "Entrez le nom du nouveau dossier:",
                "Nouveau dossier");

            if (string.IsNullOrWhiteSpace(folderName)) return;

            try
            {
                string newFolderPath = Path.Combine(basePath, folderName);
                Directory.CreateDirectory(newFolderPath);
                
                // Recharger le dossier parent
                LoadPath(basePath);
                
                // Sélectionner le nouveau dossier
                SelectPath(newFolderPath);
                
                MessageBox.Show($"Dossier créé avec succès: {folderName}",
                    "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la création du dossier: {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SelectPath(string path)
        {
            // Trouver et sélectionner le chemin dans l'arborescence
            foreach (TreeViewItem item in FolderTreeView.Items)
            {
                if (SelectPathRecursive(item, path))
                {
                    item.IsSelected = true;
                    item.BringIntoView();
                    break;
                }
            }
        }

        private bool SelectPathRecursive(TreeViewItem item, string targetPath)
        {
            string? itemPath = item.Tag as string;
            if (itemPath != null && itemPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase))
            {
                item.IsSelected = true;
                item.IsExpanded = true;
                return true;
            }

            foreach (TreeViewItem child in item.Items)
            {
                if (SelectPathRecursive(child, targetPath))
                {
                    item.IsExpanded = true;
                    return true;
                }
            }

            return false;
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedPath) || !Directory.Exists(SelectedPath))
            {
                MessageBox.Show("Veuillez sélectionner un dossier valide.",
                    "Aucun dossier sélectionné", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            SelectedPath = "";
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// Affiche une boîte de dialogue d'entrée personnalisée
        /// </summary>
        private string ShowInputDialog(string title, string prompt, string defaultValue = "")
        {
            var dialog = new Window
            {
                Title = title,
                Width = 400,
                Height = 180,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Background = new SolidColorBrush(Color.FromRgb(30, 30, 46)),
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false
            };

            var stackPanel = new StackPanel { Margin = new Thickness(20) };
            
            var promptText = new TextBlock
            {
                Text = prompt,
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 12)
            };
            stackPanel.Children.Add(promptText);

            var textBox = new TextBox
            {
                Text = defaultValue,
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 20),
                Background = new SolidColorBrush(Color.FromRgb(37, 37, 54)),
                Foreground = new SolidColorBrush(Colors.White),
                BorderBrush = new SolidColorBrush(Color.FromRgb(74, 127, 191)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(8, 6, 8, 6)
            };
            textBox.Focus();
            textBox.SelectAll();
            stackPanel.Children.Add(textBox);

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var okButton = new Button
            {
                Content = "OK",
                Width = 80,
                Height = 32,
                Margin = new Thickness(0, 0, 12, 0),
                Background = new SolidColorBrush(Color.FromRgb(0, 120, 212)),
                Foreground = new SolidColorBrush(Colors.White),
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            okButton.Click += (s, e) => { dialog.DialogResult = true; };

            var cancelButton = new Button
            {
                Content = "Annuler",
                Width = 80,
                Height = 32,
                Background = new SolidColorBrush(Color.FromRgb(107, 114, 128)),
                Foreground = new SolidColorBrush(Colors.White),
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            cancelButton.Click += (s, e) => { dialog.DialogResult = false; };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            stackPanel.Children.Add(buttonPanel);

            dialog.Content = stackPanel;
            textBox.KeyDown += (s, e) =>
            {
                if (e.Key == System.Windows.Input.Key.Enter)
                {
                    dialog.DialogResult = true;
                }
                else if (e.Key == System.Windows.Input.Key.Escape)
                {
                    dialog.DialogResult = false;
                }
            };

            if (dialog.ShowDialog() == true)
            {
                return textBox.Text.Trim();
            }

            return "";
        }
    }
}

