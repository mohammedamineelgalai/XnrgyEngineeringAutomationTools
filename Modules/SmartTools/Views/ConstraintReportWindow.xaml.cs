using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenêtre WPF moderne pour afficher le rapport de contraintes d'assemblage
    /// Migrée depuis ConstraintReport.vb de l'addin iLogic
    /// </summary>
    public partial class ConstraintReportWindow : Window
    {
        private List<ComponentConstraintInfo> _components = new List<ComponentConstraintInfo>();

        #region Win32 API pour centrage sur Inventor
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);
        
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion

        public ConstraintReportWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Centrer sur Inventor
            CenterOnInventorWindow();
            
            // Animation d'entrée si les données sont chargées
            if (_components.Count > 0)
            {
                PopulateComponentsList();
            }
        }

        /// <summary>
        /// Centre la fenêtre sur la fenêtre Inventor
        /// </summary>
        private void CenterOnInventorWindow()
        {
            try
            {
                IntPtr inventorHandle = IntPtr.Zero;
                RECT inventorRect = new RECT();

                // Chercher la fenêtre Inventor
                EnumWindows((hWnd, lParam) =>
                {
                    if (IsWindowVisible(hWnd))
                    {
                        int length = GetWindowTextLength(hWnd);
                        if (length > 0)
                        {
                            var sb = new System.Text.StringBuilder(length + 1);
                            GetWindowText(hWnd, sb, sb.Capacity);
                            string title = sb.ToString();

                            if (title.Contains("Autodesk Inventor") || title.Contains("Inventor 20"))
                            {
                                if (GetWindowRect(hWnd, out RECT rect))
                                {
                                    int width = rect.Right - rect.Left;
                                    int height = rect.Bottom - rect.Top;
                                    if (width > 800 && height > 600)
                                    {
                                        inventorHandle = hWnd;
                                        inventorRect = rect;
                                        return false; // Arrêter l'énumération
                                    }
                                }
                            }
                        }
                    }
                    return true; // Continuer l'énumération
                }, IntPtr.Zero);

                // Si Inventor trouvé, centrer dessus
                if (inventorHandle != IntPtr.Zero)
                {
                    int inventorWidth = inventorRect.Right - inventorRect.Left;
                    int inventorHeight = inventorRect.Bottom - inventorRect.Top;
                    int inventorCenterX = inventorRect.Left + inventorWidth / 2;
                    int inventorCenterY = inventorRect.Top + inventorHeight / 2;

                    this.Left = inventorCenterX - this.Width / 2;
                    this.Top = inventorCenterY - this.Height / 2;

                    // S'assurer que la fenêtre reste visible à l'écran
                    var screen = System.Windows.SystemParameters.WorkArea;
                    if (this.Left < 0) this.Left = 10;
                    if (this.Top < 0) this.Top = 10;
                    if (this.Left + this.Width > screen.Width) this.Left = screen.Width - this.Width - 10;
                    if (this.Top + this.Height > screen.Height) this.Top = screen.Height - this.Height - 10;
                }
                else
                {
                    // Fallback: centrer sur l'écran
                    var screen = System.Windows.SystemParameters.WorkArea;
                    this.Left = (screen.Width - this.Width) / 2;
                    this.Top = (screen.Height - this.Height) / 2;
                }
            }
            catch
            {
                // En cas d'erreur, centrer sur l'écran
                var screen = System.Windows.SystemParameters.WorkArea;
                this.Left = (screen.Width - this.Width) / 2;
                this.Top = (screen.Height - this.Height) / 2;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Configure les informations du document
        /// </summary>
        public void SetDocumentInfo(string assemblyName, int totalComponents, DateTime reportDate)
        {
            TxtAssemblyName.Text = $"📁 {assemblyName}";
            TxtDocName.Text = assemblyName;
            TxtDate.Text = reportDate.ToString("dd-MM-yyyy HH:mm:ss");
            TxtTotalComponents.Text = totalComponents.ToString();
        }

        /// <summary>
        /// Configure les statistiques
        /// </summary>
        public void SetStatistics(int grounded, int fullyConstrained, int underConstrained)
        {
            TxtGroundedCount.Text = grounded.ToString();
            TxtFullyConstrainedCount.Text = fullyConstrained.ToString();
            TxtUnderConstrainedCount.Text = underConstrained.ToString();
        }

        /// <summary>
        /// Ajoute les composants à afficher
        /// </summary>
        public void SetComponents(List<ComponentConstraintInfo> components)
        {
            _components = components ?? new List<ComponentConstraintInfo>();
            
            if (IsLoaded)
            {
                PopulateComponentsList();
            }
        }

        private void PopulateComponentsList()
        {
            ComponentsList.Children.Clear();
            int rowNumber = 1;

            foreach (var component in _components)
            {
                var row = CreateComponentRow(rowNumber, component);
                ComponentsList.Children.Add(row);
                rowNumber++;
            }
        }

        private Border CreateComponentRow(int rowNumber, ComponentConstraintInfo component)
        {
            // Déterminer les couleurs selon le statut
            Color statusColor;
            string statusIcon;
            Color rowBackground;

            switch (component.Status)
            {
                case ConstraintStatus.Grounded:
                    statusColor = (Color)ColorConverter.ConvertFromString("#2e7d32");
                    statusIcon = "⚓";
                    rowBackground = (Color)ColorConverter.ConvertFromString("#e8f5e9");
                    break;
                case ConstraintStatus.FullyConstrained:
                    statusColor = (Color)ColorConverter.ConvertFromString("#1976d2");
                    statusIcon = "✓";
                    rowBackground = (Color)ColorConverter.ConvertFromString("#e3f2fd");
                    break;
                default: // UnderConstrained
                    statusColor = (Color)ColorConverter.ConvertFromString("#ff6f00");
                    statusIcon = "!";
                    rowBackground = (Color)ColorConverter.ConvertFromString("#fff3e0");
                    break;
            }

            // Alternance de couleurs
            if (rowNumber % 2 == 0)
            {
                rowBackground = Color.FromArgb(255, 
                    (byte)Math.Max(0, rowBackground.R - 10),
                    (byte)Math.Max(0, rowBackground.G - 10),
                    (byte)Math.Max(0, rowBackground.B - 10));
            }

            var rowBorder = new Border
            {
                Background = new SolidColorBrush(rowBackground),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(15, 10, 15, 10),
                Margin = new Thickness(0, 2, 0, 2)
            };

            // Effet hover
            rowBorder.MouseEnter += (s, e) =>
            {
                rowBorder.Background = new LinearGradientBrush(
                    Color.FromRgb(227, 242, 253),
                    Color.FromRgb(187, 222, 251),
                    0);
            };
            rowBorder.MouseLeave += (s, e) =>
            {
                rowBorder.Background = new SolidColorBrush(rowBackground);
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(250) });

            // Numéro de ligne
            var numText = new TextBlock
            {
                Text = rowNumber.ToString(),
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(55, 71, 79)),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(numText, 0);
            grid.Children.Add(numText);

            // Nom du composant
            var namePanel = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            var nameIcon = new TextBlock { Text = "🔧", FontSize = 14, Margin = new Thickness(0, 0, 8, 0) };
            var nameText = new TextBlock
            {
                Text = component.Name,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(44, 62, 80)),
                TextTrimming = TextTrimming.CharacterEllipsis,
                MaxWidth = 300
            };
            namePanel.Children.Add(nameIcon);
            namePanel.Children.Add(nameText);
            Grid.SetColumn(namePanel, 1);
            grid.Children.Add(namePanel);

            // Statut
            var statusPanel = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            var statusBadge = new Border
            {
                Width = 20,
                Height = 20,
                CornerRadius = new CornerRadius(10),
                Background = new SolidColorBrush(statusColor),
                Margin = new Thickness(0, 0, 8, 0)
            };
            var badgeText = new TextBlock
            {
                Text = statusIcon,
                FontSize = 11,
                Foreground = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontWeight = FontWeights.Bold
            };
            statusBadge.Child = badgeText;
            var statusTextBlock = new TextBlock
            {
                Text = component.StatusText,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(statusColor)
            };
            statusPanel.Children.Add(statusBadge);
            statusPanel.Children.Add(statusTextBlock);
            Grid.SetColumn(statusPanel, 2);
            grid.Children.Add(statusPanel);

            // Type
            var typeText = new TextBlock
            {
                Text = component.ComponentType,
                Foreground = new SolidColorBrush(Color.FromRgb(109, 76, 65)),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(typeText, 3);
            grid.Children.Add(typeText);

            // Analyse
            var analysisPanel = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            var analysisIcon = new TextBlock { Text = "📊", FontSize = 12, Margin = new Thickness(0, 0, 5, 0) };
            var analysisText = new TextBlock
            {
                Text = component.AnalysisMethod,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(93, 64, 55)),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            analysisPanel.Children.Add(analysisIcon);
            analysisPanel.Children.Add(analysisText);
            Grid.SetColumn(analysisPanel, 4);
            grid.Children.Add(analysisPanel);

            rowBorder.Child = grid;
            return rowBorder;
        }
    }

    /// <summary>
    /// Statut de contrainte d'un composant
    /// </summary>
    public enum ConstraintStatus
    {
        Grounded,
        FullyConstrained,
        UnderConstrained
    }

    /// <summary>
    /// Informations sur les contraintes d'un composant
    /// </summary>
    public class ComponentConstraintInfo
    {
        public string Name { get; set; } = "";
        public ConstraintStatus Status { get; set; } = ConstraintStatus.UnderConstrained;
        public string StatusText { get; set; } = "Partiellement contraint";
        public string ComponentType { get; set; } = "Pièce";
        public string AnalysisMethod { get; set; } = "";
        public int DegreesOfFreedom { get; set; } = -1;
        public int ActiveConstraints { get; set; } = 0;
        public int CriticalConstraints { get; set; } = 0;
    }
}

