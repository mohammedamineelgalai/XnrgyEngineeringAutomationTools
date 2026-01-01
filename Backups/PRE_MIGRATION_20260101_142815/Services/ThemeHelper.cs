using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Helper pour appliquer le theme aux sous-formulaires
    /// Fournit des methodes utilitaires pour uniformiser l'apparence
    /// </summary>
    public static class ThemeHelper
    {
        // === COULEURS THEME SOMBRE ===
        public static readonly Color DarkBackground = Color.FromRgb(30, 30, 46);       // #1E1E2E
        public static readonly Color DarkPanel = Color.FromRgb(37, 37, 54);            // #252536
        public static readonly Color DarkInput = Color.FromRgb(45, 45, 68);            // #2D2D44
        public static readonly Color DarkBorder = Color.FromRgb(64, 64, 96);           // #404060
        public static readonly Color DarkStatusBar = Color.FromRgb(26, 26, 40);        // #1A1A28 - FIXE
        
        // === COULEURS THEME CLAIR ===
        public static readonly Color LightBackground = Color.FromRgb(245, 247, 250);   // #F5F7FA
        public static readonly Color LightPanel = Color.FromRgb(252, 253, 255);        // #FCFDFF
        public static readonly Color LightInput = Color.FromRgb(240, 245, 252);        // #F0F5FC
        public static readonly Color LightBorder = Color.FromRgb(200, 210, 225);       // #C8D2E1
        
        // === COULEURS FIXES (ne changent pas avec le theme) ===
        public static readonly Color BleuMarine = Color.FromRgb(42, 74, 111);          // #2A4A6F
        public static readonly Color BleuMarineClair = Color.FromRgb(58, 90, 127);     // #3A5A7F
        public static readonly Color StatusBarBackground = Color.FromRgb(26, 26, 40);  // #1A1A28 - Toujours noir
        
        // === COULEURS TEXTE ===
        public static readonly Color TextWhite = Colors.White;
        public static readonly Color TextDark = Color.FromRgb(30, 30, 30);
        public static readonly Color TextMuted = Color.FromRgb(128, 128, 128);
        
        // === COULEURS ACCENT ===
        public static readonly Color Violet = Color.FromRgb(124, 58, 237);             // #7C3AED
        public static readonly Color Vert = Color.FromRgb(16, 124, 16);                // #107C10
        public static readonly Color Rouge = Color.FromRgb(232, 17, 35);               // #E81123
        public static readonly Color Orange = Color.FromRgb(255, 140, 0);              // #FF8C00
        public static readonly Color Bleu = Color.FromRgb(0, 120, 212);                // #0078D4
        public static readonly Color Cyan = Color.FromRgb(8, 145, 178);                // #0891B2

        /// <summary>
        /// Obtient le theme actuel depuis MainWindow
        /// </summary>
        public static bool IsDarkTheme => MainWindow.CurrentThemeIsDark;

        /// <summary>
        /// Obtient la couleur de fond principale selon le theme
        /// </summary>
        public static SolidColorBrush GetBackgroundBrush()
        {
            return new SolidColorBrush(IsDarkTheme ? DarkBackground : LightBackground);
        }

        /// <summary>
        /// Obtient la couleur de fond des panneaux selon le theme
        /// </summary>
        public static SolidColorBrush GetPanelBrush()
        {
            return new SolidColorBrush(IsDarkTheme ? DarkPanel : LightPanel);
        }

        /// <summary>
        /// Obtient la couleur de fond des inputs selon le theme
        /// </summary>
        public static SolidColorBrush GetInputBrush()
        {
            return new SolidColorBrush(IsDarkTheme ? DarkInput : LightInput);
        }

        /// <summary>
        /// Obtient la couleur de bordure selon le theme
        /// </summary>
        public static SolidColorBrush GetBorderBrush()
        {
            return new SolidColorBrush(IsDarkTheme ? DarkBorder : LightBorder);
        }

        /// <summary>
        /// Obtient la couleur de texte principale selon le theme
        /// </summary>
        public static SolidColorBrush GetTextBrush()
        {
            return new SolidColorBrush(IsDarkTheme ? TextWhite : TextDark);
        }

        /// <summary>
        /// Obtient la couleur bleu marine (fixe)
        /// </summary>
        public static SolidColorBrush GetBleuMarineBrush()
        {
            return new SolidColorBrush(BleuMarine);
        }

        /// <summary>
        /// Obtient la couleur de fond de la barre de statut (toujours noir)
        /// </summary>
        public static SolidColorBrush GetStatusBarBrush()
        {
            return new SolidColorBrush(StatusBarBackground);
        }

        /// <summary>
        /// Applique le theme a une fenetre complete
        /// </summary>
        public static void ApplyThemeToWindow(Window window)
        {
            if (window.Content is Grid grid)
            {
                grid.Background = GetBackgroundBrush();
            }
            else if (window.Content is Border border)
            {
                border.Background = GetBackgroundBrush();
            }
        }

        /// <summary>
        /// Applique le theme a un GroupBox
        /// </summary>
        public static void ApplyThemeToGroupBox(GroupBox groupBox)
        {
            groupBox.Background = GetPanelBrush();
            groupBox.BorderBrush = GetBleuMarineBrush();
        }

        /// <summary>
        /// Applique le theme a un TextBox
        /// </summary>
        public static void ApplyThemeToTextBox(TextBox textBox)
        {
            textBox.Background = GetInputBrush();
            textBox.Foreground = GetTextBrush();
            textBox.BorderBrush = GetBorderBrush();
        }

        /// <summary>
        /// Applique le theme a un TextBlock
        /// </summary>
        public static void ApplyThemeToTextBlock(TextBlock textBlock)
        {
            textBlock.Foreground = GetTextBrush();
        }

        /// <summary>
        /// Applique le theme a un Button
        /// </summary>
        public static void ApplyThemeToButton(Button button, bool isPrimary = false)
        {
            if (isPrimary)
            {
                button.Background = new SolidColorBrush(Bleu);
                button.Foreground = Brushes.White;
            }
            else
            {
                button.Background = GetInputBrush();
                button.Foreground = GetTextBrush();
                button.BorderBrush = GetBorderBrush();
            }
        }

        /// <summary>
        /// Applique le theme a un DataGrid
        /// </summary>
        public static void ApplyThemeToDataGrid(DataGrid dataGrid)
        {
            dataGrid.Background = GetBackgroundBrush();
            dataGrid.Foreground = GetTextBrush();
            dataGrid.BorderBrush = GetBorderBrush();
            dataGrid.RowBackground = GetPanelBrush();
            dataGrid.AlternatingRowBackground = GetInputBrush();
        }

        /// <summary>
        /// Applique le theme a un ListBox (journal)
        /// </summary>
        public static void ApplyThemeToListBox(ListBox listBox)
        {
            listBox.Background = IsDarkTheme 
                ? new SolidColorBrush(Color.FromRgb(18, 18, 28)) 
                : new SolidColorBrush(Color.FromRgb(252, 253, 255));
            listBox.Foreground = GetTextBrush();
            listBox.BorderBrush = GetBleuMarineBrush();
        }
    }
}
