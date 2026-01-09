using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Fenêtre WPF moderne pour afficher la progression des opérations Smart Save/Close
    /// Avec icônes animées et fermeture automatique
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public partial class SmartProgressWindow : Window
    {
        private Dictionary<string, OperationItem> _operations = new Dictionary<string, OperationItem>();
        private int _totalOperations = 0;
        private int _completedOperations = 0;
        private bool _hasErrors = false;
        private int _autoCloseDelay = 2; // secondes
        private DispatcherTimer? _autoCloseTimer;
        private string _operationType = "save"; // "save" ou "close"

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
        public struct RECT { public int Left, Top, Right, Bottom; }

        public SmartProgressWindow()
        {
            InitializeComponent();
            Loaded += SmartProgressWindow_Loaded;
        }

        public SmartProgressWindow(string title, string icon = "⚡") : this()
        {
            TxtTitle.Text = title;
            HeaderIconText.Text = icon;
        }

        /// <summary>
        /// Crée une fenêtre de progression pour Smart Save
        /// </summary>
        public static SmartProgressWindow CreateSmartSave(int docType, string docName, string typeText)
        {
            var window = new SmartProgressWindow("💾 Smart Save", "💾");
            window._operationType = "save";
            window.SetDocumentInfo(docName, typeText);
            window.InitializeOperationsForSave(docType, typeText, docName);
            return window;
        }

        /// <summary>
        /// Crée une fenêtre de progression pour Safe Close
        /// </summary>
        public static SmartProgressWindow CreateSafeClose(int docType, string docName, string typeText)
        {
            var window = new SmartProgressWindow("🔒 Safe Close", "🔒");
            window._operationType = "close";
            window.SetDocumentInfo(docName, typeText);
            window.InitializeOperationsForClose(docType, typeText, docName);
            return window;
        }

        /// <summary>
        /// Définit les infos du document dans le panneau d'info
        /// </summary>
        private void SetDocumentInfo(string docName, string typeText)
        {
            TxtDocName.Text = docName;
            TxtDocType.Text = typeText;
            TxtDate.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " UTC";
        }

        private void InitializeOperationsForSave(int docType, string typeText, string docName)
        {
            const int kAssemblyDocumentObject = 12290;
            const int kPartDocumentObject = 12288;
            const int kDrawingDocumentObject = 12292;

            // Ajouter l'en-tête avec info document
            TxtTitle.Text = $"💾 Smart Save V1.1 - {typeText}";

            if (docType == kAssemblyDocumentObject)
            {
                AddOperation("step1", "Étape 1: 'Default' activée (POSITION-2-PRIORITAIRE)");
                AddOperation("step2", "Étape 2: Tous les composants masqués affichés");
                AddOperation("step3", "Étape 3: Réduction de l'arborescence du navigateur");
                AddOperation("step4", "Étape 4: Mise à jour du document");
                AddOperation("step5", "Étape 5: Application de la vue isométrique");
                AddOperation("step6", "Étape 6: Masquage références (Dummy, Swing, PanFactice, AirFlow, Cut_Opening)");
                AddOperation("step7", "Étape 7: Masquage des plans de référence (WorkPlanes, Axes, Points)");
                AddOperation("step8", "Étape 8: Masquage des esquisses (Sketches 2D/3D)");
                AddOperation("step9", "Étape 9: Zoom All / Fit");
                AddOperation("step10", "Étape 10: Sauvegarde du document actif");
            }
            else if (docType == kPartDocumentObject)
            {
                AddOperation("step1", "Étape 1: Activation représentation par défaut");
                AddOperation("step2", "Étape 2: Affichage des corps cachés");
                AddOperation("step3", "Étape 3: Réduction de l'arborescence du navigateur");
                AddOperation("step4", "Étape 4: Mise à jour du document");
                AddOperation("step5", "Étape 5: Application de la vue isométrique");
                AddOperation("step6", "Étape 6: Masquage des plans de référence (WorkPlanes, Axes, Points)");
                AddOperation("step7", "Étape 7: Masquage des esquisses (Sketches 2D/3D)");
                AddOperation("step8", "Étape 8: Zoom All / Fit");
                AddOperation("step9", "Étape 9: Sauvegarde du document actif");
            }
            else if (docType == kDrawingDocumentObject)
            {
                AddOperation("step1", "Étape 1: Réduction de l'arborescence du navigateur");
                AddOperation("step2", "Étape 2: Mise à jour du document et des vues");
                AddOperation("step3", "Étape 3: Zoom All / Fit");
                AddOperation("step4", "Étape 4: Sauvegarde du document actif");
            }
            else
            {
                AddOperation("step1", "Étape 1: Mise à jour du document");
                AddOperation("step2", "Étape 2: Zoom All / Fit");
                AddOperation("step3", "Étape 3: Sauvegarde du document actif");
            }
        }

        private void InitializeOperationsForClose(int docType, string typeText, string docName)
        {
            const int kAssemblyDocumentObject = 12290;
            const int kPartDocumentObject = 12288;
            const int kDrawingDocumentObject = 12292;

            TxtTitle.Text = $"🔒 Safe Close V1.7 - {typeText}";

            if (docType == kAssemblyDocumentObject)
            {
                AddOperation("step1", "Étape 1: 'Default' activée (POSITION-2-PRIORITAIRE)");
                AddOperation("step2", "Étape 2: Tous les composants masqués affichés");
                AddOperation("step3", "Étape 3: Réduction de l'arborescence du navigateur");
                AddOperation("step4", "Étape 4: Mise à jour du document");
                AddOperation("step5", "Étape 5: Application de la vue isométrique");
                AddOperation("step6", "Étape 6: Masquage références (Dummy, Swing, PanFactice, AirFlow, Cut_Opening)");
                AddOperation("step7", "Étape 7: Masquage des plans de référence (WorkPlanes, Axes, Points)");
                AddOperation("step8", "Étape 8: Masquage des esquisses (Sketches 2D/3D)");
                AddOperation("step9", "Étape 9: Zoom All / Fit");
                AddOperation("step10", "Étape 10: Sauvegarde de tous les documents ouverts");
                AddOperation("step11", "Étape 11: Fermeture du document actif");
            }
            else if (docType == kPartDocumentObject)
            {
                AddOperation("step1", "Étape 1: Activation représentation par défaut");
                AddOperation("step2", "Étape 2: Affichage des corps cachés");
                AddOperation("step3", "Étape 3: Réduction de l'arborescence du navigateur");
                AddOperation("step4", "Étape 4: Mise à jour du document");
                AddOperation("step5", "Étape 5: Application de la vue isométrique");
                AddOperation("step6", "Étape 6: Masquage des plans de référence (WorkPlanes, Axes, Points)");
                AddOperation("step7", "Étape 7: Masquage des esquisses (Sketches 2D/3D)");
                AddOperation("step8", "Étape 8: Zoom All / Fit");
                AddOperation("step9", "Étape 9: Sauvegarde de tous les documents ouverts");
                AddOperation("step10", "Étape 10: Fermeture du document actif");
            }
            else if (docType == kDrawingDocumentObject)
            {
                AddOperation("step1", "Étape 1: Réduction de l'arborescence du navigateur");
                AddOperation("step2", "Étape 2: Mise à jour du document et des vues");
                AddOperation("step3", "Étape 3: Zoom All / Fit");
                AddOperation("step4", "Étape 4: Sauvegarde de tous les documents ouverts");
                AddOperation("step5", "Étape 5: Fermeture du document actif");
            }
            else
            {
                AddOperation("step1", "Étape 1: Mise à jour du document");
                AddOperation("step2", "Étape 2: Zoom All / Fit");
                AddOperation("step3", "Étape 3: Sauvegarde de tous les documents ouverts");
                AddOperation("step4", "Étape 4: Fermeture du document actif");
            }
        }

        private void SmartProgressWindow_Loaded(object sender, RoutedEventArgs e)
        {
            CenterOnInventorWindow();
        }

        /// <summary>
        /// Ajoute une opération à la liste
        /// </summary>
        public void AddOperation(string id, string description)
        {
            Dispatcher.Invoke(() =>
            {
                var item = new OperationItem(id, description);
                _operations[id] = item;
                OperationsList.Children.Add(item.Container);
                _totalOperations++;
                UpdateProgress();
            });
        }

        /// <summary>
        /// Met à jour le statut d'une opération
        /// </summary>
        public void UpdateOperation(string id, OperationStatus status, string? message = null)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => UpdateOperationInternal(id, status, message)));
            }
            else
            {
                UpdateOperationInternal(id, status, message);
            }
        }

        private void UpdateOperationInternal(string id, OperationStatus status, string? message)
        {
            if (_operations.TryGetValue(id, out var item))
            {
                item.SetStatus(status, message);
                
                if (status == OperationStatus.Completed || status == OperationStatus.Error || status == OperationStatus.Skipped)
                {
                    _completedOperations++;
                    if (status == OperationStatus.Error) _hasErrors = true;
                }
                
                UpdateProgress();
                
                // Forcer le rafraîchissement de l'UI
                DoEvents();
            }
        }

        /// <summary>
        /// Force le rafraîchissement de l'interface utilisateur
        /// </summary>
        private void DoEvents()
        {
            try
            {
                System.Windows.Application.Current?.Dispatcher.Invoke(
                    System.Windows.Threading.DispatcherPriority.Background,
                    new Action(delegate { }));
            }
            catch { }
        }

        /// <summary>
        /// Termine toutes les opérations et prépare la fermeture
        /// </summary>
        public void Complete(bool autoClose = true)
        {
            Dispatcher.Invoke(() =>
            {
                if (_hasErrors)
                {
                    StatusIcon.Background = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // Rouge
                    StatusIconText.Text = "✗";
                    TxtStatus.Text = "Terminé avec des erreurs";
                    BtnClose.Visibility = Visibility.Visible;
                }
                else
                {
                    StatusIcon.Background = new SolidColorBrush(Color.FromRgb(39, 174, 96)); // Vert
                    StatusIconText.Text = "✓";
                    TxtStatus.Text = $"Terminé avec succès! Fermeture dans {_autoCloseDelay}s...";
                    
                    if (autoClose)
                    {
                        StartAutoCloseTimer();
                    }
                    else
                    {
                        BtnClose.Visibility = Visibility.Visible;
                    }
                }
            });
        }

        private void StartAutoCloseTimer()
        {
            int countdown = _autoCloseDelay;
            _autoCloseTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _autoCloseTimer.Tick += (s, e) =>
            {
                countdown--;
                if (countdown <= 0)
                {
                    _autoCloseTimer?.Stop();
                    this.Close();
                }
                else
                {
                    TxtStatus.Text = $"Terminé avec succès! Fermeture dans {countdown}s...";
                }
            };
            _autoCloseTimer.Start();
        }

        private void UpdateProgress()
        {
            double progress = _totalOperations > 0 ? (double)_completedOperations / _totalOperations : 0;
            double maxWidth = this.ActualWidth - 70; // Marge pour le padding
            if (maxWidth < 100) maxWidth = 500;
            
            ProgressBar.Width = progress * maxWidth;
            TxtStatus.Text = $"Progression: {_completedOperations}/{_totalOperations}";
        }

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

                    var builder = new System.Text.StringBuilder(length + 1);
                    GetWindowText(hWnd, builder, builder.Capacity);
                    string title = builder.ToString();

                    if (title.Contains("Autodesk Inventor") || title.EndsWith(".iam") || 
                        title.EndsWith(".ipt") || title.EndsWith(".idw"))
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
                    double centerX = inventorRect.Left + (inventorRect.Right - inventorRect.Left) / 2.0;
                    double centerY = inventorRect.Top + (inventorRect.Bottom - inventorRect.Top) / 2.0;
                    this.Left = centerX - (this.Width / 2.0);
                    this.Top = centerY - (this.Height / 2.0);
                }
                else
                {
                    this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                }
            }
            catch
            {
                this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                this.DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _autoCloseTimer?.Stop();
            this.Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            _autoCloseTimer?.Stop();
            this.Close();
        }
    }

    /// <summary>
    /// États possibles d'une opération
    /// </summary>
    public enum OperationStatus
    {
        Pending,    // En attente
        InProgress, // En cours
        Completed,  // Terminé avec succès
        Error,      // Erreur
        Skipped     // Ignoré
    }

    /// <summary>
    /// Représente une opération dans la liste
    /// </summary>
    public class OperationItem
    {
        public string Id { get; }
        public Border Container { get; }
        private Border _iconBorder;
        private TextBlock _iconText;
        private TextBlock _descriptionText;
        private TextBlock _messageText;

        public OperationItem(string id, string description)
        {
            Id = id;

            // Conteneur principal
            Container = new Border
            {
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(12, 10, 12, 10),
                Margin = new Thickness(0, 0, 0, 8),
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(222, 226, 230)),
                BorderThickness = new Thickness(1)
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(36) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Icône
            _iconBorder = new Border
            {
                Width = 28,
                Height = 28,
                CornerRadius = new CornerRadius(14),
                Background = new SolidColorBrush(Color.FromRgb(149, 165, 166)), // Gris (pending)
                VerticalAlignment = VerticalAlignment.Center
            };

            _iconText = new TextBlock
            {
                Text = "○",
                FontSize = 14,
                Foreground = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            _iconBorder.Child = _iconText;
            Grid.SetColumn(_iconBorder, 0);
            grid.Children.Add(_iconBorder);

            // Texte
            var textStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            _descriptionText = new TextBlock
            {
                Text = description,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(44, 62, 80))
            };
            textStack.Children.Add(_descriptionText);

            _messageText = new TextBlock
            {
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(127, 140, 141)),
                Visibility = Visibility.Collapsed
            };
            textStack.Children.Add(_messageText);

            Grid.SetColumn(textStack, 1);
            grid.Children.Add(textStack);

            Container.Child = grid;
        }

        public void SetStatus(OperationStatus status, string? message = null)
        {
            switch (status)
            {
                case OperationStatus.Pending:
                    _iconBorder.Background = new SolidColorBrush(Color.FromRgb(149, 165, 166));
                    _iconText.Text = "○";
                    Container.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                    break;

                case OperationStatus.InProgress:
                    _iconBorder.Background = new SolidColorBrush(Color.FromRgb(52, 152, 219));
                    _iconText.Text = "◌"; // Spinning indicator
                    Container.Background = new SolidColorBrush(Color.FromRgb(235, 245, 255));
                    Container.BorderBrush = new SolidColorBrush(Color.FromRgb(52, 152, 219));
                    break;

                case OperationStatus.Completed:
                    _iconBorder.Background = new SolidColorBrush(Color.FromRgb(39, 174, 96));
                    _iconText.Text = "✓";
                    Container.Background = new SolidColorBrush(Color.FromRgb(232, 245, 233));
                    Container.BorderBrush = new SolidColorBrush(Color.FromRgb(39, 174, 96));
                    break;

                case OperationStatus.Error:
                    _iconBorder.Background = new SolidColorBrush(Color.FromRgb(231, 76, 60));
                    _iconText.Text = "✗";
                    Container.Background = new SolidColorBrush(Color.FromRgb(253, 237, 237));
                    Container.BorderBrush = new SolidColorBrush(Color.FromRgb(231, 76, 60));
                    break;

                case OperationStatus.Skipped:
                    _iconBorder.Background = new SolidColorBrush(Color.FromRgb(243, 156, 18));
                    _iconText.Text = "−";
                    Container.Background = new SolidColorBrush(Color.FromRgb(255, 248, 225));
                    Container.BorderBrush = new SolidColorBrush(Color.FromRgb(243, 156, 18));
                    break;
            }

            if (!string.IsNullOrEmpty(message))
            {
                _messageText.Text = message;
                _messageText.Visibility = Visibility.Visible;
                
                // Couleur du message selon le statut
                _messageText.Foreground = status switch
                {
                    OperationStatus.Error => new SolidColorBrush(Color.FromRgb(231, 76, 60)),
                    OperationStatus.Completed => new SolidColorBrush(Color.FromRgb(39, 174, 96)),
                    _ => new SolidColorBrush(Color.FromRgb(127, 140, 141))
                };
            }
        }
    }

    /// <summary>
    /// Wrapper pour SmartProgressWindow implémentant IProgressWindow
    /// Permet d'utiliser la fenêtre WPF avec l'interface existante
    /// </summary>
    public class SmartProgressWindowWrapper : IProgressWindow
    {
        private readonly SmartProgressWindow _window;
        private static readonly Regex _stepIdRegex = new Regex(@"step(\d+)", RegexOptions.IgnoreCase);
        private string? _lastStepId = null;

        public SmartProgressWindowWrapper(SmartProgressWindow window)
        {
            _window = window;
        }

        public async Task UpdateStepStatusAsync(string stepId, string content, string statusClass)
        {
            var status = statusClass.ToLower() switch
            {
                "completed" => OperationStatus.Completed,
                "error" => OperationStatus.Error,
                "info" => OperationStatus.InProgress,
                _ => OperationStatus.Pending
            };

            // Extraire le message du contenu (enlever les emojis et "Étape X:")
            string message = ExtractMessage(content);
            
            // Si c'est une nouvelle étape et qu'on la marque comme complétée,
            // d'abord la marquer comme "en cours" pour l'effet visuel
            if (status == OperationStatus.Completed && stepId != _lastStepId)
            {
                _window.UpdateOperation(stepId, OperationStatus.InProgress, "En cours...");
                await Task.Delay(150); // Petit délai pour l'effet visuel
            }
            
            _window.UpdateOperation(stepId, status, message);
            _lastStepId = stepId;
            
            // Petit délai pour permettre à l'UI de se rafraîchir
            await Task.Delay(50);
        }

        public Task ShowCompletionAsync(string message)
        {
            bool hasError = message.Contains("❌") || message.ToLower().Contains("erreur");
            _window.Complete(!hasError);
            return Task.CompletedTask;
        }

        public void CloseWindow()
        {
            _window.Dispatcher.Invoke(() =>
            {
                try { _window.Close(); } catch { }
            });
        }

        private string ExtractMessage(string content)
        {
            // Enlever les emojis au début
            string cleaned = Regex.Replace(content, @"^[\p{So}\p{Cs}\p{Sk}✅❌⏳ℹ️🔍👁️🌲🔄📐🙈💾🛠️🎨📋📏📄🚪🧠]+\s*", "");
            
            // Enlever "Étape X:" au début
            cleaned = Regex.Replace(cleaned, @"^Étape\s*\d+\s*:\s*", "", RegexOptions.IgnoreCase);
            
            return cleaned.Trim();
        }
    }
}

