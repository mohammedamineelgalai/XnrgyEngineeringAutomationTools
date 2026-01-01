using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace XnrgyEngineeringAutomationTools.Shared.Views
{
    /// <summary>
    /// MessageBox moderne avec theme XNRGY
    /// Utilisation: XnrgyMessageBox.Show("Message", "Titre", XnrgyMessageBoxType.Success);
    /// </summary>
    public partial class XnrgyMessageBox : Window
    {
        private XnrgyMessageBoxResult _result = XnrgyMessageBoxResult.None;

        public XnrgyMessageBox()
        {
            InitializeComponent();
            MouseLeftButtonDown += (s, e) => { if (e.ChangedButton == MouseButton.Left) DragMove(); };
            LoadLogo();
        }

        private void LoadLogo()
        {
            try
            {
                string logoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "xnrgy_logo.png");
                if (System.IO.File.Exists(logoPath))
                {
                    LogoImage.Source = new BitmapImage(new Uri(logoPath, UriKind.Absolute));
                }
                else
                {
                    // Fallback: utiliser un texte
                    LogoImage.Visibility = Visibility.Collapsed;
                }
            }
            catch
            {
                LogoImage.Visibility = Visibility.Collapsed;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _result = XnrgyMessageBoxResult.Cancel;
            Close();
        }

        /// <summary>
        /// Affiche un message avec le theme XNRGY
        /// </summary>
        public static XnrgyMessageBoxResult Show(string message, string title = "XNRGY Engineering Automation", 
            XnrgyMessageBoxType type = XnrgyMessageBoxType.Info, 
            XnrgyMessageBoxButtons buttons = XnrgyMessageBoxButtons.OK,
            Window? owner = null)
        {
            var msgBox = new XnrgyMessageBox();
            
            // Configurer le titre
            msgBox.TitleText.Text = title;
            msgBox.Title = title;
            
            // Configurer le message
            msgBox.MessageText.Text = message;
            
            // Configurer l'icone selon le type
            ConfigureIcon(msgBox, type);
            
            // Configurer les boutons
            ConfigureButtons(msgBox, buttons);
            
            // Owner
            if (owner != null)
            {
                msgBox.Owner = owner;
            }
            else
            {
                msgBox.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
            
            msgBox.ShowDialog();
            return msgBox._result;
        }

        /// <summary>
        /// Raccourci pour message de succes
        /// </summary>
        public static void ShowSuccess(string message, string title = "✅ Succès", Window? owner = null)
        {
            Show(message, title, XnrgyMessageBoxType.Success, XnrgyMessageBoxButtons.OK, owner);
        }

        /// <summary>
        /// Raccourci pour message d'erreur
        /// </summary>
        public static void ShowError(string message, string title = "❌ Erreur", Window? owner = null)
        {
            Show(message, title, XnrgyMessageBoxType.Error, XnrgyMessageBoxButtons.OK, owner);
        }

        /// <summary>
        /// Raccourci pour message d'information
        /// </summary>
        public static void ShowInfo(string message, string title = "ℹ️ Information", Window? owner = null)
        {
            Show(message, title, XnrgyMessageBoxType.Info, XnrgyMessageBoxButtons.OK, owner);
        }

        /// <summary>
        /// Raccourci pour message d'avertissement
        /// </summary>
        public static void ShowWarning(string message, string title = "⚠️ Avertissement", Window? owner = null)
        {
            Show(message, title, XnrgyMessageBoxType.Warning, XnrgyMessageBoxButtons.OK, owner);
        }

        /// <summary>
        /// Raccourci pour confirmation Oui/Non
        /// </summary>
        public static bool Confirm(string message, string title = "❓ Confirmation", Window? owner = null)
        {
            return Show(message, title, XnrgyMessageBoxType.Question, XnrgyMessageBoxButtons.YesNo, owner) == XnrgyMessageBoxResult.Yes;
        }

        private static void ConfigureIcon(XnrgyMessageBox msgBox, XnrgyMessageBoxType type)
        {
            switch (type)
            {
                case XnrgyMessageBoxType.Success:
                    msgBox.IconBorder.Background = new SolidColorBrush(Color.FromRgb(16, 124, 16));
                    msgBox.IconText.Text = "✅";
                    msgBox.IconText.Foreground = Brushes.White;
                    break;
                    
                case XnrgyMessageBoxType.Error:
                    msgBox.IconBorder.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                    msgBox.IconText.Text = "❌";
                    msgBox.IconText.Foreground = Brushes.White;
                    break;
                    
                case XnrgyMessageBoxType.Warning:
                    msgBox.IconBorder.Background = new SolidColorBrush(Color.FromRgb(255, 140, 0));
                    msgBox.IconText.Text = "⚠️";
                    msgBox.IconText.Foreground = Brushes.White;
                    break;
                    
                case XnrgyMessageBoxType.Question:
                    msgBox.IconBorder.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    msgBox.IconText.Text = "❓";
                    msgBox.IconText.Foreground = Brushes.White;
                    break;
                    
                case XnrgyMessageBoxType.Info:
                default:
                    msgBox.IconBorder.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    msgBox.IconText.Text = "ℹ️";
                    msgBox.IconText.Foreground = Brushes.White;
                    break;
            }
        }

        private static void ConfigureButtons(XnrgyMessageBox msgBox, XnrgyMessageBoxButtons buttons)
        {
            msgBox.ButtonPanel.Children.Clear();
            
            switch (buttons)
            {
                case XnrgyMessageBoxButtons.OK:
                    AddButton(msgBox, "OK", XnrgyMessageBoxResult.OK, true);
                    break;
                    
                case XnrgyMessageBoxButtons.OKCancel:
                    AddButton(msgBox, "Annuler", XnrgyMessageBoxResult.Cancel, false);
                    AddButton(msgBox, "OK", XnrgyMessageBoxResult.OK, true);
                    break;
                    
                case XnrgyMessageBoxButtons.YesNo:
                    AddButton(msgBox, "Non", XnrgyMessageBoxResult.No, false);
                    AddButton(msgBox, "Oui", XnrgyMessageBoxResult.Yes, true);
                    break;
                    
                case XnrgyMessageBoxButtons.YesNoCancel:
                    AddButton(msgBox, "Annuler", XnrgyMessageBoxResult.Cancel, false);
                    AddButton(msgBox, "Non", XnrgyMessageBoxResult.No, false);
                    AddButton(msgBox, "Oui", XnrgyMessageBoxResult.Yes, true);
                    break;
            }
        }

        private static void AddButton(XnrgyMessageBox msgBox, string text, XnrgyMessageBoxResult result, bool isPrimary)
        {
            var button = new Button
            {
                Content = text,
                MinWidth = 80,
                Height = 32,
                Margin = new Thickness(8, 0, 0, 0),
                Cursor = Cursors.Hand,
                FontSize = 13
            };

            if (isPrimary)
            {
                button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                button.Foreground = Brushes.White;
                button.BorderThickness = new Thickness(0);
            }
            else
            {
                button.Background = new SolidColorBrush(Color.FromRgb(60, 60, 60));
                button.Foreground = new SolidColorBrush(Color.FromRgb(204, 204, 204));
                button.BorderBrush = new SolidColorBrush(Color.FromRgb(80, 80, 80));
                button.BorderThickness = new Thickness(1);
            }

            // Style du bouton
            var style = new Style(typeof(Button));
            style.Setters.Add(new Setter(Button.TemplateProperty, CreateButtonTemplate(isPrimary)));
            button.Style = style;

            button.Click += (s, e) =>
            {
                msgBox._result = result;
                msgBox.Close();
            };

            msgBox.ButtonPanel.Children.Add(button);
        }

        private static ControlTemplate CreateButtonTemplate(bool isPrimary)
        {
            var template = new ControlTemplate(typeof(Button));
            
            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(4));
            borderFactory.SetValue(Border.PaddingProperty, new Thickness(12, 0, 12, 0));
            borderFactory.Name = "border";
            
            var contentFactory = new FrameworkElementFactory(typeof(ContentPresenter));
            contentFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            
            borderFactory.AppendChild(contentFactory);
            template.VisualTree = borderFactory;

            // Trigger pour hover
            var hoverTrigger = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
            if (isPrimary)
            {
                hoverTrigger.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0, 100, 180))));
            }
            else
            {
                hoverTrigger.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush(Color.FromRgb(80, 80, 80))));
            }
            template.Triggers.Add(hoverTrigger);

            return template;
        }
    }

    /// <summary>
    /// Types de messages
    /// </summary>
    public enum XnrgyMessageBoxType
    {
        Info,
        Success,
        Warning,
        Error,
        Question
    }

    /// <summary>
    /// Types de boutons
    /// </summary>
    public enum XnrgyMessageBoxButtons
    {
        OK,
        OKCancel,
        YesNo,
        YesNoCancel
    }

    /// <summary>
    /// Resultat du MessageBox
    /// </summary>
    public enum XnrgyMessageBoxResult
    {
        None,
        OK,
        Cancel,
        Yes,
        No
    }
}
