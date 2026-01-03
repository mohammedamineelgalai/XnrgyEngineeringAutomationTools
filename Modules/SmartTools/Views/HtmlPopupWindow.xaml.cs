using System;
using System.Net;
using System.Windows;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Popup HTML dynamique pour afficher du contenu formaté (iProperties, rapports, etc.)
    /// Utilise WebView2 pour un rendu HTML moderne
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public partial class HtmlPopupWindow : Window
    {
        private bool _isInitialized = false;
        private string _pendingHtml = string.Empty;

        public HtmlPopupWindow()
        {
            InitializeComponent();
            Loaded += HtmlPopupWindow_Loaded;
        }

        /// <summary>
        /// Constructeur avec titre et contenu HTML
        /// </summary>
        /// <param name="title">Titre de la fenêtre</param>
        /// <param name="htmlContent">Contenu HTML à afficher</param>
        public HtmlPopupWindow(string title, string htmlContent) : this()
        {
            TxtTitle.Text = title;
            Title = title;
            _pendingHtml = htmlContent;
        }

        private async void HtmlPopupWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Initialiser WebView2
                await WebViewContent.EnsureCoreWebView2Async(null);
                _isInitialized = true;

                // Si du HTML est en attente, l'afficher
                if (!string.IsNullOrEmpty(_pendingHtml))
                {
                    WebViewContent.NavigateToString(_pendingHtml);
                }
            }
            catch (Exception ex)
            {
                // Fallback: afficher dans un TextBlock si WebView2 n'est pas disponible
                System.Diagnostics.Debug.WriteLine($"[!] WebView2 non disponible: {ex.Message}");
                ShowFallbackContent();
            }
        }

        /// <summary>
        /// Affiche du contenu HTML dans la popup
        /// </summary>
        /// <param name="htmlContent">Contenu HTML complet</param>
        public void SetHtmlContent(string htmlContent)
        {
            if (_isInitialized)
            {
                WebViewContent.NavigateToString(htmlContent);
            }
            else
            {
                _pendingHtml = htmlContent;
            }
        }

        /// <summary>
        /// Définit le titre de la fenêtre
        /// </summary>
        /// <param name="title">Nouveau titre</param>
        public void SetTitle(string title)
        {
            TxtTitle.Text = title;
            Title = title;
        }

        private void ShowFallbackContent()
        {
            // Si WebView2 n'est pas disponible, on peut utiliser un navigateur WinForms
            // ou afficher un message d'erreur
            MessageBox.Show(
                "WebView2 n'est pas disponible. Veuillez installer le runtime Microsoft Edge WebView2.",
                "Erreur d'affichage",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// Affiche une popup HTML avec le titre et contenu spécifiés
        /// </summary>
        /// <param name="title">Titre de la fenêtre</param>
        /// <param name="htmlContent">Contenu HTML</param>
        /// <param name="owner">Fenêtre parente (optionnel)</param>
        public static void ShowHtml(string title, string htmlContent, Window? owner = null)
        {
            var popup = new HtmlPopupWindow(title, htmlContent);
            if (owner != null)
            {
                popup.Owner = owner;
                popup.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            popup.ShowDialog();
        }

        /// <summary>
        /// Génère le HTML pour afficher les iProperties d'un composant
        /// Style XNRGY avec design moderne
        /// </summary>
        /// <param name="fileName">Nom du fichier</param>
        /// <param name="fullPath">Chemin complet</param>
        /// <param name="properties">Dictionnaire des propriétés (clé, valeur)</param>
        /// <returns>Contenu HTML formaté</returns>
        public static string GenerateIPropertiesHtml(string fileName, string fullPath, 
            System.Collections.Generic.Dictionary<string, string> properties)
        {
            var sb = new System.Text.StringBuilder();
            
            // Header HTML avec styles XNRGY
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html><head><meta charset='UTF-8'>");
            sb.AppendLine("<style>");
            sb.AppendLine("* { margin: 0; padding: 0; box-sizing: border-box; }");
            sb.AppendLine("body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #1E1E2E; color: #E0E0E0; padding: 16px; }");
            sb.AppendLine("h2 { color: #0078D4; font-weight: bold; margin-bottom: 16px; font-size: 18px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }");
            sb.AppendLine("h2 span { color: #FFB74D; }");
            sb.AppendLine(".path { color: #888888; font-size: 11px; margin-bottom: 16px; word-break: break-all; background: #252536; padding: 8px; border-radius: 4px; border-left: 3px solid #0078D4; }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; background-color: #252536; border-radius: 6px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }");
            sb.AppendLine("th { background-color: #2A4A6F; color: #FFFFFF; padding: 12px 16px; text-align: left; font-weight: 600; font-size: 13px; }");
            sb.AppendLine("td { padding: 10px 16px; border-bottom: 1px solid #3D3D56; font-size: 13px; }");
            sb.AppendLine("tr:last-child td { border-bottom: none; }");
            sb.AppendLine("tr:hover { background-color: #2D2D44; }");
            sb.AppendLine(".prop-name { color: #0078D4; font-weight: 500; }");
            sb.AppendLine(".prop-value { color: #E0E0E0; }");
            sb.AppendLine(".icon { margin-right: 8px; }");
            sb.AppendLine(".undefined { color: #888888; font-style: italic; }");
            sb.AppendLine("</style></head><body>");
            
            // Titre avec nom du fichier (tronqué si trop long)
            string displayName = fileName.Length > 40 ? fileName.Substring(0, 37) + "..." : fileName;
            sb.AppendLine($"<h2><span>[i]</span> iProperties Summary - {WebUtility.HtmlEncode(displayName)}</h2>");
            
            // Chemin complet
            sb.AppendLine($"<div class='path'><strong>Chemin:</strong> {WebUtility.HtmlEncode(fullPath)}</div>");
            
            // Tableau des propriétés
            sb.AppendLine("<table>");
            sb.AppendLine("<tr><th><span class='icon'>[>]</span> Propriété</th><th><span class='icon'>[i]</span> Valeur</th></tr>");
            
            // Icônes pour chaque propriété
            var icons = new System.Collections.Generic.Dictionary<string, string>
            {
                { "Tag", "[#]" },
                { "Tag_Assy", "[#]" },
                { "Description", "[i]" },
                { "Material", "[*]" },
                { "Length", "[>]" },
                { "Width", "[>]" },
                { "Depth", "[>]" },
                { "Thickness", "[>]" },
                { "Modifié par", "[~]" },
                { "Last Saved By", "[~]" }
            };
            
            foreach (var prop in properties)
            {
                string icon = icons.ContainsKey(prop.Key) ? icons[prop.Key] : "[+]";
                string valueClass = (prop.Value == "[Non défini]" || string.IsNullOrEmpty(prop.Value)) ? "undefined" : "prop-value";
                string value = string.IsNullOrEmpty(prop.Value) ? "[Non défini]" : prop.Value;
                
                sb.AppendLine($"<tr><td class='prop-name'><span class='icon'>{icon}</span> {WebUtility.HtmlEncode(prop.Key)}</td>");
                sb.AppendLine($"<td class='{valueClass}'>{WebUtility.HtmlEncode(value)}</td></tr>");
            }
            
            sb.AppendLine("</table>");
            sb.AppendLine("</body></html>");
            
            return sb.ToString();
        }

        /// <summary>
        /// Génère le HTML pour afficher un rapport de contraintes d'assemblage
        /// </summary>
        public static string GenerateConstraintReportHtml(string assemblyName, 
            System.Collections.Generic.List<ConstraintInfo> constraints)
        {
            var sb = new System.Text.StringBuilder();
            
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html><head><meta charset='UTF-8'>");
            sb.AppendLine("<style>");
            sb.AppendLine("* { margin: 0; padding: 0; box-sizing: border-box; }");
            sb.AppendLine("body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #1E1E2E; color: #E0E0E0; padding: 16px; }");
            sb.AppendLine("h2 { color: #0078D4; margin-bottom: 16px; }");
            sb.AppendLine(".summary { background: #252536; padding: 12px; border-radius: 6px; margin-bottom: 16px; border-left: 3px solid #107C10; }");
            sb.AppendLine(".summary.warning { border-left-color: #FF8C00; }");
            sb.AppendLine(".summary.error { border-left-color: #E81123; }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; background-color: #252536; border-radius: 6px; overflow: hidden; }");
            sb.AppendLine("th { background-color: #2A4A6F; color: #FFFFFF; padding: 10px; text-align: left; }");
            sb.AppendLine("td { padding: 8px 10px; border-bottom: 1px solid #3D3D56; font-size: 12px; }");
            sb.AppendLine("tr:hover { background-color: #2D2D44; }");
            sb.AppendLine(".status-ok { color: #107C10; }");
            sb.AppendLine(".status-warning { color: #FF8C00; }");
            sb.AppendLine(".status-error { color: #E81123; }");
            sb.AppendLine("</style></head><body>");
            
            sb.AppendLine($"<h2>[>] Constraint Report - {WebUtility.HtmlEncode(assemblyName)}</h2>");
            
            // Compteurs
            int total = constraints.Count;
            int suppressed = constraints.FindAll(c => c.IsSuppressed).Count;
            int failed = constraints.FindAll(c => !c.IsHealthy).Count;
            
            string summaryClass = failed > 0 ? "error" : (suppressed > 0 ? "warning" : "");
            sb.AppendLine($"<div class='summary {summaryClass}'>");
            sb.AppendLine($"<strong>Total:</strong> {total} contraintes | ");
            sb.AppendLine($"<span class='status-ok'>OK: {total - failed - suppressed}</span> | ");
            sb.AppendLine($"<span class='status-warning'>Supprimées: {suppressed}</span> | ");
            sb.AppendLine($"<span class='status-error'>Erreurs: {failed}</span>");
            sb.AppendLine("</div>");
            
            // Tableau
            sb.AppendLine("<table>");
            sb.AppendLine("<tr><th>Composant 1</th><th>Composant 2</th><th>Type</th><th>Status</th></tr>");
            
            foreach (var c in constraints)
            {
                string statusClass = c.IsSuppressed ? "status-warning" : (c.IsHealthy ? "status-ok" : "status-error");
                string status = c.IsSuppressed ? "Supprimée" : (c.IsHealthy ? "OK" : "Erreur");
                
                sb.AppendLine($"<tr>");
                sb.AppendLine($"<td>{WebUtility.HtmlEncode(c.Component1)}</td>");
                sb.AppendLine($"<td>{WebUtility.HtmlEncode(c.Component2)}</td>");
                sb.AppendLine($"<td>{WebUtility.HtmlEncode(c.ConstraintType)}</td>");
                sb.AppendLine($"<td class='{statusClass}'>{status}</td>");
                sb.AppendLine($"</tr>");
            }
            
            sb.AppendLine("</table>");
            sb.AppendLine("</body></html>");
            
            return sb.ToString();
        }
    }

    /// <summary>
    /// Information sur une contrainte d'assemblage
    /// </summary>
    public class ConstraintInfo
    {
        public string Component1 { get; set; } = "";
        public string Component2 { get; set; } = "";
        public string ConstraintType { get; set; } = "";
        public bool IsHealthy { get; set; } = true;
        public bool IsSuppressed { get; set; } = false;
    }
}
