using System;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using XnrgyEngineeringAutomationTools.Services;
using XnrgyEngineeringAutomationTools.Shared.Views;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Views
{
    /// <summary>
    /// Popup HTML dynamique pour afficher du contenu formaté (iProperties, rapports, etc.)
    /// Utilise WebView2 pour un rendu HTML moderne
    /// Centrage automatique sur la fenêtre Inventor
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public partial class HtmlPopupWindow : Window
    {
        private bool _isInitialized = false;
        private string _pendingHtml = string.Empty;
        
        // Callback pour les modifications de propriétés
        private Action<string, string>? _onPropertyChanged;
        private Action<string, string>? _onPropertyAdded;
        private Func<string, string, bool>? _applyPropertyChange;
        private Func<string, string, bool>? _addNewProperty;

        // Win32 API pour obtenir la position de la fenêtre Inventor
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
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public HtmlPopupWindow()
        {
            InitializeComponent();
            Loaded += HtmlPopupWindow_Loaded;
        }

        /// <summary>
        /// Constructeur avec titre et contenu HTML
        /// </summary>
        public HtmlPopupWindow(string title, string htmlContent) : this()
        {
            TxtTitle.Text = title;
            Title = title;
            _pendingHtml = htmlContent;
        }
        
        /// <summary>
        /// Définir les callbacks pour la modification de propriétés
        /// </summary>
        public void SetPropertyCallbacks(
            Func<string, string, bool>? applyPropertyChange,
            Func<string, string, bool>? addNewProperty)
        {
            _applyPropertyChange = applyPropertyChange;
            _addNewProperty = addNewProperty;
        }

        private async void HtmlPopupWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Centrer sur la fenêtre Inventor AVANT d'afficher
                CenterOnInventorWindow();

                // Initialiser WebView2
                await WebViewContent.EnsureCoreWebView2Async(null);
                _isInitialized = true;

                // Configurer la communication WebView2 -> C#
                WebViewContent.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

                // Si du HTML est en attente, l'afficher
                if (!string.IsNullOrEmpty(_pendingHtml))
                {
                    WebViewContent.NavigateToString(_pendingHtml);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] WebView2 non disponible: {ex.Message}");
                ShowFallbackContent();
            }
        }

        /// <summary>
        /// Centre la fenêtre sur la fenêtre Inventor en cherchant toutes les fenêtres
        /// </summary>
        private void CenterOnInventorWindow()
        {
            try
            {
                IntPtr inventorHandle = IntPtr.Zero;
                RECT inventorRect = new RECT();
                
                // Chercher la fenêtre Inventor parmi toutes les fenêtres
                EnumWindows((hWnd, lParam) =>
                {
                    if (!IsWindowVisible(hWnd))
                        return true; // Continuer
                    
                    int length = GetWindowTextLength(hWnd);
                    if (length == 0)
                        return true;
                    
                    var sb = new System.Text.StringBuilder(length + 1);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    string title = sb.ToString();
                    
                    // Chercher une fenêtre Inventor
                    if (title.Contains("Autodesk Inventor") || 
                        title.EndsWith(".iam") || title.EndsWith(".ipt") || 
                        title.EndsWith(".idw") || title.EndsWith(".ipn") ||
                        title.Contains("Inventor Professional") ||
                        title.Contains("Inventor 20"))
                    {
                        if (GetWindowRect(hWnd, out RECT rect))
                        {
                            // Vérifier que c'est une fenêtre de taille raisonnable (pas une tooltip)
                            int width = rect.Right - rect.Left;
                            int height = rect.Bottom - rect.Top;
                            if (width > 800 && height > 600)
                            {
                                inventorHandle = hWnd;
                                inventorRect = rect;
                                return false; // Arrêter la recherche
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);

                if (inventorHandle != IntPtr.Zero)
                {
                    // Calculer le centre de la fenêtre Inventor
                    int inventorWidth = inventorRect.Right - inventorRect.Left;
                    int inventorHeight = inventorRect.Bottom - inventorRect.Top;
                    int inventorCenterX = inventorRect.Left + (inventorWidth / 2);
                    int inventorCenterY = inventorRect.Top + (inventorHeight / 2);

                    // Positionner notre fenêtre au centre d'Inventor
                    this.Left = inventorCenterX - (this.Width / 2);
                    this.Top = inventorCenterY - (this.Height / 2);

                    // S'assurer que la fenêtre reste visible à l'écran
                    var screenWidth = SystemParameters.PrimaryScreenWidth;
                    var screenHeight = SystemParameters.PrimaryScreenHeight;
                    
                    if (this.Left < 0) this.Left = 20;
                    if (this.Top < 0) this.Top = 20;
                    if (this.Left + this.Width > screenWidth) this.Left = screenWidth - this.Width - 20;
                    if (this.Top + this.Height > screenHeight) this.Top = screenHeight - this.Height - 50;

                    System.Diagnostics.Debug.WriteLine($"[+] Fenêtre centrée sur Inventor: Left={this.Left}, Top={this.Top}");
                }
                else
                {
                    // Fallback: centrer sur l'écran principal
                    this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                    this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;
                    System.Diagnostics.Debug.WriteLine("[!] Inventor non trouvé, centrage sur écran");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Erreur centrage sur Inventor: {ex.Message}");
                this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;
            }
        }

        /// <summary>
        /// Gère les messages reçus depuis WebView2 (JavaScript -> C#)
        /// </summary>
        private void CoreWebView2_WebMessageReceived(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string message = e.TryGetWebMessageAsString();
                System.Diagnostics.Debug.WriteLine($"[WebView2] Message reçu: {message}");

                if (message == "CLOSE_FORM")
                {
                    Dispatcher.Invoke(() => Close());
                }
                else if (message.StartsWith("APPLY_PROPERTY:"))
                {
                    // Format: APPLY_PROPERTY:PropertyName|PropertyValue
                    string data = message.Substring("APPLY_PROPERTY:".Length);
                    var parts = data.Split(new[] { '|' }, 2);
                    if (parts.Length == 2)
                    {
                        string propName = parts[0];
                        string propValue = parts[1];
                        
                        bool success = false;
                        if (_applyPropertyChange != null)
                        {
                            success = _applyPropertyChange(propName, propValue);
                        }
                        
                        // Envoyer le résultat au JavaScript
                        string jsResult = success ? "true" : "false";
                        Dispatcher.Invoke(() => {
                            WebViewContent.CoreWebView2.ExecuteScriptAsync($"onPropertyApplyResult('{propName}', {jsResult})");
                        });
                    }
                }
                else if (message.StartsWith("ADD_PROPERTY:"))
                {
                    // Format: ADD_PROPERTY:PropertyName|PropertyValue
                    string data = message.Substring("ADD_PROPERTY:".Length);
                    var parts = data.Split(new[] { '|' }, 2);
                    if (parts.Length >= 1)
                    {
                        string propName = parts[0];
                        string propValue = parts.Length > 1 ? parts[1] : "";
                        
                        bool success = false;
                        if (_addNewProperty != null)
                        {
                            success = _addNewProperty(propName, propValue);
                        }
                        
                        // Envoyer le résultat au JavaScript
                        string jsResult = success ? "true" : "false";
                        Dispatcher.Invoke(() => {
                            WebViewContent.CoreWebView2.ExecuteScriptAsync($"onAddPropertyResult('{WebUtility.HtmlEncode(propName)}', '{WebUtility.HtmlEncode(propValue)}', {jsResult})");
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[!] Erreur traitement message WebView2: {ex.Message}");
            }
        }

        /// <summary>
        /// Affiche du contenu HTML dans la popup
        /// </summary>
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
        public void SetTitle(string title)
        {
            TxtTitle.Text = title;
            Title = title;
        }

        private void ShowFallbackContent()
        {
            XnrgyMessageBox.ShowWarning(
                "WebView2 n'est pas disponible. Veuillez installer le runtime Microsoft Edge WebView2.",
                "Erreur d'affichage");
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
        /// Affiche une popup HTML centrée sur Inventor
        /// </summary>
        public static void ShowHtml(string title, string htmlContent, Window? owner = null)
        {
            var popup = new HtmlPopupWindow(title, htmlContent);
            popup.ShowDialog();
        }

        /// <summary>
        /// Génère le HTML pour afficher les iProperties d'un composant
        /// Style XNRGY avec design moderne, emojis, édition inline et ajout live
        /// Migré depuis SmartToolsAmineAddin iPropertiesSummary.vb V1.3
        /// </summary>
        public static string GenerateIPropertiesHtml(string fileName, string fullPath, 
            System.Collections.Generic.Dictionary<string, string> properties)
        {
            var sb = new System.Text.StringBuilder();
            
            // Icônes emoji pour chaque propriété
            var customIcons = new System.Collections.Generic.Dictionary<string, string>
            {
                {"Prefix", "🔡"}, {"Tag", "🏷️"}, {"Tag_Assy", "🔖"}, {"Description", "📝"},
                {"Material", "🧱"}, {"DesignViewRepresentation", "🖼️"}, {"Destination", "🚚"},
                {"Length", "📐"}, {"Width", "↔️"}, {"Depth", "↕️"}, {"Thickness", "📏"},
                {"Engraving", "✒️"}, {"Finish_Paint_Face", "🎨"}, {"MachineNo", "🛠️"},
                {"Part Number", "🔢"}, {"Auteur", "👤"}, {"Entreprise", "🏢"},
                {"Derniere sauvegarde", "💾"}, {"FlatPatternLength", "📐"},
                {"FlatPatternWidth", "↔️"}, {"Liner Bend Side", "🔄"}
            };
            
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang='fr'><head><meta charset='UTF-8'>");
            sb.AppendLine("<meta http-equiv='X-UA-Compatible' content='IE=edge'>");
            sb.AppendLine("<style>");
            
            // CSS Complet avec styles modernes
            sb.AppendLine(@"
                @import url('https://fonts.googleapis.com/css2?family=Noto+Color+Emoji&display=swap');
                * { font-family: 'Segoe UI Variable','Segoe UI','Roboto','Noto Color Emoji','Apple Color Emoji',sans-serif; margin: 0; padding: 0; box-sizing: border-box; }
                body { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); margin: 0; padding: 20px; min-height: 100vh; }
                .container { background: rgba(255,255,255,0.98); border-radius: 16px; padding: 28px; box-shadow: 0 10px 40px rgba(0,0,0,0.25); max-width: 920px; min-width: 880px; margin: 0 auto; }
                h1 { color: #2c3e50; margin: 0 0 8px 0; font-size: 26px; text-align: center; font-weight: 700; letter-spacing: 1px; }
                .file-path { color: #7f8c8d; margin: 0 0 20px 0; font-size: 13px; text-align: center; font-weight: 400; overflow: hidden; text-overflow: ellipsis; max-width: 100%; padding: 12px 20px; background: #f8f9fa; border-radius: 8px; border-left: 4px solid #667eea; word-break: break-all; }
                
                /* Status message */
                .status { margin: 0 auto 18px auto; padding: 12px 20px; border-radius: 8px; text-align: center; font-weight: 600; display: none; font-size: 15px; width: fit-content; min-width: 340px; max-width: 90%; }
                .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; display: block; }
                .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; display: block; }
                .status.info { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; display: block; }
                
                /* Buttons bar */
                .buttons-bar { display: flex; align-items: center; justify-content: center; gap: 20px; margin-bottom: 24px; flex-wrap: wrap; }
                .btn-close { background: linear-gradient(45deg, #e74c3c, #c0392b); color: white; padding: 12px 32px; font-size: 15px; font-weight: 600; border: none; border-radius: 25px; cursor: pointer; transition: all 0.3s; text-transform: uppercase; letter-spacing: 0.5px; box-shadow: 0 4px 15px rgba(231,76,60,0.3); }
                .btn-close:hover { background: linear-gradient(45deg, #c0392b, #a93226); transform: translateY(-2px); box-shadow: 0 6px 20px rgba(231,76,60,0.4); }
                .btn-apply-all { background: linear-gradient(45deg, #3498db, #2980b9); color: white; padding: 12px 32px; font-size: 15px; font-weight: 600; border: none; border-radius: 25px; cursor: pointer; transition: all 0.3s; text-transform: uppercase; letter-spacing: 0.5px; box-shadow: 0 4px 15px rgba(52,152,219,0.3); }
                .btn-apply-all:hover { background: linear-gradient(45deg, #2980b9, #1a5276); transform: translateY(-2px); box-shadow: 0 6px 20px rgba(52,152,219,0.4); }
                .btn-apply-all.modified { background: linear-gradient(45deg, #27ae60, #2ecc71); animation: pulse 1.5s infinite; box-shadow: 0 4px 15px rgba(39,174,96,0.4); }
                @keyframes pulse { 0%, 100% { box-shadow: 0 0 8px 2px rgba(39,174,96,0.4); } 50% { box-shadow: 0 0 20px 6px rgba(39,174,96,0.7); } }
                
                /* Properties list */
                .properties-list { display: flex; flex-direction: column; gap: 12px; max-height: 450px; overflow-y: auto; padding-right: 10px; }
                .property-row { display: flex; align-items: center; gap: 14px; width: 100%; }
                .property-oblong { display: flex; align-items: center; justify-content: flex-start; width: 300px; min-width: 300px; max-width: 300px; height: 48px; border-radius: 24px; background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border: 2px solid #dee2e6; font-size: 14px; font-weight: 600; box-shadow: 0 2px 8px rgba(0,0,0,0.05); padding-left: 18px; overflow: hidden; white-space: nowrap; transition: all 0.2s; }
                .property-oblong:hover { border-color: #667eea; box-shadow: 0 4px 12px rgba(102,126,234,0.15); }
                .property-value-oblong { flex: 1; height: 48px; display: flex; align-items: center; background: #ffffff; border-radius: 24px; border: 2px solid #dee2e6; box-shadow: 0 2px 8px rgba(0,0,0,0.05); padding: 0 18px; font-size: 14px; overflow: hidden; transition: all 0.2s; }
                .property-value-oblong:hover { border-color: #667eea; }
                .property-value-oblong:focus-within { border-color: #667eea; box-shadow: 0 0 0 3px rgba(102,126,234,0.2); }
                .property-label { font-weight: 600; color: #2c3e50; text-overflow: ellipsis; overflow: hidden; }
                .property-icon { margin-right: 10px; font-size: 18px; }
                
                /* Input styling */
                input[type='text'] { width: 100%; border: none; outline: none; background: transparent; font-size: 14px; font-family: inherit; font-weight: 500; color: #34495e; }
                input[type='text'].modified { color: #27ae60; font-weight: 600; }
                input[type='text']::placeholder { color: #adb5bd; font-style: italic; }
                
                /* Ligne d'ajout */
                .add-section { margin-top: 20px; padding-top: 20px; border-top: 2px dashed #dee2e6; }
                .add-section-title { text-align: center; font-size: 14px; font-weight: 600; color: #27ae60; margin-bottom: 12px; }
                .add-property-row { display: flex; align-items: center; gap: 14px; width: 100%; }
                .add-property-row .property-oblong { background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); border-color: #81c784; }
                .add-property-row .property-value-oblong { background: #f1f8e9; border-color: #81c784; }
                .add-property-row input::placeholder { color: #66bb6a; }
                
                /* Bouton Ajouter centré en bas */
                .add-button-container { display: flex; justify-content: center; margin-top: 16px; }
                .btn-add { background: linear-gradient(45deg, #27ae60, #2ecc71); color: white; padding: 14px 40px; font-size: 16px; font-weight: 600; border: none; border-radius: 30px; cursor: pointer; transition: all 0.3s; text-transform: uppercase; letter-spacing: 0.5px; box-shadow: 0 4px 15px rgba(39,174,96,0.3); }
                .btn-add:hover:not(:disabled) { background: linear-gradient(45deg, #229954, #27ae60); transform: translateY(-2px); box-shadow: 0 6px 20px rgba(39,174,96,0.4); }
                .btn-add:disabled { background: #bdc3c7; cursor: not-allowed; transform: none; box-shadow: none; }
                
                /* Icon colors */
                .icon-prefix { color: #e67e22; } .icon-tag { color: #16a085; } .icon-tagassy { color: #9b59b6; }
                .icon-description { color: #2c3e50; } .icon-material { color: #c0392b; } .icon-dimension { color: #2980b9; }
                .icon-destination { color: #7f8c8d; } .icon-engraving { color: #34495e; } .icon-paint { color: #8e44ad; }
                .icon-machine { color: #27ae60; } .icon-design { color: #1abc9c; } .icon-default { color: #3498db; }
                .icon-new { color: #27ae60; }
                
                .footer { text-align: center; font-size: 12px; color: #95a5a6; margin-top: 20px; padding-top: 14px; border-top: 1px solid #ecf0f1; }
                
                /* Animation pour nouvelles lignes */
                @keyframes slideIn { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
                .property-row.new { animation: slideIn 0.3s ease-out; }
            ");
            sb.AppendLine("</style></head><body>");
            
            sb.AppendLine("<div class='container'>");
            
            // Titre
            string displayName = fileName.Length > 45 ? fileName.Substring(0, 42) + "..." : fileName;
            sb.AppendLine($"<h1>📋 {WebUtility.HtmlEncode(displayName)}</h1>");
            sb.AppendLine($"<div class='file-path'>🔍 {WebUtility.HtmlEncode(fullPath)}</div>");
            
            // Message de status
            sb.AppendLine("<div id='status' class='status'></div>");
            
            // Boutons principaux : Fermer et Appliquer
            sb.AppendLine("<div class='buttons-bar'>");
            sb.AppendLine("<button type='button' id='closeBtn' class='btn-close'>✖️ Fermer</button>");
            sb.AppendLine("<button type='button' id='applyAllBtn' class='btn-apply-all'>💾 Appliquer</button>");
            sb.AppendLine("</div>");
            
            // Liste des propriétés éditables
            sb.AppendLine("<div class='properties-list' id='propertiesList'>");
            
            foreach (var prop in properties)
            {
                string icon = customIcons.ContainsKey(prop.Key) ? customIcons[prop.Key] : "🔸";
                string iconClass = GetIconClass(prop.Key);
                string value = prop.Value ?? "";
                string escapedKey = WebUtility.HtmlEncode(prop.Key).Replace("'", "\\'");
                
                sb.AppendLine($"<div class='property-row' data-prop='{escapedKey}'>");
                sb.AppendLine($"  <div class='property-oblong'>");
                sb.AppendLine($"    <span class='property-icon {iconClass}'>{icon}</span>");
                sb.AppendLine($"    <span class='property-label'>{WebUtility.HtmlEncode(prop.Key)}</span>");
                sb.AppendLine($"  </div>");
                sb.AppendLine($"  <div class='property-value-oblong'>");
                sb.AppendLine($"    <input type='text' data-key='{escapedKey}' value=\"{WebUtility.HtmlEncode(value)}\" data-original=\"{WebUtility.HtmlEncode(value)}\" placeholder='[Non défini]'/>");
                sb.AppendLine($"  </div>");
                sb.AppendLine("</div>");
            }
            
            sb.AppendLine("</div>"); // Fin properties-list
            
            // Section d'ajout de nouvelle propriété
            sb.AppendLine("<div class='add-section'>");
            sb.AppendLine("<div class='add-section-title'>➕ Ajouter une nouvelle propriété</div>");
            sb.AppendLine("<div class='add-property-row'>");
            sb.AppendLine("  <div class='property-oblong'>");
            sb.AppendLine("    <span class='property-icon icon-new'>🆕</span>");
            sb.AppendLine("    <input type='text' id='newPropertyName' placeholder='Nom de la propriété...' style='width:200px; font-weight:600;'/>");
            sb.AppendLine("  </div>");
            sb.AppendLine("  <div class='property-value-oblong'>");
            sb.AppendLine("    <input type='text' id='newPropertyValue' placeholder='Valeur initiale...'/>");
            sb.AppendLine("  </div>");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class='add-button-container'>");
            sb.AppendLine("  <button type='button' id='addPropertyBtn' class='btn-add' disabled>➕ AJOUTER LA PROPRIÉTÉ</button>");
            sb.AppendLine("</div>");
            sb.AppendLine("</div>");
            
            // Footer
            sb.AppendLine("<div class='footer'>🏷️ Smart iProperties Summary V1.4 - By Mohammed Amine Elgalai - XNRGY Climate Systems ULC</div>");
            sb.AppendLine("</div>");
            
            // JavaScript
            sb.AppendLine(@"
                <script>
                var originalValues = {};
                var pendingChanges = [];
                var applyAllBtn;
                var newNameInput;
                var newValueInput;
                var addBtn;
                
                function showStatus(message, type) {
                    var status = document.getElementById('status');
                    status.textContent = message;
                    status.className = 'status ' + type;
                    if (type !== 'info') {
                        setTimeout(function() { status.className = 'status'; }, 4000);
                    }
                }
                
                function hideStatus() {
                    document.getElementById('status').className = 'status';
                }
                
                // Initialiser les valeurs originales
                function initValues() {
                    document.querySelectorAll('input[data-key]').forEach(function(input) {
                        originalValues[input.dataset.key] = input.dataset.original || '';
                    });
                }
                
                // Vérifier s'il y a des modifications
                function checkForChanges() {
                    if (!applyAllBtn) return false;
                    var hasChanges = false;
                    document.querySelectorAll('input[data-key]').forEach(function(input) {
                        var key = input.dataset.key;
                        var originalValue = originalValues[key] || '';
                        if (input.value !== originalValue) {
                            hasChanges = true;
                            input.classList.add('modified');
                        } else {
                            input.classList.remove('modified');
                        }
                    });
                    
                    if (hasChanges) {
                        applyAllBtn.classList.add('modified');
                    } else {
                        applyAllBtn.classList.remove('modified');
                    }
                    return hasChanges;
                }
                
                // Bouton Appliquer global - applique toutes les modifications
                function handleApplyAll() {
                    var modifiedInputs = document.querySelectorAll('input[data-key].modified');
                    if (modifiedInputs.length === 0) {
                        showStatus('ℹ️ Aucune modification à appliquer', 'info');
                        return;
                    }
                    
                    pendingChanges = [];
                    modifiedInputs.forEach(function(input) {
                        pendingChanges.push({ key: input.dataset.key, value: input.value });
                    });
                    
                    showStatus('⏳ Application de ' + pendingChanges.length + ' modification(s)...', 'info');
                    applyNextChange();
                }
                
                function applyNextChange() {
                    if (pendingChanges.length === 0) {
                        showStatus('✅ Toutes les modifications ont été appliquées!', 'success');
                        checkForChanges();
                        return;
                    }
                    var change = pendingChanges.shift();
                    window.chrome.webview.postMessage('APPLY_PROPERTY:' + change.key + '|' + change.value);
                }
                
                // Callback du résultat d'application
                function onPropertyApplyResult(propName, success) {
                    var input = document.querySelector('input[data-key=""' + propName + '""]');
                    
                    if (success) {
                        if (input) {
                            originalValues[propName] = input.value;
                            input.dataset.original = input.value;
                            input.classList.remove('modified');
                        }
                        // Continuer avec la prochaine modification
                        applyNextChange();
                    } else {
                        showStatus('❌ Erreur lors de la mise à jour de ""' + propName + '""', 'error');
                        pendingChanges = []; // Arrêter en cas d'erreur
                        checkForChanges();
                    }
                }
                
                // === SECTION AJOUT DE PROPRIÉTÉ ===
                function updateAddBtnState() {
                    if (!newNameInput || !addBtn) return;
                    var name = newNameInput.value.trim();
                    addBtn.disabled = (name.length === 0);
                }
                
                function setupAddPropertySection() {
                    newNameInput = document.getElementById('newPropertyName');
                    newValueInput = document.getElementById('newPropertyValue');
                    addBtn = document.getElementById('addPropertyBtn');
                    
                    if (!newNameInput || !addBtn) {
                        console.error('Elements not found!');
                        return;
                    }
                    
                    // Événements pour activer/désactiver le bouton
                    newNameInput.oninput = updateAddBtnState;
                    newNameInput.onkeyup = updateAddBtnState;
                    newNameInput.onchange = updateAddBtnState;
                    newNameInput.onkeypress = function(e) {
                        if (e.key === 'Enter' && newNameInput.value.trim()) {
                            newValueInput.focus();
                        }
                    };
                    
                    if (newValueInput) {
                        newValueInput.onkeypress = function(e) {
                            if (e.key === 'Enter' && newNameInput.value.trim()) {
                                addBtn.click();
                            }
                        };
                    }
                    
                    // Click sur le bouton Ajouter
                    addBtn.onclick = function() {
                        var name = newNameInput.value.trim();
                        var value = newValueInput.value;
                        
                        if (!name) {
                            showStatus('❌ Le nom de la propriété est requis!', 'error');
                            return;
                        }
                        
                        if (originalValues.hasOwnProperty(name)) {
                            showStatus('❌ Cette propriété existe déjà!', 'error');
                            return;
                        }
                        
                        showStatus('⏳ Ajout de la propriété...', 'info');
                        window.chrome.webview.postMessage('ADD_PROPERTY:' + name + '|' + value);
                    };
                }
                
                // Callback du résultat d'ajout
                function onAddPropertyResult(propName, propValue, success) {
                    if (success) {
                        showStatus('✅ Propriété ""' + propName + '"" ajoutée avec succès!', 'success');
                        
                        // Ajouter la nouvelle ligne dans la liste
                        var list = document.getElementById('propertiesList');
                        
                        var newRow = document.createElement('div');
                        newRow.className = 'property-row new';
                        newRow.dataset.prop = propName;
                        newRow.innerHTML = 
                            ""<div class='property-oblong'>"" +
                            ""<span class='property-icon icon-new'>✨</span>"" +
                            ""<span class='property-label'>"" + propName + ""</span>"" +
                            ""</div>"" +
                            ""<div class='property-value-oblong'>"" +
                            ""<input type='text' data-key='"" + propName + ""' value='"" + propValue + ""' data-original='"" + propValue + ""' placeholder='[Non défini]'/>"" +
                            ""</div>"";
                        
                        list.appendChild(newRow);
                        
                        // Enregistrer la valeur originale
                        originalValues[propName] = propValue;
                        
                        // Attacher les événements à la nouvelle ligne
                        var newInput = newRow.querySelector('input[data-key]');
                        if (newInput) {
                            newInput.addEventListener('input', function(e) {
                                checkForChanges();
                            });
                        }
                        
                        // Vider les champs d'ajout
                        newNameInput.value = '';
                        newValueInput.value = '';
                        addBtn.disabled = true;
                        
                        // Scroller vers la nouvelle ligne
                        newRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    } else {
                        showStatus('❌ Erreur lors de l\\'ajout de la propriété', 'error');
                    }
                }
                
                // Initialisation au chargement de la page
                window.onload = function() {
                    applyAllBtn = document.getElementById('applyAllBtn');
                    initValues();
                    setupAddPropertySection();
                    
                    // Attacher événements aux inputs existants
                    document.querySelectorAll('input[data-key]').forEach(function(input) {
                        input.oninput = function(e) { checkForChanges(); };
                    });
                    
                    // Bouton Appliquer global
                    if (applyAllBtn) {
                        applyAllBtn.onclick = handleApplyAll;
                    }
                    
                    // Bouton Fermer
                    document.getElementById('closeBtn').onclick = function() {
                        window.chrome.webview.postMessage('CLOSE_FORM');
                    };
                };
                </script>
            ");
            
            sb.AppendLine("</body></html>");
            
            return sb.ToString();
        }

        /// <summary>
        /// Retourne la classe CSS pour la couleur de l'icône selon le nom de propriété
        /// </summary>
        private static string GetIconClass(string propertyName)
        {
            return propertyName switch
            {
                "Prefix" => "icon-prefix",
                "Tag" => "icon-tag",
                "Tag_Assy" => "icon-tagassy",
                "Description" => "icon-description",
                "Material" => "icon-material",
                "Length" or "Width" or "Depth" or "Thickness" or "FlatPatternLength" or "FlatPatternWidth" => "icon-dimension",
                "Destination" => "icon-destination",
                "Engraving" => "icon-engraving",
                "Finish_Paint_Face" => "icon-paint",
                "MachineNo" => "icon-machine",
                "DesignViewRepresentation" => "icon-design",
                _ => "icon-default"
            };
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
