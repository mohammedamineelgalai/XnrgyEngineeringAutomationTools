using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace XnrgyEngineeringAutomationTools.Modules.SmartTools.Services
{
    /// <summary>
    /// Service pour générer tous les HTML utilisés dans SmartTools
    /// Converti depuis SmartToolsAmineAddin
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public static class HtmlGeneratorService
    {
        /// <summary>
        /// Génère le HTML pour iPropertiesSummary avec formulaire d'édition complet
        /// </summary>
        public static string GenerateIPropertiesSummaryHtml(string fileName, string filePath, 
            Dictionary<string, string> propValues, List<PropertyMeta> keys)
        {
            var html = new StringBuilder();
            int propertyOblongWidth = 400;
            int valueOblongWidth = 420;

            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html><head><meta charset='UTF-8'>");
            html.AppendLine("<meta http-equiv='X-UA-Compatible' content='IE=edge'>");
            html.AppendLine("<style>");
            html.AppendLine("body { font-family: 'Segoe UI Variable','Segoe UI',Arial,sans-serif; background: linear-gradient(135deg,#667eea 0%,#764ba2 100%); margin:0; padding:15px; min-height:100vh; }");
            html.AppendLine($".container {{ background:rgba(255,255,255,0.98); border-radius:13px; padding:22px 10px 18px 10px; box-shadow:0 8px 28px rgba(0,0,0,0.13); max-width:880px; min-width:880px; min-height:1000px; margin:auto; }}");
            html.AppendLine("h1 { color:#2c3e50; margin:0 0 5px 0; font-size:25px; text-align:center; font-weight:700; letter-spacing:1px; }");
            html.AppendLine(".file-path { color:#7f8c8d; margin:0 0 18px 0; font-size:14px; text-align:center; font-weight:400; overflow:hidden; text-overflow:ellipsis; max-width:100%; padding:0 20px; }");
            html.AppendLine(".status { margin: 0 auto 18px auto; padding:9px 18px; border-radius:7px; text-align:center; font-weight:600; display:none; font-size:15px; width:fit-content; min-width:340px; max-width:90%; position:relative; top:-8px; z-index:2; }");
            html.AppendLine(".status.success { background:#d4edda; color:#155724; border:1px solid #c3e6cb; }");
            html.AppendLine(".status.error { background:#f8d7da; color:#721c24; border:1px solid #f5c6cb; }");
            html.AppendLine(".buttons-bar { display:flex; align-items:center; justify-content:center; gap:20px; margin-bottom:18px; }");
            html.AppendLine(".btn-apply { background:linear-gradient(45deg,#2186eb,#037cd4); color:white; box-shadow:0 2px 8px rgba(52,152,219,0.08); padding:11px 30px; font-size:15px; font-weight:600; border:none; border-radius:22px; cursor:pointer; transition:all 0.2s; text-transform:uppercase; letter-spacing:0.2px; height:42px; min-width:120px; }");
            html.AppendLine(".btn-close { background:linear-gradient(45deg,#e74c3c,#c0392b); color:white; padding:11px 30px; font-size:15px; font-weight:600; border:none; border-radius:22px; cursor:pointer; transition:all 0.2s; text-transform:uppercase; letter-spacing:0.2px; height:42px; min-width:120px; }");
            html.AppendLine(".btn-add { background:linear-gradient(45deg,#27ae60,#2ecc71); color:white; padding:11px 30px; font-size:15px; font-weight:600; border:none; border-radius:22px; cursor:pointer; transition:all 0.2s; text-transform:uppercase; letter-spacing:0.2px; height:42px; min-width:140px; box-shadow:0 2px 8px rgba(39,174,96,0.08); }");
            html.AppendLine(".btn-add:hover { background:linear-gradient(45deg,#229954,#27ae60); transform: translateY(-1px); box-shadow:0 4px 12px rgba(39,174,96,0.15); }");
            html.AppendLine(".btn-apply.modified { background:linear-gradient(45deg,#2ecc71,#27ae60); box-shadow:0 0 8px 1px #27ae6044, 0 4px 12px rgba(39,174,96,0.09); animation:pulse 1s infinite; }");
            html.AppendLine("@keyframes pulse { 0% { box-shadow:0 0 8px 1px #27ae6044; } 50% { box-shadow:0 0 16px 4px #27ae6099; } 100% { box-shadow:0 0 8px 1px #27ae6044; } }");
            html.AppendLine(".properties-list { display:flex; flex-direction:column; gap:14px; margin-bottom:18px; }");
            html.AppendLine(".property-row { display:flex; align-items:center; gap:18px; width:100%; }");
            html.AppendLine($".property-oblong {{ display:flex; align-items:center; justify-content:flex-start; width:{propertyOblongWidth}px; min-width:{propertyOblongWidth}px; max-width:{propertyOblongWidth}px; height:46px; border-radius:28px; background:#f7fafd; border:1.4px solid #bfc7d3; font-size:16px; font-weight:600; box-shadow:0 1px 5px rgba(52,152,219,0.04); letter-spacing:0.3px; padding-left:17px; overflow:hidden; white-space:nowrap; }}");
            html.AppendLine($".property-value-oblong {{ flex:1; width:{valueOblongWidth}px; min-width:{valueOblongWidth}px; max-width:{valueOblongWidth}px; height:46px; display:flex; align-items:center; background:#fff; border-radius:28px; border:1.4px solid #bfc7d3; box-shadow:0 1px 4px rgba(52,152,219,0.04); padding:0 16px; font-size:16px; overflow-x:auto; overflow-y:hidden; }}");
            html.AppendLine(".property-label { font-weight:600; color:#34495e; font-family:inherit; font-size:inherit; }");
            html.AppendLine(".property-icon { margin-right:8px; font-size:21px; font-weight:700; vertical-align:middle; }");
            html.AppendLine("input[type='text'] { width:100%; border:none; outline:none; background:transparent; font-size:16px; font-family:inherit; font-weight:500; color:#222; padding:0; text-overflow:ellipsis; overflow:visible; white-space:nowrap; transition:font-size 0.2s; }");
            html.AppendLine("input[type='text'].modified { color:#27ae60; background:transparent; }");
            html.AppendLine(".modal-overlay { position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:1000; display:none; }");
            html.AppendLine(".modal-content { position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); background:white; border-radius:15px; padding:30px; box-shadow:0 10px 30px rgba(0,0,0,0.3); min-width:500px; max-width:600px; }");
            html.AppendLine(".form-group { margin-bottom:20px; }");
            html.AppendLine(".form-label { display:block; font-weight:600; color:#34495e; margin-bottom:8px; font-size:16px; }");
            html.AppendLine(".form-input { width:100%; padding:12px 16px; border:2px solid #bfc7d3; border-radius:8px; font-size:16px; font-family:inherit; box-sizing:border-box; }");
            html.AppendLine(".form-select { width:100%; padding:12px 16px; border:2px solid #bfc7d3; border-radius:8px; font-size:16px; font-family:inherit; background:white; box-sizing:border-box; }");
            html.AppendLine(".modal-buttons { display:flex; gap:15px; justify-content:center; margin-top:25px; }");
            html.AppendLine(".btn-modal-confirm { background:linear-gradient(45deg,#27ae60,#2ecc71); color:white; padding:12px 25px; border:none; border-radius:8px; font-size:16px; font-weight:600; cursor:pointer; }");
            html.AppendLine(".btn-modal-cancel { background:linear-gradient(45deg,#95a5a6,#7f8c8d); color:white; padding:12px 25px; border:none; border-radius:8px; font-size:16px; font-weight:600; cursor:pointer; }");
            html.AppendLine(".footer { text-align:center; font-size:14px; color:#888; margin-top:20px; }");
            html.AppendLine("</style></head><body>");
            html.AppendLine("<div class='container'>");
            html.AppendLine($"<h1>📋 {WebUtility.HtmlEncode(fileName)}</h1>");
            html.AppendLine($"<div class='file-path'>🔍 {WebUtility.HtmlEncode(filePath)}</div>");
            html.AppendLine("<div id='status' class='status'></div>");
            html.AppendLine("<div class='buttons-bar'>");
            html.AppendLine("<button type='button' id='closeBtn' class='btn-close'>Close</button>");
            html.AppendLine("<button type='button' id='addBtn' class='btn-add'>➕ Ajouter</button>");
            html.AppendLine("<button type='button' id='applyBtn' class='btn-apply'>Apply</button>");
            html.AppendLine("</div>");
            html.AppendLine("<div class='properties-list' id='propertiesList'>");

            foreach (var meta in keys)
            {
                string value = propValues.ContainsKey(meta.Key) ? propValues[meta.Key] : "";
                html.AppendLine("<div class='property-row'>");
                html.AppendLine($"  <div class='property-oblong' data-key='{WebUtility.HtmlEncode(meta.Key)}'>");
                html.AppendLine($"    <span class='property-icon'>{WebUtility.HtmlEncode(meta.Icon)}</span>");
                html.AppendLine($"    <span class='property-label'>{WebUtility.HtmlEncode(meta.Key)}</span>");
                html.AppendLine("  </div>");
                html.AppendLine($"  <div class='property-value-oblong'>");
                html.AppendLine($"    <input type='text' id='{WebUtility.HtmlEncode(meta.Key)}' value=\"{WebUtility.HtmlEncode(value).Replace("\"", "&quot;")}\" data-original=\"{WebUtility.HtmlEncode(value).Replace("\"", "&quot;")}\" />");
                html.AppendLine("  </div>");
                html.AppendLine("</div>");
            }

            html.AppendLine("</div>");
            html.AppendLine("<div id='addPropertyModal' class='modal-overlay'>");
            html.AppendLine("  <div class='modal-content'>");
            html.AppendLine("    <div class='modal-header'>");
            html.AppendLine("      <h2 class='modal-title'>➕ Ajouter une nouvelle propriété</h2>");
            html.AppendLine("    </div>");
            html.AppendLine("    <div class='form-group'>");
            html.AppendLine("      <label class='form-label' for='newPropertyName'>Nom de la propriété:</label>");
            html.AppendLine("      <input type='text' id='newPropertyName' class='form-input' placeholder='Ex: CustomProperty1'>");
            html.AppendLine("    </div>");
            html.AppendLine("    <div class='form-group'>");
            html.AppendLine("      <label class='form-label' for='newPropertyType'>Type de propriété:</label>");
            html.AppendLine("      <select id='newPropertyType' class='form-select'>");
            html.AppendLine("        <option value='text'>📝 Texte</option>");
            html.AppendLine("        <option value='date'>📅 Date</option>");
            html.AppendLine("        <option value='number'>🔢 Nombre</option>");
            html.AppendLine("        <option value='boolean'>✅ Oui/Non</option>");
            html.AppendLine("      </select>");
            html.AppendLine("    </div>");
            html.AppendLine("    <div class='form-group'>");
            html.AppendLine("      <label class='form-label' for='newPropertyValue'>Valeur initiale:</label>");
            html.AppendLine("      <input type='text' id='newPropertyValue' class='form-input' placeholder='Valeur par défaut'>");
            html.AppendLine("    </div>");
            html.AppendLine("    <div class='modal-buttons'>");
            html.AppendLine("      <button type='button' id='modalCancelBtn' class='btn-modal-cancel'>Annuler</button>");
            html.AppendLine("      <button type='button' id='modalConfirmBtn' class='btn-modal-confirm'>Créer</button>");
            html.AppendLine("    </div>");
            html.AppendLine("  </div>");
            html.AppendLine("</div>");
            html.AppendLine("<div class='footer'>Smart iProperties généré par Smart Tools Amine Add-in - iProperties Summary V1.3 @2025</div>");
            html.AppendLine("</div>");

            // JavaScript
            html.AppendLine("<script>");
            html.AppendLine("var originalValues = {};");
            html.AppendLine("var currentValues = {};");
            html.AppendLine("function initializeValues() {");
            foreach (var meta in keys)
            {
                string value = propValues.ContainsKey(meta.Key) ? propValues[meta.Key] : "";
                html.AppendLine($"    originalValues['{meta.Key.Replace("'", "\\'")}'] = '{value.Replace("'", "\\'")}';");
                html.AppendLine($"    currentValues['{meta.Key.Replace("'", "\\'")}'] = '{value.Replace("'", "\\'")}';");
            }
            html.AppendLine("}");
            html.AppendLine("function trackChanges() {");
            html.AppendLine("    var inputs = document.querySelectorAll('input[type=\"text\"]');");
            html.AppendLine("    var applyBtn = document.getElementById('applyBtn');");
            html.AppendLine("    function updateApplyBtnState() {");
            html.AppendLine("        var anyModified = false;");
            html.AppendLine("        for (var key in currentValues) {");
            html.AppendLine("            if (currentValues[key] !== originalValues[key]) { anyModified = true; break; }");
            html.AppendLine("        }");
            html.AppendLine("        if (anyModified) { applyBtn.classList.add('modified'); } else { applyBtn.classList.remove('modified'); }");
            html.AppendLine("    }");
            html.AppendLine("    for (var i = 0; i < inputs.length; i++) {");
            html.AppendLine("        inputs[i].addEventListener('input', function(e) {");
            html.AppendLine("            var key = e.target.id;");
            html.AppendLine("            var value = e.target.value;");
            html.AppendLine("            currentValues[key] = value;");
            html.AppendLine("            if (value !== originalValues[key]) { e.target.classList.add('modified'); }");
            html.AppendLine("            else { e.target.classList.remove('modified'); }");
            html.AppendLine("            updateApplyBtnState();");
            html.AppendLine("        });");
            html.AppendLine("    }");
            html.AppendLine("    updateApplyBtnState();");
            html.AppendLine("}");
            html.AppendLine("function showStatus(message, isSuccess) {");
            html.AppendLine("    var status = document.getElementById('status');");
            html.AppendLine("    status.textContent = message;");
            html.AppendLine("    status.className = 'status ' + (isSuccess ? 'success' : 'error');");
            html.AppendLine("    status.style.display = 'block';");
            html.AppendLine("    setTimeout(function() { status.style.display = 'none'; }, 3000);");
            html.AppendLine("}");
            html.AppendLine("function applyChanges() {");
            html.AppendLine("    try {");
            html.AppendLine("        var hasChanges = false;");
            html.AppendLine("        var changedProperties = '';");
            html.AppendLine("        for (var key in currentValues) {");
            html.AppendLine("            if (currentValues[key] !== originalValues[key]) { hasChanges = true; changedProperties += key + '=' + currentValues[key] + '|'; }");
            html.AppendLine("        }");
            html.AppendLine("        if (hasChanges) {");
            html.AppendLine("            if (window.chrome && window.chrome.webview) {");
            html.AppendLine("                window.chrome.webview.postMessage(JSON.stringify({action: 'ApplyChanges', data: changedProperties}));");
            html.AppendLine("            }");
            html.AppendLine("            showStatus('Properties successfully updated!', true);");
            html.AppendLine("            for (var key in currentValues) { originalValues[key] = currentValues[key]; document.getElementById(key).classList.remove('modified'); }");
            html.AppendLine("            document.getElementById('applyBtn').classList.remove('modified');");
            html.AppendLine("        } else { showStatus('No changes detected.', false); }");
            html.AppendLine("    } catch(e) { showStatus('Erreur: ' + e.message, false); }");
            html.AppendLine("}");
            html.AppendLine("function showAddPropertyModal() {");
            html.AppendLine("    document.getElementById('addPropertyModal').style.display = 'block';");
            html.AppendLine("    document.getElementById('newPropertyName').focus();");
            html.AppendLine("}");
            html.AppendLine("function hideAddPropertyModal() {");
            html.AppendLine("    document.getElementById('addPropertyModal').style.display = 'none';");
            html.AppendLine("    document.getElementById('newPropertyName').value = '';");
            html.AppendLine("    document.getElementById('newPropertyValue').value = '';");
            html.AppendLine("    document.getElementById('newPropertyType').value = 'text';");
            html.AppendLine("}");
            html.AppendLine("function addNewProperty() {");
            html.AppendLine("    var name = document.getElementById('newPropertyName').value.trim();");
            html.AppendLine("    var type = document.getElementById('newPropertyType').value;");
            html.AppendLine("    var value = document.getElementById('newPropertyValue').value;");
            html.AppendLine("    if (!name) { showStatus('Le nom de la propriété est requis!', false); return; }");
            html.AppendLine("    if (currentValues.hasOwnProperty(name)) { showStatus('Cette propriété existe déjà!', false); return; }");
            html.AppendLine("    try {");
            html.AppendLine("        if (window.chrome && window.chrome.webview) {");
            html.AppendLine("            window.chrome.webview.postMessage(JSON.stringify({action: 'AddNewProperty', data: {name: name, type: type, value: value}}));");
            html.AppendLine("        }");
            html.AppendLine("        hideAddPropertyModal();");
            html.AppendLine("        showStatus('Propriété ' + name + ' ajoutée avec succès!', true);");
            html.AppendLine("    } catch(e) { showStatus('Erreur: ' + e.message, false); }");
            html.AppendLine("}");
            html.AppendLine("window.onload = function() {");
            html.AppendLine("    initializeValues(); trackChanges();");
            html.AppendLine("    document.getElementById('applyBtn').onclick = applyChanges;");
            html.AppendLine("    document.getElementById('closeBtn').onclick = function() {");
            html.AppendLine("        if (window.chrome && window.chrome.webview) {");
            html.AppendLine("            window.chrome.webview.postMessage(JSON.stringify({action: 'CloseWindow'}));");
            html.AppendLine("        }");
            html.AppendLine("    };");
            html.AppendLine("    document.getElementById('addBtn').onclick = showAddPropertyModal;");
            html.AppendLine("    document.getElementById('modalCancelBtn').onclick = hideAddPropertyModal;");
            html.AppendLine("    document.getElementById('modalConfirmBtn').onclick = addNewProperty;");
            html.AppendLine("    document.getElementById('addPropertyModal').onclick = function(e) { if (e.target === this) hideAddPropertyModal(); };");
            html.AppendLine("};");
            html.AppendLine("</script>");
            html.AppendLine("</body></html>");

            return html.ToString();
        }

        /// <summary>
        /// Génère le HTML pour Smart Save avec barre de progression et changement dynamique de couleur du texte
        /// </summary>
        public static string GenerateSmartSaveProgressHtml(string docName, string typeText, List<string> steps, string primaryColor = "#2e7d32", string secondaryColor = "#4caf50")
        {
            var html = new StringBuilder();
            
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html lang='fr'>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset='UTF-8'>");
            html.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
            html.AppendLine("    <title>💾 Smart Save - V1.1 @2025 - By Mohammed Amine Elgalai</title>");
            html.AppendLine("    <style>");
            html.AppendLine("        @import url('https://fonts.googleapis.com/css2?family=Noto+Color+Emoji&display=swap');");
            html.AppendLine("        * { font-family: 'Segoe UI', 'Roboto', 'Noto Color Emoji', 'Apple Color Emoji', sans-serif; }");
            html.AppendLine($"        body {{ margin: 15px; background: linear-gradient(135deg, {primaryColor} 0%, {secondaryColor} 100%); font-size: 16px; min-height: calc(100vh - 30px); color: white; }}");
            html.AppendLine("        .container { max-width: 95%; margin: 0 auto; background: rgba(255,255,255,0.95); padding: 25px; border-radius: 12px; box-shadow: 0 8px 25px rgba(0,0,0,0.3); color: #333; }");
            html.AppendLine($"        h2 {{ color: {primaryColor}; font-size: 24px; text-align: center; margin-bottom: 20px; font-weight: bold; text-shadow: 1px 1px 2px rgba(0,0,0,0.1); }}");
            html.AppendLine($"        .info-box {{ background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); padding: 15px; border-radius: 8px; margin-bottom: 20px; border: 2px solid {secondaryColor}; }}");
            html.AppendLine($"        .info-box strong {{ color: {primaryColor}; }}");
            html.AppendLine("        ul { list-style: none; padding: 0; margin: 0; }");
            html.AppendLine("        li { margin: 8px 0; font-size: 15px; padding: 8px 12px; border-radius: 6px; display: flex; align-items: center; background: rgba(248,249,250,0.8); border-left: 4px solid #90a4ae; position: relative; overflow: hidden; transition: all 0.3s ease; }");
            html.AppendLine("        li .step-text { position: relative; z-index: 2; transition: color 0.3s ease; }");
            html.AppendLine("        li .progress-bar { position: absolute; left: 0; top: 0; height: 100%; width: 0%; background: linear-gradient(90deg, rgba(76,175,80,0.3) 0%, rgba(76,175,80,0.6) 100%); transition: width 0.5s ease; z-index: 1; }");
            html.AppendLine($"        li.completed {{ background: rgba(232,245,233,0.9); border-left-color: {secondaryColor}; }}");
            html.AppendLine($"        li.completed .progress-bar {{ width: 100%; background: linear-gradient(90deg, rgba(76,175,80,0.5) 0%, rgba(76,175,80,0.8) 100%); }}");
            html.AppendLine("        li.error { background: rgba(255,235,238,0.9); border-left-color: #f44336; }");
            html.AppendLine("        li.info { background: rgba(227,242,253,0.9); border-left-color: #2196f3; }");
            html.AppendLine("        li.active { background: rgba(255,243,224,0.9); border-left-color: #ff9800; }");
            html.AppendLine("        li.active .progress-bar { width: 50%; background: linear-gradient(90deg, rgba(255,152,0,0.3) 0%, rgba(255,152,0,0.6) 100%); }");
            html.AppendLine("        .emoji { font-size: 18px; margin-right: 10px; min-width: 25px; }");
            html.AppendLine($"        .completion {{ text-align: center; font-size: 18px; font-weight: bold; color: {primaryColor}; margin: 20px 0; padding: 15px; background: rgba(232,245,233,0.8); border-radius: 8px; display: none; }}");
            html.AppendLine($"        .btn-close {{ display: block; width: 150px; padding: 12px; margin: 20px auto; font-size: 16px; font-weight: bold; cursor: pointer; border: none; border-radius: 8px; background: {secondaryColor}; color: white; transition: all 0.3s; }}");
            html.AppendLine($"        .btn-close:hover {{ background: {primaryColor}; transform: scale(1.05); }}");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            html.AppendLine("    <div class='container'>");
            html.AppendLine($"        <h2><span class='emoji'>💾</span> Smart Save V1.1 - {WebUtility.HtmlEncode(typeText)}</h2>");
            html.AppendLine("        <div class='info-box'>");
            html.AppendLine($"            <span class='emoji'>📄</span> <strong>Document:</strong> {WebUtility.HtmlEncode(docName)}<br>");
            html.AppendLine($"            <span class='emoji'>📋</span> <strong>Type:</strong> {WebUtility.HtmlEncode(typeText)}<br>");
            html.AppendLine($"            <span class='emoji'>📅</span> <strong>Date:</strong> {DateTime.Now:yyyy-MM-dd HH:mm:ss} UTC<br>");
            html.AppendLine("            <span class='emoji'>👨‍💻</span> <strong>Développé par:</strong> Mohammed Amine Elgalai");
            html.AppendLine("        </div>");
            html.AppendLine("        <ul>");

            for (int i = 0; i < steps.Count; i++)
            {
                html.AppendLine($"            <li id='step{i + 1}'>");
                html.AppendLine("                <div class='progress-bar'></div>");
                html.AppendLine($"                <span class='step-text'><span class='emoji'>⏳</span> {WebUtility.HtmlEncode(steps[i])}</span>");
                html.AppendLine("            </li>");
            }

            html.AppendLine("        </ul>");
            html.AppendLine("        <div id='completion' class='completion'></div>");
            html.AppendLine("        <button class='btn-close' onclick='closeForm()'>✅ Fermer</button>");
            html.AppendLine("    </div>");

            // JavaScript pour le changement dynamique de couleur
            html.AppendLine("    <script>");
            html.AppendLine("        function getLuminance(r, g, b) {");
            html.AppendLine("            var [rs, gs, bs] = [r, g, b].map(c => {");
            html.AppendLine("                c = c / 255;");
            html.AppendLine("                return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);");
            html.AppendLine("            });");
            html.AppendLine("            return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;");
            html.AppendLine("        }");
            html.AppendLine("        function hexToRgb(hex) {");
            html.AppendLine("            var result = /^#?([a-f\\d]{2})([a-f\\d]{2})([a-f\\d]{2})$/i.exec(hex);");
            html.AppendLine("            return result ? {");
            html.AppendLine("                r: parseInt(result[1], 16),");
            html.AppendLine("                g: parseInt(result[2], 16),");
            html.AppendLine("                b: parseInt(result[3], 16)");
            html.AppendLine("            } : null;");
            html.AppendLine("        }");
            html.AppendLine("        function getBackgroundColor(element) {");
            html.AppendLine("            var style = window.getComputedStyle(element);");
            html.AppendLine("            var bgColor = style.backgroundColor || style.background;");
            html.AppendLine("            if (bgColor === 'rgba(0, 0, 0, 0)' || bgColor === 'transparent') {");
            html.AppendLine("                var parent = element.parentElement;");
            html.AppendLine("                while (parent && (bgColor === 'rgba(0, 0, 0, 0)' || bgColor === 'transparent')) {");
            html.AppendLine("                    var parentStyle = window.getComputedStyle(parent);");
            html.AppendLine("                    bgColor = parentStyle.backgroundColor || parentStyle.background;");
            html.AppendLine("                    parent = parent.parentElement;");
            html.AppendLine("                }");
            html.AppendLine("            }");
            html.AppendLine("            return bgColor;");
            html.AppendLine("        }");
            html.AppendLine("        function parseColor(color) {");
            html.AppendLine("            if (color.startsWith('#')) return hexToRgb(color);");
            html.AppendLine("            if (color.startsWith('rgb')) {");
            html.AppendLine("                var matches = color.match(/\\d+/g);");
            html.AppendLine("                return matches ? { r: parseInt(matches[0]), g: parseInt(matches[1]), b: parseInt(matches[2]) } : null;");
            html.AppendLine("            }");
            html.AppendLine("            if (color.startsWith('rgba')) {");
            html.AppendLine("                var matches = color.match(/\\d+/g);");
            html.AppendLine("                return matches ? { r: parseInt(matches[0]), g: parseInt(matches[1]), b: parseInt(matches[2]) } : null;");
            html.AppendLine("            }");
            html.AppendLine("            return null;");
            html.AppendLine("        }");
            html.AppendLine("        function updateTextColor(stepElement) {");
            html.AppendLine("            var progressBar = stepElement.querySelector('.progress-bar');");
            html.AppendLine("            var stepText = stepElement.querySelector('.step-text');");
            html.AppendLine("            if (!progressBar || !stepText) return;");
            html.AppendLine("            var progressWidth = parseFloat(progressBar.style.width) || 0;");
            html.AppendLine("            if (progressWidth > 0) {");
            html.AppendLine("                var bgColor = getBackgroundColor(progressBar);");
            html.AppendLine("                var rgb = parseColor(bgColor);");
            html.AppendLine("                if (rgb) {");
            html.AppendLine("                    var luminance = getLuminance(rgb.r, rgb.g, rgb.b);");
            html.AppendLine("                    stepText.style.color = luminance > 0.5 ? '#000000' : '#FFFFFF';");
            html.AppendLine("                    stepText.style.textShadow = luminance > 0.5 ? '1px 1px 2px rgba(255,255,255,0.8)' : '1px 1px 2px rgba(0,0,0,0.8)';");
            html.AppendLine("                    stepText.style.fontWeight = 'bold';");
            html.AppendLine("                }");
            html.AppendLine("            } else {");
            html.AppendLine("                stepText.style.color = '';");
            html.AppendLine("                stepText.style.textShadow = '';");
            html.AppendLine("                stepText.style.fontWeight = '';");
            html.AppendLine("            }");
            html.AppendLine("        }");
            html.AppendLine("        function observeProgressBars() {");
            html.AppendLine("            var steps = document.querySelectorAll('li[id^=\"step\"]');");
            html.AppendLine("            steps.forEach(function(step) {");
            html.AppendLine("                var progressBar = step.querySelector('.progress-bar');");
            html.AppendLine("                if (progressBar) {");
            html.AppendLine("                    var observer = new MutationObserver(function() {");
            html.AppendLine("                        updateTextColor(step);");
            html.AppendLine("                    });");
            html.AppendLine("                    observer.observe(progressBar, { attributes: true, attributeFilter: ['style'] });");
            html.AppendLine("                    updateTextColor(step);");
            html.AppendLine("                }");
            html.AppendLine("            });");
            html.AppendLine("        }");
            html.AppendLine("        function closeForm() {");
            html.AppendLine("            if (window.chrome && window.chrome.webview) {");
            html.AppendLine("                window.chrome.webview.postMessage(JSON.stringify({action: 'CloseWindow'}));");
            html.AppendLine("            }");
            html.AppendLine("        }");
            html.AppendLine("        window.onload = function() {");
            html.AppendLine("            observeProgressBars();");
            html.AppendLine("            setInterval(function() {");
            html.AppendLine("                var steps = document.querySelectorAll('li[id^=\"step\"]');");
            html.AppendLine("                steps.forEach(updateTextColor);");
            html.AppendLine("            }, 100);");
            html.AppendLine("        };");
            html.AppendLine("    </script>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            return html.ToString();
        }

        /// <summary>
        /// Génère le HTML pour Safe Close avec barre de progression et changement dynamique de couleur du texte
        /// </summary>
        public static string GenerateSafeCloseProgressHtml(string docName, string typeText, List<string> steps, string primaryColor = "#0d47a1", string secondaryColor = "#1976d2")
        {
            // Même structure que Smart Save mais avec les couleurs bleues
            return GenerateSmartSaveProgressHtml(docName, typeText, steps, primaryColor, secondaryColor)
                .Replace("Smart Save", "Safe Close")
                .Replace("💾", "🔒")
                .Replace("Sauvegarde", "Fermeture");
        }

        /// <summary>
        /// Génère le HTML pour HideBox avec barre de progression et changement dynamique de couleur du texte
        /// </summary>
        public static string GenerateHideBoxProgressHtml(string actionEnCours, int nombreTotal, int nombreVisibles, int nombreCaches, bool isXnrgyDoor = false)
        {
            var html = new StringBuilder();
            string couleurAction = actionEnCours.Contains("RÉAFFICHAGE") ? "#28a745" : "#dc3545";

            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html><head><meta charset='UTF-8'>");
            html.AppendLine("<style>");
            html.AppendLine("body { font-family: 'Segoe UI', Arial, sans-serif; background: linear-gradient(135deg, #667eea, #764ba2); margin: 0; padding: 10px; color: #333; overflow: auto; }");
            html.AppendLine(".container { background: rgba(255,255,255,0.95); border-radius: 10px; padding: 15px; box-shadow: 0 10px 20px rgba(0,0,0,0.3); width: 100%; min-height: 300px; box-sizing: border-box; display: flex; flex-direction: column; justify-content: flex-start; }");
            html.AppendLine("h1 { color: #444; margin: 0 0 20px 0; font-size: 20px; text-align: center; }");
            html.AppendLine($".step {{ margin: 10px 0; padding: 12px; background: #f8f9fa; border-radius: 8px; border-left: 4px solid {couleurAction}; text-align: left; font-size: 16px; line-height: 1.4; position: relative; overflow: hidden; transition: all 0.3s ease; }}");
            html.AppendLine(".step .step-text { position: relative; z-index: 2; transition: color 0.3s ease; }");
            html.AppendLine($".step .progress-bar {{ position: absolute; left: 0; top: 0; height: 100%; width: 0%; background: linear-gradient(90deg, {couleurAction}33 0%, {couleurAction}66 100%); transition: width 0.5s ease; z-index: 1; }}");
            html.AppendLine(".step span { margin-right: 10px; }");
            html.AppendLine(".analysis { background: #e3f2fd; border-left-color: #2196f3; margin-bottom: 15px; }");
            html.AppendLine(".spinner { display: inline-block; width: 18px; height: 18px; border: 2px solid #f3f3f3; border-top: 2px solid #3498db; border-radius: 50%; animation: spin 1s linear infinite; }");
            html.AppendLine("@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }");
            html.AppendLine("</style></head><body>");
            html.AppendLine("<div class='container'>");
            html.AppendLine("<h1>🔄 Smart Hide/Show Box V1.5 - Mode Intelligent</h1>");
            html.AppendLine("<div class='step analysis'>");
            html.AppendLine($"<span>🔍</span> Analyse: {nombreTotal} composants détectés ({nombreVisibles} visibles, {nombreCaches} cachés)<br>");
            html.AppendLine("<span>🔧</span> Inclut: Box, Template, Multi_Opening, Cut_Opening, tous types de Dummy, AirFlow, PanFactice, etc.");
            html.AppendLine("</div>");

            if (isXnrgyDoor)
            {
                html.AppendLine("<div class='step analysis' style='background:#fff3e0; border-left-color:#ff9800;'>");
                html.AppendLine("<span>🚪</span> <strong>Xnrgy_Door détecté:</strong> Gestion spéciale des OpenDummy_1<br>");
                html.AppendLine($"<span>🔄</span> <strong>Alternance intelligente:</strong> {actionEnCours} basé sur l'état précédent<br>");
                html.AppendLine("<span>🎯</span> <strong>Optimisation:</strong> Détection des External_Swing et Internal_Swing");
                html.AppendLine("</div>");
            }

            html.AppendLine($"<div id='step1' class='step'>");
            html.AppendLine("    <div class='progress-bar'></div>");
            html.AppendLine($"    <span class='step-text'><span class='spinner'></span> Étape 1: {actionEnCours} des composants de référence en cours...</span>");
            html.AppendLine("</div>");
            html.AppendLine($"<div id='step2' class='step'>");
            html.AppendLine("    <div class='progress-bar'></div>");
            html.AppendLine("    <span class='step-text'><span class='spinner'></span> Étape 2: Mise à jour du document en cours...</span>");
            html.AppendLine("</div>");
            html.AppendLine("</div>");

            // JavaScript pour le changement dynamique de couleur (même logique que Smart Save)
            html.AppendLine("<script>");
            html.AppendLine("        function getLuminance(r, g, b) {");
            html.AppendLine("            var [rs, gs, bs] = [r, g, b].map(c => {");
            html.AppendLine("                c = c / 255;");
            html.AppendLine("                return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);");
            html.AppendLine("            });");
            html.AppendLine("            return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;");
            html.AppendLine("        }");
            html.AppendLine("        function hexToRgb(hex) {");
            html.AppendLine("            var result = /^#?([a-f\\d]{2})([a-f\\d]{2})([a-f\\d]{2})$/i.exec(hex);");
            html.AppendLine("            return result ? {");
            html.AppendLine("                r: parseInt(result[1], 16),");
            html.AppendLine("                g: parseInt(result[2], 16),");
            html.AppendLine("                b: parseInt(result[3], 16)");
            html.AppendLine("            } : null;");
            html.AppendLine("        }");
            html.AppendLine("        function getBackgroundColor(element) {");
            html.AppendLine("            var style = window.getComputedStyle(element);");
            html.AppendLine("            var bgColor = style.backgroundColor || style.background;");
            html.AppendLine("            if (bgColor === 'rgba(0, 0, 0, 0)' || bgColor === 'transparent') {");
            html.AppendLine("                var parent = element.parentElement;");
            html.AppendLine("                while (parent && (bgColor === 'rgba(0, 0, 0, 0)' || bgColor === 'transparent')) {");
            html.AppendLine("                    var parentStyle = window.getComputedStyle(parent);");
            html.AppendLine("                    bgColor = parentStyle.backgroundColor || parentStyle.background;");
            html.AppendLine("                    parent = parent.parentElement;");
            html.AppendLine("                }");
            html.AppendLine("            }");
            html.AppendLine("            return bgColor;");
            html.AppendLine("        }");
            html.AppendLine("        function parseColor(color) {");
            html.AppendLine("            if (color.startsWith('#')) return hexToRgb(color);");
            html.AppendLine("            if (color.startsWith('rgb')) {");
            html.AppendLine("                var matches = color.match(/\\d+/g);");
            html.AppendLine("                return matches ? { r: parseInt(matches[0]), g: parseInt(matches[1]), b: parseInt(matches[2]) } : null;");
            html.AppendLine("            }");
            html.AppendLine("            if (color.startsWith('rgba')) {");
            html.AppendLine("                var matches = color.match(/\\d+/g);");
            html.AppendLine("                return matches ? { r: parseInt(matches[0]), g: parseInt(matches[1]), b: parseInt(matches[2]) } : null;");
            html.AppendLine("            }");
            html.AppendLine("            return null;");
            html.AppendLine("        }");
            html.AppendLine("        function updateTextColor(stepElement) {");
            html.AppendLine("            var progressBar = stepElement.querySelector('.progress-bar');");
            html.AppendLine("            var stepText = stepElement.querySelector('.step-text');");
            html.AppendLine("            if (!progressBar || !stepText) return;");
            html.AppendLine("            var progressWidth = parseFloat(progressBar.style.width) || 0;");
            html.AppendLine("            if (progressWidth > 0) {");
            html.AppendLine("                var bgColor = getBackgroundColor(progressBar);");
            html.AppendLine("                var rgb = parseColor(bgColor);");
            html.AppendLine("                if (rgb) {");
            html.AppendLine("                    var luminance = getLuminance(rgb.r, rgb.g, rgb.b);");
            html.AppendLine("                    stepText.style.color = luminance > 0.5 ? '#000000' : '#FFFFFF';");
            html.AppendLine("                    stepText.style.textShadow = luminance > 0.5 ? '1px 1px 2px rgba(255,255,255,0.8)' : '1px 1px 2px rgba(0,0,0,0.8)';");
            html.AppendLine("                    stepText.style.fontWeight = 'bold';");
            html.AppendLine("                }");
            html.AppendLine("            } else {");
            html.AppendLine("                stepText.style.color = '';");
            html.AppendLine("                stepText.style.textShadow = '';");
            html.AppendLine("                stepText.style.fontWeight = '';");
            html.AppendLine("            }");
            html.AppendLine("        }");
            html.AppendLine("        function observeProgressBars() {");
            html.AppendLine("            var steps = document.querySelectorAll('.step[id^=\"step\"]');");
            html.AppendLine("            steps.forEach(function(step) {");
            html.AppendLine("                var progressBar = step.querySelector('.progress-bar');");
            html.AppendLine("                if (progressBar) {");
            html.AppendLine("                    var observer = new MutationObserver(function() {");
            html.AppendLine("                        updateTextColor(step);");
            html.AppendLine("                    });");
            html.AppendLine("                    observer.observe(progressBar, { attributes: true, attributeFilter: ['style'] });");
            html.AppendLine("                    updateTextColor(step);");
            html.AppendLine("                }");
            html.AppendLine("            });");
            html.AppendLine("        }");
            html.AppendLine("        window.onload = function() {");
            html.AppendLine("            observeProgressBars();");
            html.AppendLine("            setInterval(function() {");
            html.AppendLine("                var steps = document.querySelectorAll('.step[id^=\"step\"]');");
            html.AppendLine("                steps.forEach(updateTextColor);");
            html.AppendLine("            }, 100);");
            html.AppendLine("        };");
            html.AppendLine("</script>");
            html.AppendLine("</body></html>");

            return html.ToString();
        }
    }

    /// <summary>
    /// Métadonnées d'une propriété
    /// </summary>
    public class PropertyMeta
    {
        public string Key { get; set; } = "";
        public string Icon { get; set; } = "🔸";
        public string SetName { get; set; } = "Custom";

        public PropertyMeta(string key, string icon, string setName)
        {
            Key = key;
            Icon = icon;
            SetName = setName;
        }
    }
}


