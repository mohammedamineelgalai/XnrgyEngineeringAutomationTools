using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using iTextSharp.text.pdf;
using NLog;

namespace XnrgyEngineeringAutomationTools.Modules.CreateModule.Services
{
    /// <summary>
    /// Service pour remplir les PDFs de couverture BatchPrint avec les informations du projet
    /// PDFs source: C:\Vault\Engineering\Library\Xnrgy_Module\6-Shop Drawing PDF\Production\
    /// PDFs destination: C:\Vault\Engineering\Projects\{project}\REF{ref}\M{module}\6-Shop Drawing PDF\Production\
    /// By Mohammed Amine Elgalai - XNRGY Climate Systems ULC
    /// </summary>
    public class PdfCoverService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        // Chemin source des templates PDF
        private const string TEMPLATE_PDF_PATH = @"C:\Vault\Engineering\Library\Xnrgy_Module\6-Shop Drawing PDF\Production";
        
        // Liste des fichiers PDF de couverture
        private static readonly string[] PDF_COVER_FILES = new[]
        {
            "01-Plancher.pdf",
            "02-Machines.pdf",
            "03-Mur.pdf",
            "04-Plafond.pdf",
            "05-Portes.pdf",
            "06-Xnfan.pdf",
            "07-Electricite.pdf",
            "08-L-10 @ L-50.pdf",
            "09-L-60 @ L-100.pdf",
            "10-Submital.pdf",
            "11-KnockDown.pdf"
        };
        
        // Noms des champs PDF (decouverts par analyse PowerShell le 08/01/2026)
        private const string FIELD_PROJECT = "NUMBER";
        private const string FIELD_REF = "Dropdown7";
        private const string FIELD_MOD = "Dropdown10";
        private const string FIELD_JOB_TITLE = "Nom of the job";
        
        private readonly Action<string, string>? _logCallback;

        /// <summary>
        /// Constructeur avec callback de log optionnel
        /// </summary>
        public PdfCoverService(Action<string, string>? logCallback = null)
        {
            _logCallback = logCallback;
        }

        /// <summary>
        /// Remplit tous les PDFs de couverture avec les informations du projet
        /// </summary>
        /// <param name="destinationPath">Chemin destination (ex: C:\Vault\Engineering\Projects\12345\REF01\M01)</param>
        /// <param name="projectNumber">Numero de projet (ex: 12345)</param>
        /// <param name="reference">Reference sans prefixe (ex: 01)</param>
        /// <param name="module">Module sans prefixe (ex: 01)</param>
        /// <param name="jobTitle">Titre du job/projet (ex: HOSPITAL ABC MONTREAL)</param>
        /// <returns>Nombre de PDFs remplis avec succes</returns>
        public int FillAllCoverPdfs(string destinationPath, string projectNumber, string reference, string module, string jobTitle)
        {
            int successCount = 0;
            int failCount = 0;
            int notFoundCount = 0;
            var pdfDestFolder = Path.Combine(destinationPath, "6-Shop Drawing PDF", "Production");
            
            // ═══════════════════════════════════════════════════════════════════════
            // DEBUT OPERATION: REMPLISSAGE PDF COVERS BATCHPRINT
            // ═══════════════════════════════════════════════════════════════════════
            Logger.Info("[PdfCoverService] ═══════════════════════════════════════════════════════════");
            Logger.Info("[PdfCoverService] DEBUT FillAllCoverPdfs()");
            Logger.Info($"[PdfCoverService] pdfDestFolder = {pdfDestFolder}");
            
            Log("", "INFO");
            Log("═══════════════════════════════════════════════════════════════", "INFO");
            Log("[>] REMPLISSAGE PDFs DE COUVERTURE BATCHPRINT", "START");
            Log("═══════════════════════════════════════════════════════════════", "INFO");
            
            // Afficher les parametres
            Log($"   Dossier: {pdfDestFolder}", "INFO");
            Log($"   Projet: {projectNumber}", "INFO");
            Log($"   REF: {reference} | MOD: {module}", "INFO");
            Log($"   Job Title: {jobTitle}", "INFO");
            Logger.Info($"[PdfCoverService] Parametres: Projet={projectNumber}, REF={reference}, MOD={module}");
            Logger.Info($"[PdfCoverService] Job Title: {jobTitle}");
            
            // Verifier que le dossier destination existe
            Logger.Info($"[PdfCoverService] Verification dossier: {pdfDestFolder}");
            bool folderExists = Directory.Exists(pdfDestFolder);
            Logger.Info($"[PdfCoverService] Directory.Exists = {folderExists}");
            
            if (!folderExists)
            {
                Log($"[!] Dossier PDF non trouve: {pdfDestFolder}", "WARN");
                Logger.Warn($"[PdfCoverService] DOSSIER INEXISTANT: {pdfDestFolder}");
                return 0;
            }
            
            // Lister les fichiers PDF dans le dossier
            var existingPdfs = Directory.GetFiles(pdfDestFolder, "*.pdf");
            Logger.Info($"[PdfCoverService] PDFs dans le dossier: {existingPdfs.Length}");
            foreach (var pdf in existingPdfs)
            {
                Logger.Info($"[PdfCoverService]   - {Path.GetFileName(pdf)}");
            }
            
            Log($"[>] Traitement de {PDF_COVER_FILES.Length} fichiers PDF...", "INFO");
            Logger.Info($"[PdfCoverService] Traitement de {PDF_COVER_FILES.Length} fichiers de la liste");
            
            var results = new StringBuilder();
            results.AppendLine("[RESULTATS DETAILLES]");
            
            foreach (var pdfFile in PDF_COVER_FILES)
            {
                var pdfPath = Path.Combine(pdfDestFolder, pdfFile);
                Logger.Info($"[PdfCoverService] Traitement: {pdfFile}");
                Logger.Info($"[PdfCoverService]   Chemin complet: {pdfPath}");
                Logger.Info($"[PdfCoverService]   File.Exists: {File.Exists(pdfPath)}");
                
                if (!File.Exists(pdfPath))
                {
                    notFoundCount++;
                    Log($"   [~] {pdfFile}: Non trouve", "WARN");
                    Logger.Warn($"[PdfCoverService]   FICHIER NON TROUVE: {pdfFile}");
                    results.AppendLine($"   [-] {pdfFile}: Non trouve");
                    continue;
                }
                
                try
                {
                    Logger.Info($"[PdfCoverService]   Appel FillCoverPdfWithDetails...");
                    var fillResult = FillCoverPdfWithDetails(pdfPath, projectNumber, reference, module, jobTitle);
                    Logger.Info($"[PdfCoverService]   Resultat: Success={fillResult.Success}, FieldsFilled={fillResult.FieldsFilled}");
                    
                    if (fillResult.Success)
                    {
                        successCount++;
                        Log($"   [+] {pdfFile}", "SUCCESS");
                        Logger.Info($"[PdfCoverService]   [+] {pdfFile}: OK ({fillResult.FieldsFilled} champs)");
                        results.AppendLine($"   [+] {pdfFile}: OK ({fillResult.FieldsFilled} champs)");
                    }
                    else
                    {
                        failCount++;
                        Log($"   [-] {pdfFile}: {fillResult.ErrorMessage}", "ERROR");
                        Logger.Error($"[PdfCoverService]   [-] {pdfFile}: {fillResult.ErrorMessage}");
                        results.AppendLine($"   [-] {pdfFile}: ECHEC - {fillResult.ErrorMessage}");
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    Log($"   [-] {pdfFile}: Exception - {ex.Message}", "ERROR");
                    Logger.Error($"[PdfCoverService]   EXCEPTION: {ex.Message}");
                    Logger.Error($"[PdfCoverService]   StackTrace: {ex.StackTrace}");
                    results.AppendLine($"   [-] {pdfFile}: EXCEPTION - {ex.Message}");
                }
            }
            
            // ═══════════════════════════════════════════════════════════════════════
            // RESUME FINAL
            // ═══════════════════════════════════════════════════════════════════════
            Log("───────────────────────────────────────────────────────────────", "INFO");
            
            if (successCount > 0 && failCount == 0)
            {
                Log($"[+] SUCCES: {successCount}/{PDF_COVER_FILES.Length} PDFs remplis", "SUCCESS");
                Logger.Info($"[+] SUCCES COMPLET: {successCount}/{PDF_COVER_FILES.Length} PDFs remplis");
            }
            else if (successCount > 0)
            {
                Log($"[!] PARTIEL: {successCount} OK, {failCount} echecs, {notFoundCount} non trouves", "WARN");
                Logger.Warn($"[!] SUCCES PARTIEL: {successCount} OK, {failCount} echecs, {notFoundCount} non trouves");
            }
            else
            {
                Log($"[-] ECHEC: Aucun PDF rempli (0/{PDF_COVER_FILES.Length})", "ERROR");
                Logger.Error($"[-] ECHEC TOTAL: 0/{PDF_COVER_FILES.Length} PDFs remplis");
            }
            
            Log("═══════════════════════════════════════════════════════════════", "INFO");
            Logger.Info("═══════════════════════════════════════════════════════════════");
            Logger.Info($"[FIN] Remplissage PDF Covers termine");
            Logger.Info(results.ToString());
            
            return successCount;
        }

        /// <summary>
        /// Resultat detaille du remplissage d'un PDF
        /// </summary>
        public class FillResult
        {
            public bool Success { get; set; }
            public int FieldsFilled { get; set; }
            public string ErrorMessage { get; set; } = "";
            public List<string> FieldsSet { get; set; } = new List<string>();
        }

        /// <summary>
        /// Remplit un PDF de couverture avec details sur les champs remplis
        /// </summary>
        private FillResult FillCoverPdfWithDetails(string pdfPath, string projectNumber, string reference, string module, string jobTitle)
        {
            var result = new FillResult();
            
            if (!File.Exists(pdfPath))
            {
                result.ErrorMessage = "Fichier introuvable";
                return result;
            }
            
            try
            {
                var tempPath = pdfPath + ".tmp";
                
                using (var reader = new PdfReader(pdfPath))
                using (var stamper = new PdfStamper(reader, new FileStream(tempPath, FileMode.Create)))
                {
                    var form = stamper.AcroFields;
                    
                    // Logger tous les champs disponibles (niveau DEBUG)
                    var availableFields = new List<string>();
                    foreach (var fieldName in form.Fields.Keys)
                    {
                        availableFields.Add(fieldName.ToString());
                    }
                    Logger.Debug($"      Champs disponibles dans {Path.GetFileName(pdfPath)}: {string.Join(", ", availableFields)}");
                    
                    // Remplir les 4 champs avec suivi detaille
                    if (SetFieldWithTracking(form, FIELD_PROJECT, projectNumber, result))
                        Logger.Debug($"      {FIELD_PROJECT} = {projectNumber}");
                    
                    if (SetFieldWithTracking(form, FIELD_REF, reference, result))
                        Logger.Debug($"      {FIELD_REF} = {reference}");
                    
                    if (SetFieldWithTracking(form, FIELD_MOD, module, result))
                        Logger.Debug($"      {FIELD_MOD} = {module}");
                    
                    if (SetFieldWithTracking(form, FIELD_JOB_TITLE, jobTitle, result))
                        Logger.Debug($"      {FIELD_JOB_TITLE} = {jobTitle}");
                }
                
                // Remplacer le fichier original
                File.Delete(pdfPath);
                File.Move(tempPath, pdfPath);
                
                result.Success = result.FieldsFilled > 0;
                if (!result.Success)
                {
                    result.ErrorMessage = "Aucun champ trouve dans le PDF";
                }
                
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                Logger.Error($"      Exception: {ex.Message}");
                
                // Nettoyer le fichier temporaire
                var tempPath = pdfPath + ".tmp";
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
                
                return result;
            }
        }

        /// <summary>
        /// Definit un champ et trace le resultat
        /// </summary>
        private bool SetFieldWithTracking(AcroFields form, string fieldName, string value, FillResult result)
        {
            if (form.Fields.ContainsKey(fieldName))
            {
                bool success = form.SetField(fieldName, value);
                if (success)
                {
                    result.FieldsFilled++;
                    result.FieldsSet.Add($"{fieldName}={value}");
                }
                return success;
            }
            return false;
        }

        /// <summary>
        /// Analyse un PDF et retourne la liste des champs de formulaire
        /// Utile pour le debug et la decouverte des noms de champs
        /// </summary>
        public List<string> GetPdfFormFields(string pdfPath)
        {
            var fields = new List<string>();
            
            if (!File.Exists(pdfPath))
                return fields;
            
            try
            {
                using (var reader = new PdfReader(pdfPath))
                {
                    var form = reader.AcroFields;
                    foreach (var field in form.Fields)
                    {
                        fields.Add($"{field.Key} = '{form.GetField(field.Key.ToString())}'");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"[-] Erreur lecture champs PDF: {ex.Message}");
            }
            
            return fields;
        }

        /// <summary>
        /// Log un message via le callback et le logger
        /// </summary>
        private void Log(string message, string level)
        {
            _logCallback?.Invoke(message, level);
            
            // Ne pas doubler les logs dans le logger - deja loggé ailleurs
        }
    }
}
