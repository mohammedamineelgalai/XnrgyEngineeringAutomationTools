// =============================================================================
// ExcelManagerService.cs - Gestionnaire Excel pour DXF Verifier
// MIGRATION EXACTE depuis ExcelManager.vb - NE PAS MODIFIER LA LOGIQUE
// Auteur original: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
// Version: 1.2 - Portage C# depuis VB.NET
// =============================================================================
// [!!!] CE CODE A ETE CALIBRE - NE PAS TOUCHER LA LOGIQUE CRITIQUE [!!!]
// =============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Services
{
    /// <summary>
    /// ExcelManager Ultra-Avancé v1.2.0 - DXF-CSV vs PDF Verifier v1.2
    /// Gestion robuste CSV, Excel avec renommage automatique et formatage professionnel
    /// Version portée depuis VB.NET avec logique IDENTIQUE
    /// </summary>
    public static class ExcelManagerService
    {
        // Événement pour le logging (sera connecté au journal de l'UI)
        public static event Action<string, string>? OnLog;

        #region Lecture CSV

        /// <summary>
        /// Lit le contenu d'un fichier CSV - Version restaurée et optimisée
        /// Préserve les valeurs exactes du CSV sans les modifier (logique critique)
        /// </summary>
        public static List<DxfItem> ReadCsvFile(string csvPath)
        {
            var dxfList = new List<DxfItem>();
            var tagsDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Détecter l'encodage du fichier
                var encoding = DetectEncoding(csvPath);
                Log("FileIO", $"[>] Lecture du fichier CSV: {csvPath} avec encodage {encoding.EncodingName}");

                using (var reader = new StreamReader(csvPath, encoding))
                {
                    int lineNumber = 0;
                    bool hasHeader = false;

                    while (!reader.EndOfStream)
                    {
                        lineNumber++;
                        string? line = reader.ReadLine();

                        // Ignorer les lignes vides
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        // Vérifier si la première ligne est un en-tête
                        if (lineNumber == 1 && (
                            line.Contains("Qtée") || line.Contains("Tag") ||
                            line.Contains("QTY") || line.Contains("Qty")))
                        {
                            hasHeader = true;
                            continue;
                        }

                        // Détecter le séparateur utilisé
                        char separator = DetectSeparator(line);

                        // Diviser la ligne selon le séparateur
                        string[] parts = line.Split(new char[] { separator }, StringSplitOptions.None);

                        if (parts.Length >= 2)
                        {
                            // Format du CSV: "Quantité, Tag, [Matériau]"
                            int quantity = 0;

                            // Nettoyer les chaînes
                            string qtyStr = CleanCsvValue(parts[0]);
                            string tag = CleanCsvValue(parts[1]);

                            // Récupérer le matériau s'il existe
                            string material = "";
                            if (parts.Length >= 3)
                            {
                                material = CleanCsvValue(parts[2]);
                            }

                            // Nettoyer le tag - Supprimer l'extension .dxf si présente
                            if (tag.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase))
                            {
                                tag = tag.Substring(0, tag.Length - 4);
                            }

                            // Convertir la quantité (IMPORTANT: toujours respecter la valeur exacte du CSV)
                            if (int.TryParse(qtyStr, out quantity) && quantity > 0)
                            {
                                // CRITIQUE: Si ce tag existe déjà, nous prenons la PREMIÈRE occurrence
                                if (!tagsDict.TryAdd(tag, quantity))
                                {
                                    Log("FileIO", $"[!] ATTENTION: Tag duplique dans le CSV - {tag}, " +
                                        $"Valeur originale: {tagsDict[tag]}, Nouvelle valeur ignoree: {quantity}");
                                }
                                else
                                {
                                    // Créer l'objet DxfItem avec valeur exacte du CSV
                                    var dxfItem = new DxfItem
                                    {
                                        Quantity = quantity,
                                        Tag = tag,
                                        Material = material,
                                        FoundInPdf = false,
                                        PdfQuantity = 0
                                    };
                                    dxfList.Add(dxfItem);
                                }
                            }
                            else
                            {
                                Log("FileIO", $"[!] Avertissement: Quantite invalide dans le CSV ligne {lineNumber}: '{qtyStr}' pour le tag {tag}");
                            }
                        }
                        else
                        {
                            Log("FileIO", $"[!] Avertissement: Format de ligne incorrect au CSV ligne {lineNumber}: {line}");
                        }
                    }
                }

                // Trier la liste par tag avant de la retourner
                dxfList = dxfList.OrderBy(item => item.Tag).ToList();

                Log("FileIO", $"[+] Lecture CSV terminee: {dxfList.Count} elements charges et tries par ordre alphabetique");
            }
            catch (IOException ex)
            {
                Log("Error", $"[-] ERREUR E/S lors de la lecture du CSV: {ex.Message}");
                throw;
            }
            catch (UnauthorizedAccessException ex)
            {
                Log("Error", $"[-] ERREUR d'acces lors de la lecture du CSV: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] ERREUR critique lors de la lecture du CSV: {ex.Message}");
                throw;
            }

            return dxfList;
        }

        /// <summary>
        /// Lecture CSV comme lignes de Cut List pour la nouvelle architecture multi-pass
        /// Compatible avec PdfAnalyzerService
        /// </summary>
        public static Dictionary<string, PdfAnalyzerService.CsvRow> ReadCsvFileAsCutListRows(string csvPath)
        {
            var csvDict = new Dictionary<string, PdfAnalyzerService.CsvRow>(StringComparer.OrdinalIgnoreCase);

            try
            {
                Log("FileIO", $"[>] Lecture CSV pour Cut List Simple: {csvPath}");

                var encoding = DetectEncoding(csvPath);
                using (var reader = new StreamReader(csvPath, encoding))
                {
                    int lineNumber = 0;

                    while (!reader.EndOfStream)
                    {
                        lineNumber++;
                        string? line = reader.ReadLine();

                        if (string.IsNullOrWhiteSpace(line)) continue;

                        // Ignorer les en-têtes
                        if (IsHeaderLine(line))
                        {
                            continue;
                        }

                        // Traiter la ligne CSV
                        ProcessCsvLine(line, lineNumber, csvDict);
                    }
                }

                Log("FileIO", $"[+] CSV Cut List charge: {csvDict.Count} tags uniques");
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] Erreur lecture CSV Cut List: {ex.Message}");
                throw;
            }

            return csvDict;
        }

        #endregion

        #region Écriture Excel

        /// <summary>
        /// Écrit les données de DXF dans le fichier Excel - Version EPPlus moderne et robuste
        /// AVEC mise à jour automatique des cellules projet (C1: Date, C2: Projet, C3: Ref, C4: Module)
        /// </summary>
        public static void WriteToExcel(string excelPath, List<DxfItem> dxfItems)
        {
            try
            {
                // Vérifier et fermer Excel si le fichier est ouvert
                CloseExcelIfFileIsOpen(excelPath);

                // Trier les éléments DXF par tag avant de les écrire
                dxfItems = dxfItems.OrderBy(item => item.Tag).ToList();

                Log("FileIO", $"[>] Debut ecriture Excel: {excelPath} avec {dxfItems.Count} elements");

                using (var package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    ExcelWorksheet? worksheet = null;

                    // Trouver ou créer la feuille DXF_vs_PDF
                    if (package.Workbook.Worksheets.Any(ws => ws.Name == "DXF_vs_PDF"))
                    {
                        worksheet = package.Workbook.Worksheets["DXF_vs_PDF"];
                    }
                    else if (package.Workbook.Worksheets.Count > 0)
                    {
                        worksheet = package.Workbook.Worksheets[0];
                        worksheet.Name = "DXF_vs_PDF";
                    }
                    else
                    {
                        worksheet = package.Workbook.Worksheets.Add("DXF_vs_PDF");
                    }

                    // Mise à jour automatique des cellules projet
                    UpdateProjectCellsInExcel(worksheet, excelPath);

                    // Ligne de départ (A11 selon les spécifications)
                    int startRow = 11;

                    // Nettoyer les données existantes
                    if (worksheet.Dimension != null)
                    {
                        int lastRow = Math.Max(worksheet.Dimension.End.Row, startRow + dxfItems.Count + 10);
                        int lastCol = Math.Max(worksheet.Dimension.End.Column, 15);

                        // Nettoyer TOUTE la zone de données
                        worksheet.Cells[10, 1, lastRow, lastCol].Clear();

                        // Double sécurité: nettoyer spécifiquement les colonnes F et au-delà
                        for (int row = 10; row <= lastRow; row++)
                        {
                            for (int col = 6; col <= lastCol; col++)
                            {
                                worksheet.Cells[row, col].Value = null;
                                worksheet.Cells[row, col].Clear();
                            }
                        }
                    }
                    else
                    {
                        worksheet.Cells[10, 1, startRow + dxfItems.Count + 10, 15].Clear();
                    }

                    // Configurer les en-têtes à la ligne 10
                    worksheet.Cells[10, 1].Value = "Qté pièces CSV/DXF";
                    worksheet.Cells[10, 2].Value = "Tag CSV/DXF";
                    worksheet.Cells[10, 3].Value = "Matériau CSV/DXF";
                    worksheet.Cells[10, 4].Value = "Qté trouvée dans PDF";
                    worksheet.Cells[10, 5].Value = "Résultat d'analyse PDF";

                    // Mettre en forme les en-têtes
                    using (var headerRange = worksheet.Cells[10, 1, 10, 5])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    // Écrire les données
                    for (int i = 0; i < dxfItems.Count; i++)
                    {
                        int currentRow = startRow + i;

                        // Écrire les données dans les colonnes
                        worksheet.Cells[currentRow, 1].Value = dxfItems[i].Quantity;
                        worksheet.Cells[currentRow, 2].Value = dxfItems[i].Tag;
                        worksheet.Cells[currentRow, 3].Value = dxfItems[i].Material;
                        worksheet.Cells[currentRow, 4].Value = dxfItems[i].PdfQuantity;

                        // Écrire le statut
                        string status = "[-] Tag non trouve dans le PDF";
                        Color statusColor = Color.FromArgb(255, 102, 102); // Rouge clair

                        if (dxfItems[i].FoundInPdf)
                        {
                            if (dxfItems[i].Quantity == dxfItems[i].PdfQuantity)
                            {
                                status = "[+] Tag trouve, quantite OK";
                                statusColor = Color.FromArgb(144, 238, 144); // Vert clair
                            }
                            else if (dxfItems[i].PdfQuantity > 0)
                            {
                                status = $"[!] Tag trouve, qty differente: CSV={dxfItems[i].Quantity} vs PDF={dxfItems[i].PdfQuantity}";
                                statusColor = Color.FromArgb(255, 165, 0); // Orange
                            }
                            else
                            {
                                status = "[?] Tag trouve mais quantite = 0 dans PDF";
                                statusColor = Color.FromArgb(173, 216, 230); // Bleu clair
                            }
                        }

                        worksheet.Cells[currentRow, 5].Value = status;

                        // Appliquer la couleur selon le statut
                        using (var rowRange = worksheet.Cells[currentRow, 1, currentRow, 5])
                        {
                            rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rowRange.Style.Fill.BackgroundColor.SetColor(statusColor);

                            // Couleur de police appropriée selon le statut
                            if (status.Contains("[-]") || status.Contains("[!]"))
                            {
                                rowRange.Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                rowRange.Style.Font.Color.SetColor(Color.Black);
                            }
                        }

                        // Nettoyage explicite de la colonne F et au-delà
                        for (int col = 6; col <= 15; col++)
                        {
                            worksheet.Cells[currentRow, col].Value = null;
                            worksheet.Cells[currentRow, col].Clear();
                        }
                    }

                    // Auto-ajuster les colonnes
                    worksheet.Cells.AutoFitColumns();

                    // Sauvegarder le fichier
                    package.Save();

                    Log("FileIO", $"[+] Excel sauvegarde avec succes: {dxfItems.Count} elements ecrits");
                }
            }
            catch (IOException ex)
            {
                Log("Error", $"[-] Erreur E/S lors de l'ecriture dans Excel: {ex.Message}");
                throw new InvalidOperationException("Erreur lors de l'ecriture dans Excel: " + ex.Message, ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                Log("Error", $"[-] Erreur d'acces lors de l'ecriture dans Excel: {ex.Message}");
                throw new UnauthorizedAccessException("Erreur d'acces lors de l'ecriture dans Excel: " + ex.Message, ex);
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] Erreur lors de l'ecriture dans Excel: {ex.Message}");
                throw new InvalidOperationException("Erreur lors de l'ecriture dans Excel: " + ex.Message, ex);
            }
        }

        #endregion

        #region Gestion fichiers Excel avec template

        /// <summary>
        /// Gère intelligemment le fichier Excel avec renommage automatique et copie depuis template
        /// </summary>
        public static string ManageExcelFileWithTemplate(string projectNumber, string reference, string moduleNumber, string modulePath)
        {
            try
            {
                // Nettoyer la référence (enlever "REF" si présent)
                string refNumber = reference;
                if (refNumber.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
                {
                    refNumber = refNumber.Substring(3).PadLeft(2, '0');
                }

                // Nom de fichier cible avec format correct
                string targetFileName = $"{projectNumber}-{refNumber}-{moduleNumber}_Décompte de DXF_DXF Count.xlsx";
                string checkListFileName = $"{projectNumber}-{refNumber}-{moduleNumber}_Liste de vérification_Check List.xlsm";

                // Chemins
                string documentsDir = Path.Combine(modulePath, "0-Documents");
                string targetFilePath = Path.Combine(documentsDir, targetFileName);
                string checkListFilePath = Path.Combine(documentsDir, checkListFileName);

                // Créer le dossier 0-Documents s'il n'existe pas
                if (!Directory.Exists(documentsDir))
                {
                    Directory.CreateDirectory(documentsDir);
                    Log("FileIO", $"[i] Dossier cree: {documentsDir}");
                }

                // Template paths - V1.2 Vault paths
                string templateDir = @"C:\Vault\Engineering\Library\Xnrgy_Module\0-Documents";
                string templateExcelPath = Path.Combine(templateDir, "XXXXX-XX-MXX_Décompte de DXF_DXF Count.xlsx");
                string templateCheckListPath = Path.Combine(templateDir, "XXXXX-XX-MXX_Liste de vérification_Check List.xlsm");

                Log("FileIO", $"[>] Gestion fichier Excel pour: {targetFileName}");

                bool needsProjectCellsUpdate = false;

                // 1. Gérer le fichier Excel principal
                if (!File.Exists(targetFilePath))
                {
                    // Vérifier s'il y a un fichier avec l'ancien format
                    var existingTemplateFiles = Directory.GetFiles(documentsDir, "*_Décompte de DXF_DXF Count.xlsx");

                    if (existingTemplateFiles.Length > 0)
                    {
                        // Renommer le fichier existant
                        string oldFilePath = existingTemplateFiles[0];
                        File.Move(oldFilePath, targetFilePath);
                        needsProjectCellsUpdate = true;
                        Log("FileIO", $"[+] Fichier renomme: {Path.GetFileName(oldFilePath)} -> {targetFileName}");
                    }
                    else if (File.Exists(templateExcelPath))
                    {
                        // Copier depuis le template
                        File.Copy(templateExcelPath, targetFilePath);
                        needsProjectCellsUpdate = true;
                        Log("FileIO", $"[+] Fichier copie depuis template: {targetFileName}");
                    }
                    else
                    {
                        Log("FileIO", $"[!] Template introuvable: {templateExcelPath}");
                    }
                }
                else
                {
                    Log("FileIO", $"[+] Fichier Excel deja existant: {targetFileName}");
                }

                // Mettre à jour les cellules projet si fichier copié/renommé
                if (needsProjectCellsUpdate && File.Exists(targetFilePath))
                {
                    Log("FileIO", "[>] Mise a jour des cellules projet dans le fichier Excel...");
                    try
                    {
                        using (var package = new ExcelPackage(new FileInfo(targetFilePath)))
                        {
                            if (package.Workbook.Worksheets.Count > 0)
                            {
                                var worksheet = package.Workbook.Worksheets[0];
                                UpdateProjectCellsInExcel(worksheet, targetFilePath);
                                package.Save();
                                Log("FileIO", $"[+] Cellules projet mises a jour dans {targetFileName}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("FileIO", $"[!] Erreur mise a jour cellules dans {targetFileName}: {ex.Message}");
                    }
                }

                // 2. Gérer le fichier Check List (bonus)
                if (!File.Exists(checkListFilePath))
                {
                    var existingCheckListFiles = Directory.GetFiles(documentsDir, "*_Liste de vérification_Check List.xlsm");

                    if (existingCheckListFiles.Length > 0)
                    {
                        string oldCheckListPath = existingCheckListFiles[0];
                        File.Move(oldCheckListPath, checkListFilePath);
                        Log("FileIO", $"[+] Check List renommee: {Path.GetFileName(oldCheckListPath)} -> {checkListFileName}");
                    }
                    else if (File.Exists(templateCheckListPath))
                    {
                        File.Copy(templateCheckListPath, checkListFilePath);
                        Log("FileIO", $"[+] Check List copiee depuis template: {checkListFileName}");
                    }
                }

                // Retourner le chemin du fichier Excel principal
                return targetFilePath;
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] Erreur gestion fichier Excel: {ex.Message}");
                return string.Empty;
            }
        }

        #endregion

        #region Méthodes utilitaires

        /// <summary>
        /// Nettoie une valeur CSV
        /// </summary>
        private static string CleanCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            string result = value.Trim();

            // Supprimer les guillemets si présents
            if (result.StartsWith("\"") && result.EndsWith("\""))
            {
                result = result.Substring(1, result.Length - 2);
            }

            return result.Trim();
        }

        /// <summary>
        /// Détecte l'encodage d'un fichier
        /// </summary>
        private static Encoding DetectEncoding(string filePath)
        {
            byte[] buffer = new byte[4];

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                int bytesToRead = Math.Min(4, (int)fs.Length);
                if (bytesToRead > 0)
                {
                    int totalBytesRead = 0;
                    while (totalBytesRead < bytesToRead)
                    {
                        int bytesRead = fs.Read(buffer, totalBytesRead, bytesToRead - totalBytesRead);
                        if (bytesRead == 0) break;
                        totalBytesRead += bytesRead;
                    }
                }
            }

            // UTF-8 BOM: EF BB BF
            if (buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
            {
                return Encoding.UTF8;
            }

            // UTF-32 BOM (BE): 00 00 FE FF
            if (buffer[0] == 0x0 && buffer[1] == 0x0 && buffer[2] == 0xFE && buffer[3] == 0xFF)
            {
                return Encoding.UTF32;
            }

            // UTF-32 BOM (LE): FF FE 00 00
            if (buffer[0] == 0xFF && buffer[1] == 0xFE && buffer[2] == 0x0 && buffer[3] == 0x0)
            {
                return Encoding.UTF32;
            }

            // UTF-16 BOM (BE): FE FF
            if (buffer[0] == 0xFE && buffer[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode;
            }

            // UTF-16 BOM (LE): FF FE
            if (buffer[0] == 0xFF && buffer[1] == 0xFE)
            {
                return Encoding.Unicode;
            }

            // Si aucun BOM n'est détecté, utiliser l'encodage par défaut
            return Encoding.Default;
        }

        /// <summary>
        /// Détecte le séparateur utilisé dans une ligne CSV
        /// </summary>
        private static char DetectSeparator(string line)
        {
            int commaCount = line.Count(c => c == ',');
            int semicolonCount = line.Count(c => c == ';');
            int tabCount = line.Count(c => c == '\t');

            if (commaCount > semicolonCount && commaCount > tabCount)
            {
                return ',';
            }
            else if (semicolonCount > commaCount && semicolonCount > tabCount)
            {
                return ';';
            }
            else
            {
                return '\t';
            }
        }

        private static bool IsHeaderLine(string line)
        {
            return line.Contains("Qtée") || line.Contains("Tag") || line.Contains("QTY") || line.Contains("Qty");
        }

        private static void ProcessCsvLine(string line, int lineNumber, Dictionary<string, PdfAnalyzerService.CsvRow> csvDict)
        {
            try
            {
                char separator = DetectSeparator(line);
                string[] parts = line.Split(new char[] { separator }, StringSplitOptions.None);

                if (parts.Length >= 2)
                {
                    int quantity = 0;
                    string qtyStr = CleanCsvValue(parts[0]);
                    string tag = CleanCsvValue(parts[1]);
                    string material = parts.Length >= 3 ? CleanCsvValue(parts[2]) : "";

                    // Nettoyer le tag
                    if (tag.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase))
                    {
                        tag = tag.Substring(0, tag.Length - 4);
                    }

                    tag = NormalizeTag(tag);

                    if (int.TryParse(qtyStr, out quantity) && quantity > 0)
                    {
                        csvDict.TryAdd(tag, new PdfAnalyzerService.CsvRow(tag, quantity, material));
                    }
                }
            }
            catch (Exception ex)
            {
                Log("FileIO", $"[!] Erreur traitement ligne {lineNumber}: {ex.Message}");
            }
        }

        private static string NormalizeTag(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return string.Empty;
            return tag.ToUpper(CultureInfo.InvariantCulture).Replace("_", "-").Trim();
        }

        /// <summary>
        /// Met à jour automatiquement les cellules projet dans Excel (C1, C2, C3, C4)
        /// </summary>
        private static void UpdateProjectCellsInExcel(ExcelWorksheet worksheet, string excelPath)
        {
            try
            {
                // Extraire les informations du projet depuis le nom de fichier
                string fileName = Path.GetFileNameWithoutExtension(excelPath);

                // Pattern pour extraire: XXXXX-XX-MXX_Décompte de DXF_DXF Count
                string projectPattern = @"^(\d{4,6})-(\d{1,2})-(M\d{1,2})_";
                var match = Regex.Match(fileName, projectPattern, RegexOptions.IgnoreCase);

                if (match.Success)
                {
                    string projectNumber = match.Groups[1].Value;
                    string refNumber = match.Groups[2].Value;
                    string moduleNumber = match.Groups[3].Value;

                    // Extraire le numéro du module (enlever le 'M')
                    string moduleNumericPart = moduleNumber.Substring(1);

                    // Date de l'analyse au format YYYY-MM-DD
                    string analysisDate = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                    Log("FileIO", "[i] Mise a jour cellules Excel:");
                    Log("FileIO", $"  - C1 (Date): {analysisDate}");
                    Log("FileIO", $"  - C2 (Numero Job): {projectNumber}");
                    Log("FileIO", $"  - C3 (Ref): {refNumber}");
                    Log("FileIO", $"  - C4 (Module): {moduleNumericPart}");

                    // Mettre à jour les cellules
                    worksheet.Cells["C1"].Value = analysisDate;
                    worksheet.Cells["C2"].Value = projectNumber;
                    worksheet.Cells["C3"].Value = int.Parse(refNumber, CultureInfo.InvariantCulture);
                    worksheet.Cells["C4"].Value = int.Parse(moduleNumericPart, CultureInfo.InvariantCulture);

                    // Appliquer un formatage approprié
                    worksheet.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C1"].Style.Font.Bold = true;
                    worksheet.Cells["C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C2"].Style.Font.Bold = true;
                    worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C3"].Style.Font.Bold = true;
                    worksheet.Cells["C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C4"].Style.Font.Bold = true;

                    Log("FileIO", "[+] Cellules projet mises a jour automatiquement");
                }
                else
                {
                    // Mettre au moins la date
                    string analysisDate = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    worksheet.Cells["C1"].Value = analysisDate;
                    worksheet.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C1"].Style.Font.Bold = true;

                    Log("FileIO", $"[i] Date d'analyse mise a jour: {analysisDate}");
                }
            }
            catch (Exception ex)
            {
                Log("FileIO", $"[!] Erreur mise a jour cellules projet: {ex.Message}");
            }
        }

        /// <summary>
        /// Ferme automatiquement Excel si le fichier spécifié est ouvert
        /// </summary>
        private static void CloseExcelIfFileIsOpen(string excelPath)
        {
            try
            {
                if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                {
                    return;
                }

                string fileName = Path.GetFileName(excelPath);

                // Vérifier si le fichier est verrouillé
                if (!IsFileInUse(excelPath))
                {
                    return;
                }

                Log("FileIO", $"[!] Fichier Excel verrouille detecte: {fileName}");

                // Obtenir tous les processus Excel en cours
                var excelProcesses = Process.GetProcessesByName("EXCEL");

                if (excelProcesses.Length == 0)
                {
                    Log("FileIO", "[!] Fichier verrouille mais aucun processus Excel trouve");
                    return;
                }

                Log("FileIO", $"[>] {excelProcesses.Length} processus Excel trouves - Fermeture en cours...");

                // Fermeture progressive des processus Excel
                foreach (var excelProcess in excelProcesses)
                {
                    try
                    {
                        if (!excelProcess.HasExited)
                        {
                            Log("FileIO", $"[>] Fermeture du processus Excel PID: {excelProcess.Id}");

                            // Essayer fermeture propre d'abord
                            excelProcess.CloseMainWindow();

                            // Attendre 3 secondes pour fermeture propre
                            if (!excelProcess.WaitForExit(3000))
                            {
                                // Forcer la fermeture si nécessaire
                                excelProcess.Kill();
                                Log("FileIO", $"[!] Processus Excel force a se fermer: PID {excelProcess.Id}");
                            }
                            else
                            {
                                Log("FileIO", $"[+] Processus Excel ferme proprement: PID {excelProcess.Id}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("FileIO", $"[!] Erreur fermeture processus Excel PID {excelProcess.Id}: {ex.Message}");
                    }
                    finally
                    {
                        excelProcess.Dispose();
                    }
                }

                // Attendre que tous les processus se ferment
                Thread.Sleep(2000);

                // Vérification finale
                if (IsFileInUse(excelPath))
                {
                    Log("Error", $"[-] ECHEC: Fichier toujours verrouille apres fermeture Excel: {fileName}");
                    throw new InvalidOperationException($"Impossible de liberer le fichier Excel: {fileName}. Veuillez fermer Excel manuellement et reessayer.");
                }
                else
                {
                    Log("FileIO", $"[+] Fichier Excel libere avec succes: {fileName}");
                }
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] Erreur lors de la fermeture d'Excel: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Vérifie si un fichier est actuellement en cours d'utilisation (verrouillé)
        /// </summary>
        private static bool IsFileInUse(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false; // Fichier non verrouillé
                }
            }
            catch (IOException)
            {
                return true; // Fichier verrouillé
            }
            catch
            {
                return false; // Autres erreurs = considérer comme non verrouillé
            }
        }

        #endregion

        #region Logging

        private static void Log(string category, string message)
        {
            OnLog?.Invoke(category, message);
        }

        #endregion
    }

    /// <summary>
    /// Représente un élément DXF avec ses propriétés
    /// </summary>
    public class DxfItem
    {
        public int Quantity { get; set; }
        public string Tag { get; set; } = "";
        public string Material { get; set; } = "";
        public bool FoundInPdf { get; set; }
        public int PdfQuantity { get; set; }
    }
}
