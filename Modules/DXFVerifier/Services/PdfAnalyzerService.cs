// =============================================================================
// PdfAnalyzerService.cs - Moteur d'analyse PDF pour DXF Verifier
// MIGRATION EXACTE depuis PdfAnalyzer.vb - NE PAS MODIFIER LA LOGIQUE
// Auteur original: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
// Version: 1.2 - Portage C# depuis VB.NET
// =============================================================================
// [!!!] CE CODE A ETE CALIBRE PENDANT 1 MOIS - NE PAS TOUCHER LA LOGIQUE [!!!]
// =============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Services
{
    /// <summary>
    /// Analyseur PDF optimisé pour DXF-CSV vs PDF Verifier v1.2
    /// Utilise exclusivement UglyToad.PdfPig pour des performances optimales
    /// Version portée depuis VB.NET avec logique IDENTIQUE
    /// </summary>
    public static class PdfAnalyzerService
    {
        #region Structures de données

        /// <summary>
        /// Représente un élément de texte avec sa position dans le PDF
        /// </summary>
        private sealed class TextElement
        {
            public string Text { get; set; } = "";
            public double X { get; set; }
            public double Y { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public int PageNumber { get; set; }

            public override string ToString() => $"{Text} @ ({X:F2}, {Y:F2})";
        }

        /// <summary>
        /// Représente une ligne de texte regroupée
        /// </summary>
        private sealed class TextLine
        {
            public List<TextElement> Elements { get; set; } = new List<TextElement>();
            public double Y { get; set; }
            public int PageNumber { get; set; }
            public string Text { get; set; } = "";

            public void BuildText()
            {
                // Trier les éléments par position X et reconstruire le texte
                Elements = Elements.OrderBy(e => e.X).ToList();
                Text = string.Join(" ", Elements.Select(e => e.Text.Trim()));
            }
        }

        /// <summary>
        /// Représente un item extrait du PDF avec plus de détails sur sa source
        /// </summary>
        public class PdfItem
        {
            public string Tag { get; set; } = "";
            public string Material { get; set; } = "";
            public int Quantity { get; set; }
            public int LineNumber { get; set; }
            public int PageNumber { get; set; }
            public double Confidence { get; set; }
            public string SourceType { get; set; } = ""; // "TABLE" ou "ISOLATED" ou "BALLON"
            public int TableNumber { get; set; } // Numéro du tableau si trouvé dans un tableau
        }

        /// <summary>
        /// Structure pour représenter une ligne CSV (compatibilité)
        /// </summary>
        public class CsvRow
        {
            public string Tag { get; set; } = "";
            public int Quantity { get; set; }
            public string Material { get; set; } = "";

            public CsvRow() { }

            public CsvRow(string tag, int quantity, string material)
            {
                Tag = tag;
                Quantity = quantity;
                Material = material;
            }

            public override string ToString() => $"{Tag}: {Quantity} ({Material})";
        }

        #endregion

        #region Champs statiques

        // Variable pour stocker le nombre de pages du dernier PDF analysé
        private static int _lastAnalyzedPageCount;

        /// <summary>
        /// Propriété publique pour obtenir le nombre de pages du dernier PDF analysé
        /// </summary>
        public static int LastAnalyzedPageCount => _lastAnalyzedPageCount;

        // Événement pour le logging (sera connecté au journal de l'UI)
        public static event Action<string, string>? OnLog;

        #endregion

        #region API Publique

        /// <summary>
        /// Point d'entrée principal - extrait tous les tableaux du PDF
        /// Version restaurée avec la logique simple et efficace
        /// </summary>
        /// <param name="pdfPath">Chemin complet du fichier PDF</param>
        /// <returns>Dictionnaire de tags avec leurs quantités</returns>
        public static Dictionary<string, int> ExtractTablesFromPdf(string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(pdfPath))
            {
                return new Dictionary<string, int>();
            }

            Log("PdfAnalysis", "========== DEBUT EXTRACTION PDF ==========");
            Log("PdfAnalysis", $"Fichier: {pdfPath}");

            var results = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var startTime = DateTime.Now;

                // 1. Extraction principale avec la logique simple et efficace
                var extractedItems = ExtractStructuredTables(pdfPath);
                Log("PdfAnalysis", $"Extraction structuree: {extractedItems.Count} items trouves");

                // 2. Conversion en dictionnaire (CORRIGÉE pour préserver quantités 0)
                foreach (var item in extractedItems)
                {
                    if (!string.IsNullOrWhiteSpace(item.Tag) && item.Quantity >= 0) // >= 0 au lieu de > 0
                    {
                        if (results.ContainsKey(item.Tag))
                        {
                            // Ne jamais écraser avec une quantité plus petite (SAUF si c'est 0 explicite)
                            if (item.Quantity > results[item.Tag] ||
                                (item.Quantity == 0 && results[item.Tag] == 1)) // Cas spécial : 0 explicite écrase 1 par défaut
                            {
                                Log("PdfAnalysis", $"[!] Mise a jour quantite pour {item.Tag}: {results[item.Tag]} -> {item.Quantity}");
                                results[item.Tag] = item.Quantity;
                            }
                        }
                        else
                        {
                            results[item.Tag] = item.Quantity; // Inclut maintenant les quantités 0
                        }
                    }
                }

                // 3. Log final
                var duration = DateTime.Now - startTime;
                Log("PdfAnalysis", $"Extraction terminee: {results.Count} tags uniques");
                Log("PdfAnalysis", $"Duree: {duration.TotalSeconds:F2} secondes");
                Log("PdfAnalysis", "========== FIN EXTRACTION PDF ==========");

                return results;
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] ERREUR CRITIQUE: {ex.Message}");
                Log("Error", $"Stack Trace: {ex.StackTrace}");
                return new Dictionary<string, int>();
            }
        }

        /// <summary>
        /// Version simple compatible avec MainForm
        /// AVEC stratégie de recherche CSV → PDF en deux étapes
        /// </summary>
        public static Dictionary<string, int> ExtractTablesFromPdfSimple(string pdfPath, Dictionary<string, CsvRow>? csvReference = null)
        {
            if (string.IsNullOrWhiteSpace(pdfPath))
            {
                return new Dictionary<string, int>();
            }

            try
            {
                Log("PdfAnalysis", "=== Debut extraction PDF Simple ===");
                Log("PdfAnalysis", $"Fichier: {pdfPath}");
                if (csvReference != null)
                {
                    Log("PdfAnalysis", $"Reference CSV: {csvReference.Count} elements");
                }

                // Déléguer à la méthode principale
                var results = ExtractTablesFromPdf(pdfPath);

                // STRATÉGIE CSV → PDF : Corriger les quantités des tags échués avec référence CSV
                if (csvReference != null && csvReference.Count > 0)
                {
                    // Pour chaque tag trouvé avec source BALLON, utiliser la quantité CSV si disponible
                    foreach (var kvp in results.ToList())
                    {
                        if (csvReference.TryGetValue(kvp.Key, out var csvRow))
                        {
                            // Si le tag PDF a une quantité par défaut (1) et qu'on a une référence CSV, utiliser CSV
                            if (kvp.Value == 1 && csvRow.Quantity > 1)
                            {
                                results[kvp.Key] = csvRow.Quantity;
                                Log("PdfAnalysis", $"[+] Quantite corrigee par CSV: {kvp.Key} = {csvRow.Quantity}");
                            }
                        }
                    }

                    // Filtrer les résultats selon la référence CSV si fournie (comme avant)
                    var filteredResults = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    foreach (var kvp in results)
                    {
                        // Si le tag existe dans la référence CSV, l'inclure
                        if (csvReference.ContainsKey(kvp.Key))
                        {
                            filteredResults[kvp.Key] = kvp.Value;
                        }
                    }

                    Log("PdfAnalysis", $"Filtrage: {results.Count} -> {filteredResults.Count} tags");
                    results = filteredResults;
                }

                Log("PdfAnalysis", $"Extraction terminee: {results.Count} tags extraits");
                Log("PdfAnalysis", "=== Fin extraction PDF Simple ===");

                return results;
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] ERREUR dans ExtractTablesFromPdfSimple: {ex.Message}");
                Log("Error", $"Stack Trace: {ex.StackTrace}");
                return new Dictionary<string, int>();
            }
        }

        #endregion

        #region Extraction structurée des tableaux - Logique simple restaurée

        /// <summary>
        /// Extraction principale avec la logique simple et efficace d'hier
        /// Restaurée depuis la version .NET 6 qui fonctionnait parfaitement
        /// </summary>
        private static List<PdfItem> ExtractStructuredTables(string pdfPath)
        {
            var allItems = new List<PdfItem>();
            var foundTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // Pour éviter les doublons

            if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath))
            {
                Log("Error", $"[-] Fichier PDF introuvable: {pdfPath}");
                return allItems;
            }

            Log("PdfAnalysis", "=== Debut extraction structuree (logique simple) ===");

            try
            {
                using (var document = PdfDocument.Open(pdfPath))
                {
                    Log("PdfAnalysis", $"Document ouvert: {document.NumberOfPages} pages");

                    // Stocker le nombre de pages pour exposition publique
                    _lastAnalyzedPageCount = document.NumberOfPages;

                    // Parcourir toutes les pages
                    for (int pageNum = 1; pageNum <= document.NumberOfPages; pageNum++)
                    {
                        var page = document.GetPage(pageNum);

                        // Extraire tous les mots de la page
                        var words = page.GetWords().ToList();

                        // Convertir en TextElements
                        var textElements = new List<TextElement>();
                        foreach (var word in words)
                        {
                            if (!string.IsNullOrWhiteSpace(word.Text))
                            {
                                textElements.Add(new TextElement
                                {
                                    Text = word.Text,
                                    X = word.BoundingBox.Left,
                                    Y = word.BoundingBox.Bottom,
                                    Width = word.BoundingBox.Width,
                                    Height = word.BoundingBox.Height,
                                    PageNumber = pageNum
                                });
                            }
                        }

                        // Grouper en lignes
                        var lines = GroupIntoLines(textElements, pageNum);

                        // Détecter les tableaux (méthode simple)
                        var tables = DetectTables(lines);
                        Log("PdfAnalysis", $"Page {pageNum}: {tables.Count} tableaux detectes");

                        // Extraire les données de chaque tableau
                        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
                        {
                            var table = tables[tableIndex];
                            Log("PdfAnalysis", $"  Tableau {tableIndex + 1}: {table.Count} lignes");

                            // Analyser le tableau (méthode simple)
                            var tableItems = AnalyzeTable(table, pageNum, tableIndex + 1);

                            // Ajouter les items de tableau et marquer les tags comme trouvés
                            foreach (var item in tableItems)
                            {
                                foundTags.Add(item.Tag);
                                allItems.Add(item);
                            }

                            Log("PdfAnalysis", $"  Tableau {tableIndex + 1}: {tableItems.Count} items extraits");
                        }

                        // Analyse complémentaire: chercher des tags isolés hors tableaux (SEULEMENT ceux pas encore trouvés)
                        var isolatedItems = FindIsolatedTags(lines, pageNum, foundTags);
                        if (isolatedItems.Count > 0)
                        {
                            Log("PdfAnalysis", $"  Page {pageNum}: {isolatedItems.Count} tags isoles complementaires trouves");
                            allItems.AddRange(isolatedItems);
                        }
                    }
                }

                Log("PdfAnalysis", $"=== Extraction terminee: {allItems.Count} items total ===");
                Log("PdfAnalysis", $"Tags uniques trouves: {foundTags.Count}");
                return allItems;
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] Erreur dans ExtractStructuredTables: {ex.Message}");
                Log("Error", $"Stack Trace: {ex.StackTrace}");
                return new List<PdfItem>();
            }
        }

        /// <summary>
        /// Groupe les éléments de texte en lignes
        /// </summary>
        private static List<TextLine> GroupIntoLines(List<TextElement> elements, int pageNum)
        {
            var lines = new List<TextLine>();

            if (elements.Count == 0) return lines;

            // Trier par Y décroissant (haut vers bas), puis par X
            var sorted = elements.OrderByDescending(e => e.Y).ThenBy(e => e.X).ToList();

            // Tolérance Y pour considérer des éléments sur la même ligne
            const double Y_TOLERANCE = 2.0;

            var currentLine = new TextLine
            {
                Y = sorted[0].Y,
                PageNumber = pageNum
            };
            currentLine.Elements.Add(sorted[0]);

            for (int i = 1; i < sorted.Count; i++)
            {
                var element = sorted[i];

                // Si l'élément est sur la même ligne (Y proche)
                if (Math.Abs(element.Y - currentLine.Y) <= Y_TOLERANCE)
                {
                    currentLine.Elements.Add(element);
                }
                else
                {
                    // Construire le texte de la ligne actuelle et l'ajouter
                    currentLine.BuildText();
                    lines.Add(currentLine);

                    // Commencer une nouvelle ligne
                    currentLine = new TextLine
                    {
                        Y = element.Y,
                        PageNumber = pageNum
                    };
                    currentLine.Elements.Add(element);
                }
            }

            // Ajouter la dernière ligne
            if (currentLine.Elements.Count > 0)
            {
                currentLine.BuildText();
                lines.Add(currentLine);
            }

            return lines;
        }

        /// <summary>
        /// Détecte les tableaux dans les lignes de texte (méthode simple)
        /// </summary>
        private static List<List<TextLine>> DetectTables(List<TextLine> lines)
        {
            var tables = new List<List<TextLine>>();
            List<TextLine>? currentTable = null;

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];

                // Détecter un en-tête de tableau (méthode simple)
                if (IsTableHeader(line.Text))
                {
                    // Si un tableau est en cours, le sauvegarder
                    if (currentTable != null && currentTable.Count >= 2)
                    {
                        tables.Add(currentTable);
                    }

                    // Commencer un nouveau tableau
                    currentTable = new List<TextLine> { line };
                }
                else if (currentTable != null)
                {
                    // Vérifier si cette ligne fait partie du tableau
                    if (IsTableDataRow(line.Text))
                    {
                        currentTable.Add(line);
                    }
                    else
                    {
                        // Fin du tableau
                        if (currentTable.Count >= 2)
                        {
                            tables.Add(currentTable);
                        }
                        currentTable = null;
                    }
                }
            }

            // Ajouter le dernier tableau si nécessaire
            if (currentTable != null && currentTable.Count >= 2)
            {
                tables.Add(currentTable);
            }

            return tables;
        }

        /// <summary>
        /// Vérifie si une ligne est un en-tête de tableau (méthode simple)
        /// </summary>
        private static bool IsTableHeader(string lineText)
        {
            if (string.IsNullOrWhiteSpace(lineText)) return false;

            var lower = lineText.ToLowerInvariant();

            // Chercher les mots-clés d'en-tête (variations possibles)
            bool hasTag = lower.Contains("tag") || lower.Contains("ref") ||
                         lower.Contains("part") || lower.Contains("no.");

            bool hasQty = lower.Contains("qty") || lower.Contains("qte") ||
                         lower.Contains("qtee") || lower.Contains("qté") ||
                         lower.Contains("quantité") || lower.Contains("quantite");

            // Un en-tête valide doit avoir au moins Tag et Qty
            return hasTag && hasQty;
        }

        /// <summary>
        /// Vérifie si une ligne contient des données de tableau (méthode simple)
        /// </summary>
        private static bool IsTableDataRow(string lineText)
        {
            if (string.IsNullOrWhiteSpace(lineText)) return false;

            // Pattern pour un tag valide: XXX0000-0000
            string tagPattern = @"[A-Z]{2,3}\d{1,4}[-_]?\d{1,4}";
            bool hasTag = Regex.IsMatch(lineText, tagPattern, RegexOptions.IgnoreCase);

            // Doit avoir au moins un nombre (quantité potentielle)
            bool hasNumber = Regex.IsMatch(lineText, @"\d+");

            return hasTag && hasNumber;
        }

        /// <summary>
        /// Analyse un tableau pour extraire les données (méthode simple restaurée)
        /// </summary>
        private static List<PdfItem> AnalyzeTable(List<TextLine> table, int pageNum, int tableNum)
        {
            var items = new List<PdfItem>();

            if (table.Count < 2) return items;

            // Analyser l'en-tête pour comprendre la structure
            var headerLine = table[0];
            var columnPositions = AnalyzeHeaderColumns(headerLine);

            // Parcourir les lignes de données (sauter l'en-tête)
            for (int i = 1; i < table.Count; i++)
            {
                var dataLine = table[i];
                var item = ExtractItemFromLine(dataLine, columnPositions);

                if (item != null)
                {
                    item.PageNumber = pageNum;
                    item.LineNumber = i;
                    item.SourceType = "TABLE";
                    item.TableNumber = tableNum;
                    items.Add(item);

                    Log("PdfAnalysis", $"    [+] Tag={item.Tag}, Qty={item.Quantity}, Mat={item.Material}");
                }
            }

            return items;
        }

        /// <summary>
        /// Analyse l'en-tête pour déterminer les positions des colonnes (méthode simple)
        /// </summary>
        private static Dictionary<string, double> AnalyzeHeaderColumns(TextLine headerLine)
        {
            var columns = new Dictionary<string, double>();

            if (headerLine?.Elements == null) return columns;

            foreach (var element in headerLine.Elements)
            {
                if (string.IsNullOrWhiteSpace(element.Text)) continue;

                var lower = element.Text.ToLowerInvariant();

                if (lower.Contains("tag") || lower.Contains("ref"))
                {
                    columns["TAG"] = element.X;
                }
                else if (lower.Contains("qty") || lower.Contains("qte") || lower.Contains("qtee"))
                {
                    columns["QTY"] = element.X;
                }
                else if (lower.Contains("material") || lower.Contains("materiel") || lower.Contains("mat"))
                {
                    columns["MATERIAL"] = element.X;
                }
            }

            // Si certaines colonnes manquent, essayer de les déduire
            if (!columns.ContainsKey("TAG") && headerLine.Elements.Count > 0)
            {
                columns["TAG"] = headerLine.Elements[0].X;
            }

            if (!columns.ContainsKey("QTY") && headerLine.Elements.Count > 1)
            {
                // La quantité est généralement après le tag
                columns["QTY"] = headerLine.Elements[1].X;
            }

            return columns;
        }

        /// <summary>
        /// Extrait un item d'une ligne de données (méthode simple restaurée)
        /// MODIFICATION: Préserver les quantités 0 explicites trouvées dans les tableaux
        /// </summary>
        private static PdfItem? ExtractItemFromLine(TextLine dataLine, Dictionary<string, double> columnPositions)
        {
            if (dataLine == null || string.IsNullOrWhiteSpace(dataLine.Text)) return null;

            // Parser la ligne selon les règles strictes et simples
            var lineText = dataLine.Text;

            // 1. Chercher le tag (premier mot qui match le pattern)
            string tagPattern = @"([A-Z]{2,3}\d{1,4}[-_]?\d{1,4})";
            var tagMatch = Regex.Match(lineText, tagPattern, RegexOptions.IgnoreCase);

            if (!tagMatch.Success) return null;

            var item = new PdfItem
            {
                Tag = tagMatch.Groups[1].Value.ToUpper(CultureInfo.InvariantCulture).Replace("_", "-")
            };

            // 2. Chercher la première valeur numérique après le tag = Qty
            var afterTag = lineText.Substring(tagMatch.Index + tagMatch.Length);
            var qtyMatch = Regex.Match(afterTag, @"^\s+(\d+)");

            // Variable pour distinguer quantité trouvée vs non trouvée
            bool quantityFound = false;

            if (qtyMatch.Success)
            {
                if (int.TryParse(qtyMatch.Groups[1].Value, out int qty))
                {
                    item.Quantity = qty;
                    quantityFound = true;
                    // Quantité extraite avec succès (peut être 0, 1, 2, etc.)
                }
            }

            // 3. Tout ce qui est entre le tag et la quantité = Material
            if (qtyMatch.Success)
            {
                int materialStart = tagMatch.Index + tagMatch.Length;
                int materialEnd = tagMatch.Index + tagMatch.Length + qtyMatch.Index;

                if (materialEnd > materialStart)
                {
                    item.Material = lineText.Substring(materialStart, materialEnd - materialStart).Trim();
                    // Ignorer les chiffres dans le matériau selon les règles
                    item.Material = Regex.Replace(item.Material, @"\d+", "").Trim();
                }
            }

            // 4. Validation modifiée pour préserver les quantités 0 explicites
            if (quantityFound)
            {
                // Si une quantité a été trouvée dans le PDF, la respecter (même si c'est 0)
                if (item.Quantity < 0 || item.Quantity > 1000)
                {
                    // Seules les quantités négatives ou aberrantes sont remplacées
                    item.Quantity = 1; // Valeur par défaut pour quantités invalides
                    Log("PdfAnalysis", $"[!] Quantite aberrante corrigee pour {item.Tag}: remplacee par 1");
                }
                else if (item.Quantity == 0)
                {
                    // Quantité 0 explicite préservée
                    Log("PdfAnalysis", $"[i] Quantite 0 explicite preservee pour {item.Tag}");
                }
            }
            else
            {
                // Aucune quantité trouvée = valeur par défaut
                item.Quantity = 1;
            }

            // 5. Calculer la confiance (ajustée pour quantités 0)
            item.Confidence = 0.0;
            if (!string.IsNullOrWhiteSpace(item.Tag)) item.Confidence += 0.5;
            if (quantityFound)
            {
                item.Confidence += 0.3; // Même confiance que la quantité soit 0 ou autre
            }
            else
            {
                item.Confidence += 0.1; // Confiance réduite si quantité par défaut
            }
            if (!string.IsNullOrWhiteSpace(item.Material)) item.Confidence += 0.2;

            return item.Confidence >= 0.5 ? item : null;
        }

        /// <summary>
        /// Cherche des tags isolés hors des tableaux (méthode simple)
        /// SEULEMENT pour les tags pas encore trouvés dans les tableaux
        /// AVEC recherche complémentaire pour tags échués (ballons, cartouches, etc.)
        /// </summary>
        private static List<PdfItem> FindIsolatedTags(List<TextLine> lines, int pageNum, HashSet<string> foundTags)
        {
            var items = new List<PdfItem>();

            foreach (var line in lines)
            {
                // Ignorer les en-têtes de tableau
                if (IsTableHeader(line.Text)) continue;

                // STRATÉGIE 1: Chercher des tags avec pattern standard (Tag + Quantité)
                string tagPattern = @"([A-Z]{2,3}\d{1,4}[-_]?\d{1,4})\s+(\d+)";
                var matches = Regex.Matches(line.Text, tagPattern, RegexOptions.IgnoreCase);

                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        var tagNormalized = match.Groups[1].Value.ToUpper(CultureInfo.InvariantCulture).Replace("_", "-");

                        // SEULEMENT si le tag n'a pas déjà été trouvé dans un tableau
                        if (!foundTags.Contains(tagNormalized))
                        {
                            var item = new PdfItem
                            {
                                Tag = tagNormalized,
                                PageNumber = pageNum,
                                Confidence = 0.7,
                                SourceType = "ISOLATED",
                                TableNumber = 0
                            };

                            if (int.TryParse(match.Groups[2].Value, out int qty) &&
                                qty > 0 && qty <= 1000)
                            {
                                item.Quantity = qty;
                                items.Add(item);
                                foundTags.Add(tagNormalized); // Marquer comme trouvé
                            }
                        }
                    }
                }

                // STRATÉGIE 2: Chercher des tags échués SANS quantité (ballons, cartouches, etc.)
                // Pattern pour tag seul (sera traité plus tard avec quantité CSV de référence)
                string tagOnlyPattern = @"([A-Z]{2,3}\d{1,4}[-_]?\d{1,4})";
                var tagOnlyMatches = Regex.Matches(line.Text, tagOnlyPattern, RegexOptions.IgnoreCase);

                foreach (Match tagMatch in tagOnlyMatches)
                {
                    if (tagMatch.Success)
                    {
                        var tagNormalized = tagMatch.Groups[1].Value.ToUpper(CultureInfo.InvariantCulture).Replace("_", "-");

                        // SEULEMENT si le tag n'a pas déjà été trouvé et ne fait pas partie d'un pattern tag+quantité
                        if (!foundTags.Contains(tagNormalized))
                        {
                            // Vérifier que ce n'est pas déjà inclus dans un pattern tag+quantité de la même ligne
                            bool isAlreadyInTagQtyPattern = matches.Cast<Match>()
                                .Any(m => m.Groups[1].Value.ToUpper(CultureInfo.InvariantCulture).Replace("_", "-") == tagNormalized);

                            if (!isAlreadyInTagQtyPattern)
                            {
                                // Créer un item "échoué" avec quantité par défaut 1 (sera remplacée par CSV si référence fournie)
                                var item = new PdfItem
                                {
                                    Tag = tagNormalized,
                                    PageNumber = pageNum,
                                    Quantity = 1, // Quantité par défaut pour tags échués
                                    Confidence = 0.5, // Confiance réduite pour tags sans quantité
                                    SourceType = "BALLON", // Nouveau type de source
                                    TableNumber = 0
                                };

                                items.Add(item);
                                foundTags.Add(tagNormalized); // Marquer comme trouvé
                            }
                        }
                    }
                }
            }

            return items;
        }

        #endregion

        #region Logging

        private static void Log(string category, string message)
        {
            OnLog?.Invoke(category, message);
        }

        #endregion
    }
}
