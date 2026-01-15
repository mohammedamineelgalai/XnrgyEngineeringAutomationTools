// =============================================================================
// PDF TABLE EXTRACTION BENCHMARK - XNRGY DXF Verifier
// =============================================================================
// Objectif: Tester plusieurs moteurs/méthodes d'extraction de tableaux PDF
// Référence: CSV = vérité terrain (181 tags)
// PDF: 02-Machines.pdf
// =============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace XnrgyEngineeringAutomationTools.Tests
{
    /// <summary>
    /// Benchmark de différentes méthodes d'extraction de tableaux PDF
    /// </summary>
    public class PdfExtractionBenchmark
    {
        // Fichiers de référence
        private const string PDF_PATH = @"C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines.pdf";
        private const string CSV_PATH = @"C:\Vault\Engineering\Projects\10381\REF13\M02\5_Exportation\Sheet_Metal_Nesting\Punch\10381-13-M02.csv";

        // Pattern pour les tags XNRGY
        private static readonly Regex TagPattern = new Regex(
            @"([A-Z]{2,4}\d{3,4}[-_]\d{3,4})",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Résultats de référence (CSV)
        private HashSet<string> _referenceTagsCSV = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Point d'entrée du benchmark
        /// </summary>
        public static void RunBenchmark()
        {
            var benchmark = new PdfExtractionBenchmark();
            benchmark.Execute();
        }

        /// <summary>
        /// Exécute tous les tests de benchmark
        /// </summary>
        public void Execute()
        {
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine("PDF TABLE EXTRACTION BENCHMARK - XNRGY DXF Verifier");
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine();

            // 1. Charger la référence CSV
            LoadCSVReference();

            // 2. Tester chaque méthode d'extraction
            var results = new List<BenchmarkResult>();

            results.Add(TestMethod1_SimpleLineByLine());
            results.Add(TestMethod2_WordExtractor_Default());
            results.Add(TestMethod3_WordExtractor_NearestNeighbour());
            results.Add(TestMethod4_RecursiveXYCut());
            results.Add(TestMethod5_DocstrumBounding());
            results.Add(TestMethod6_PhysicalLayout());
            results.Add(TestMethod7_GridBasedReconstruction());
            results.Add(TestMethod8_ColumnPositionAnalysis());
            results.Add(TestMethod9_YCoordinateClustering());
            results.Add(TestMethod10_CombinedApproach());

            // 3. Afficher le résumé comparatif
            PrintSummary(results);

            // 4. Analyser les tags manquants pour la meilleure méthode
            AnalyzeMissingTags(results);
        }

        /// <summary>
        /// Charge les tags de référence depuis le CSV
        /// </summary>
        private void LoadCSVReference()
        {
            Console.WriteLine("[>] Chargement de la reference CSV...");

            if (!File.Exists(CSV_PATH))
            {
                Console.WriteLine($"[-] ERREUR: CSV non trouve: {CSV_PATH}");
                return;
            }

            var lines = File.ReadAllLines(CSV_PATH);
            foreach (var line in lines.Skip(1)) // Skip header
            {
                var parts = line.Split(';', ',');
                if (parts.Length > 0)
                {
                    var tag = parts[0].Trim().ToUpperInvariant().Replace("_", "-");
                    if (TagPattern.IsMatch(tag))
                    {
                        _referenceTagsCSV.Add(tag);
                    }
                }
            }

            Console.WriteLine($"[+] Reference CSV: {_referenceTagsCSV.Count} tags uniques");
            Console.WriteLine();
        }

        #region Test Methods

        /// <summary>
        /// Méthode 1: Extraction ligne par ligne simple (baseline)
        /// </summary>
        private BenchmarkResult TestMethod1_SimpleLineByLine()
        {
            var methodName = "Method 1: Simple Line-by-Line";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var text = page.Text;
                        var matches = TagPattern.Matches(text);

                        foreach (Match match in matches)
                        {
                            var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                            extractedTags.Add(tag);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 2: WordExtractor par défaut
        /// </summary>
        private BenchmarkResult TestMethod2_WordExtractor_Default()
        {
            var methodName = "Method 2: WordExtractor Default";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var words = page.GetWords();

                        foreach (var word in words)
                        {
                            if (TagPattern.IsMatch(word.Text))
                            {
                                var tag = TagPattern.Match(word.Text).Groups[1].Value
                                    .ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 3: WordExtractor NearestNeighbour
        /// </summary>
        private BenchmarkResult TestMethod3_WordExtractor_NearestNeighbour()
        {
            var methodName = "Method 3: WordExtractor NearestNeighbour";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var words = NearestNeighbourWordExtractor.Instance.GetWords(page.Letters);

                        foreach (var word in words)
                        {
                            if (TagPattern.IsMatch(word.Text))
                            {
                                var tag = TagPattern.Match(word.Text).Groups[1].Value
                                    .ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 4: RecursiveXYCut Page Segmenter
        /// </summary>
        private BenchmarkResult TestMethod4_RecursiveXYCut()
        {
            var methodName = "Method 4: RecursiveXYCut Segmenter";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var words = page.GetWords();
                        var blocks = RecursiveXYCut.Instance.GetBlocks(words);

                        foreach (var block in blocks)
                        {
                            var blockText = string.Join(" ", block.TextLines.SelectMany(l => l.Words).Select(w => w.Text));
                            var matches = TagPattern.Matches(blockText);

                            foreach (Match match in matches)
                            {
                                var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 5: DocstrumBoundingBoxes Segmenter
        /// </summary>
        private BenchmarkResult TestMethod5_DocstrumBounding()
        {
            var methodName = "Method 5: DocstrumBoundingBoxes";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var words = page.GetWords();
                        var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(words);

                        foreach (var block in blocks)
                        {
                            var blockText = string.Join(" ", block.TextLines.SelectMany(l => l.Words).Select(w => w.Text));
                            var matches = TagPattern.Matches(blockText);

                            foreach (Match match in matches)
                            {
                                var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 6: Physical Layout Analysis
        /// </summary>
        private BenchmarkResult TestMethod6_PhysicalLayout()
        {
            var methodName = "Method 6: Physical Layout Analysis";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        // Grouper les lettres par position Y (lignes)
                        var letters = page.Letters.ToList();
                        var lineGroups = letters
                            .GroupBy(l => Math.Round(l.GlyphRectangle.Bottom, 1))
                            .OrderByDescending(g => g.Key);

                        foreach (var lineGroup in lineGroups)
                        {
                            // Trier les lettres par X (gauche à droite)
                            var sortedLetters = lineGroup.OrderBy(l => l.GlyphRectangle.Left);
                            var lineText = string.Concat(sortedLetters.Select(l => l.Value));

                            var matches = TagPattern.Matches(lineText);
                            foreach (Match match in matches)
                            {
                                var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 7: Grid-Based Table Reconstruction
        /// </summary>
        private BenchmarkResult TestMethod7_GridBasedReconstruction()
        {
            var methodName = "Method 7: Grid-Based Reconstruction";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var words = page.GetWords().ToList();

                        // Détecter les colonnes (positions X uniques)
                        var xPositions = words
                            .Select(w => Math.Round(w.BoundingBox.Left, 0))
                            .Distinct()
                            .OrderBy(x => x)
                            .ToList();

                        // Détecter les lignes (positions Y uniques)
                        var yPositions = words
                            .Select(w => Math.Round(w.BoundingBox.Bottom, 0))
                            .Distinct()
                            .OrderByDescending(y => y)
                            .ToList();

                        // Pour chaque cellule de la grille
                        foreach (var y in yPositions)
                        {
                            var rowWords = words
                                .Where(w => Math.Abs(Math.Round(w.BoundingBox.Bottom, 0) - y) < 5)
                                .OrderBy(w => w.BoundingBox.Left)
                                .Select(w => w.Text);

                            var rowText = string.Join(" ", rowWords);

                            var matches = TagPattern.Matches(rowText);
                            foreach (Match match in matches)
                            {
                                var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 8: Column Position Analysis (détection de structure de tableau)
        /// </summary>
        private BenchmarkResult TestMethod8_ColumnPositionAnalysis()
        {
            var methodName = "Method 8: Column Position Analysis";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var words = page.GetWords().ToList();

                        // Trouver les mots qui sont des tags
                        var tagWords = words.Where(w => TagPattern.IsMatch(w.Text)).ToList();

                        // Pour chaque tag trouvé, chercher la quantité associée
                        foreach (var tagWord in tagWords)
                        {
                            var tag = TagPattern.Match(tagWord.Text).Groups[1].Value
                                .ToUpperInvariant().Replace("_", "-");

                            // Chercher si c'est dans un contexte de tableau (mots alignés)
                            var sameRowWords = words
                                .Where(w => Math.Abs(w.BoundingBox.Bottom - tagWord.BoundingBox.Bottom) < 5)
                                .OrderBy(w => w.BoundingBox.Left)
                                .ToList();

                            // Si on a plusieurs mots sur la même ligne = probablement un tableau
                            if (sameRowWords.Count >= 2)
                            {
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 9: Y-Coordinate Clustering avec tolérance
        /// </summary>
        private BenchmarkResult TestMethod9_YCoordinateClustering()
        {
            var methodName = "Method 9: Y-Coordinate Clustering";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        var words = page.GetWords().ToList();

                        // Clustering des mots par Y avec tolérance de 3 points
                        var clusters = new List<List<Word>>();
                        var tolerance = 3.0;

                        foreach (var word in words.OrderByDescending(w => w.BoundingBox.Bottom))
                        {
                            var y = word.BoundingBox.Bottom;
                            var existingCluster = clusters.FirstOrDefault(c =>
                                c.Any(w => Math.Abs(w.BoundingBox.Bottom - y) < tolerance));

                            if (existingCluster != null)
                            {
                                existingCluster.Add(word);
                            }
                            else
                            {
                                clusters.Add(new List<Word> { word });
                            }
                        }

                        // Analyser chaque cluster (ligne)
                        foreach (var cluster in clusters)
                        {
                            var lineText = string.Join(" ", cluster.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text));

                            var matches = TagPattern.Matches(lineText);
                            foreach (Match match in matches)
                            {
                                var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        /// <summary>
        /// Méthode 10: Approche combinée (meilleure des précédentes)
        /// </summary>
        private BenchmarkResult TestMethod10_CombinedApproach()
        {
            var methodName = "Method 10: Combined Approach";
            Console.WriteLine($"[>] Test: {methodName}");

            var extractedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var document = PdfDocument.Open(PDF_PATH))
                {
                    foreach (var page in document.GetPages())
                    {
                        // Approche 1: Extraction directe des mots
                        var words = page.GetWords().ToList();
                        foreach (var word in words)
                        {
                            if (TagPattern.IsMatch(word.Text))
                            {
                                var tag = TagPattern.Match(word.Text).Groups[1].Value
                                    .ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }

                        // Approche 2: Reconstruction des lignes par position Y
                        var lineGroups = words
                            .GroupBy(w => Math.Round(w.BoundingBox.Bottom / 3.0) * 3.0)
                            .OrderByDescending(g => g.Key);

                        foreach (var lineGroup in lineGroups)
                        {
                            var lineText = string.Join("", lineGroup.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text + " "));

                            var matches = TagPattern.Matches(lineText);
                            foreach (Match match in matches)
                            {
                                var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                                extractedTags.Add(tag);
                            }
                        }

                        // Approche 3: Texte brut de la page
                        var pageText = page.Text;
                        var pageMatches = TagPattern.Matches(pageText);
                        foreach (Match match in pageMatches)
                        {
                            var tag = match.Groups[1].Value.ToUpperInvariant().Replace("_", "-");
                            extractedTags.Add(tag);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[-] Erreur: {ex.Message}");
            }

            return EvaluateResult(methodName, extractedTags);
        }

        #endregion

        #region Evaluation

        /// <summary>
        /// Évalue les résultats d'une méthode
        /// </summary>
        private BenchmarkResult EvaluateResult(string methodName, HashSet<string> extractedTags)
        {
            var result = new BenchmarkResult
            {
                MethodName = methodName,
                ExtractedTags = extractedTags,
                TotalExtracted = extractedTags.Count,
                TotalReference = _referenceTagsCSV.Count
            };

            // Tags trouvés qui sont dans la référence (True Positives)
            result.CorrectTags = extractedTags.Intersect(_referenceTagsCSV, StringComparer.OrdinalIgnoreCase).ToHashSet();
            result.TruePositives = result.CorrectTags.Count;

            // Tags manquants (False Negatives)
            result.MissingTags = _referenceTagsCSV.Except(extractedTags, StringComparer.OrdinalIgnoreCase).ToHashSet();
            result.FalseNegatives = result.MissingTags.Count;

            // Tags en trop (False Positives)
            result.ExtraTags = extractedTags.Except(_referenceTagsCSV, StringComparer.OrdinalIgnoreCase).ToHashSet();
            result.FalsePositives = result.ExtraTags.Count;

            // Métriques
            result.Precision = result.TotalExtracted > 0
                ? (double)result.TruePositives / result.TotalExtracted * 100
                : 0;

            result.Recall = result.TotalReference > 0
                ? (double)result.TruePositives / result.TotalReference * 100
                : 0;

            result.F1Score = (result.Precision + result.Recall) > 0
                ? 2 * (result.Precision * result.Recall) / (result.Precision + result.Recall)
                : 0;

            // Afficher le résultat
            Console.WriteLine($"    Extraits: {result.TotalExtracted} | Corrects: {result.TruePositives} | " +
                            $"Manquants: {result.FalseNegatives} | En trop: {result.FalsePositives}");
            Console.WriteLine($"    Precision: {result.Precision:F1}% | Recall: {result.Recall:F1}% | F1: {result.F1Score:F1}%");
            Console.WriteLine();

            return result;
        }

        /// <summary>
        /// Affiche le résumé comparatif
        /// </summary>
        private void PrintSummary(List<BenchmarkResult> results)
        {
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine("RESUME COMPARATIF");
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine();
            Console.WriteLine($"Reference CSV: {_referenceTagsCSV.Count} tags uniques");
            Console.WriteLine();

            // Trier par F1 Score
            var sorted = results.OrderByDescending(r => r.F1Score).ToList();

            Console.WriteLine($"{"Rang",-5} {"Methode",-40} {"Recall",-10} {"Precision",-10} {"F1",-10} {"Manquants",-10}");
            Console.WriteLine("-".PadRight(95, '-'));

            int rank = 1;
            foreach (var result in sorted)
            {
                Console.WriteLine($"{rank,-5} {result.MethodName,-40} {result.Recall,7:F1}% {result.Precision,9:F1}% {result.F1Score,7:F1}% {result.FalseNegatives,8}");
                rank++;
            }

            Console.WriteLine();
            Console.WriteLine($"[+] MEILLEURE METHODE: {sorted[0].MethodName}");
            Console.WriteLine($"    Recall: {sorted[0].Recall:F1}% ({sorted[0].TruePositives}/{sorted[0].TotalReference} tags trouves)");
            Console.WriteLine($"    Precision: {sorted[0].Precision:F1}%");
            Console.WriteLine($"    Tags manquants: {sorted[0].FalseNegatives}");
            Console.WriteLine();
        }

        /// <summary>
        /// Analyse détaillée des tags manquants
        /// </summary>
        private void AnalyzeMissingTags(List<BenchmarkResult> results)
        {
            var best = results.OrderByDescending(r => r.F1Score).First();

            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine("ANALYSE DES TAGS MANQUANTS (Meilleure methode)");
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine();

            if (best.MissingTags.Count == 0)
            {
                Console.WriteLine("[+] Aucun tag manquant! 100% des tags trouves.");
                return;
            }

            Console.WriteLine($"Tags manquants ({best.MissingTags.Count}):");
            foreach (var tag in best.MissingTags.OrderBy(t => t))
            {
                Console.WriteLine($"  - {tag}");
            }

            // Analyser les patterns des tags manquants
            Console.WriteLine();
            Console.WriteLine("Analyse des patterns:");

            var prefixes = best.MissingTags
                .Select(t => Regex.Match(t, @"^([A-Z]+)").Groups[1].Value)
                .GroupBy(p => p)
                .OrderByDescending(g => g.Count());

            foreach (var prefix in prefixes)
            {
                Console.WriteLine($"  Prefixe '{prefix.Key}': {prefix.Count()} tags manquants");
            }

            // Tester si les tags manquants existent dans le PDF brut
            Console.WriteLine();
            Console.WriteLine("Verification dans le PDF brut:");

            using (var document = PdfDocument.Open(PDF_PATH))
            {
                var fullText = string.Join("\n", document.GetPages().Select(p => p.Text));

                int foundInRaw = 0;
                foreach (var tag in best.MissingTags)
                {
                    if (fullText.Contains(tag, StringComparison.OrdinalIgnoreCase))
                    {
                        foundInRaw++;
                        Console.WriteLine($"  [+] '{tag}' existe dans le PDF brut");
                    }
                    else
                    {
                        Console.WriteLine($"  [-] '{tag}' N'EXISTE PAS dans le PDF brut");
                    }
                }

                Console.WriteLine();
                Console.WriteLine($"Resultat: {foundInRaw}/{best.MissingTags.Count} tags manquants sont dans le PDF brut");
                Console.WriteLine($"=> {best.MissingTags.Count - foundInRaw} tags sont vraiment absents du PDF");
            }
        }

        #endregion
    }

    /// <summary>
    /// Résultat d'un test de benchmark
    /// </summary>
    public class BenchmarkResult
    {
        public string MethodName { get; set; } = "";
        public HashSet<string> ExtractedTags { get; set; } = new HashSet<string>();
        public HashSet<string> CorrectTags { get; set; } = new HashSet<string>();
        public HashSet<string> MissingTags { get; set; } = new HashSet<string>();
        public HashSet<string> ExtraTags { get; set; } = new HashSet<string>();

        public int TotalExtracted { get; set; }
        public int TotalReference { get; set; }
        public int TruePositives { get; set; }
        public int FalseNegatives { get; set; }
        public int FalsePositives { get; set; }

        public double Precision { get; set; }
        public double Recall { get; set; }
        public double F1Score { get; set; }
    }
}
