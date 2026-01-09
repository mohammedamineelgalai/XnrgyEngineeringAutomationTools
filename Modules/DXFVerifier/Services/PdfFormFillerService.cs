// =============================================================================
// PdfFormFillerService.cs - Remplisseur de formulaire PDF pour DXF Verifier
// MIGRATION EXACTE depuis PdfFormFiller.vb - NE PAS MODIFIER LA LOGIQUE
// Auteur original: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
// Version: 1.2 - Portage C# depuis VB.NET
// =============================================================================
// Methode: Modifier /V des annotations + NeedAppearances=true pour Adobe
// =============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using iTextSharp.text.pdf;

namespace XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Services
{
    /// <summary>
    /// Module pour remplir les champs de formulaire PDF avec les quantités de vérification DXF
    /// Utilise iTextSharp.LGPLv2.Core pour modifier les widgets de la page de couverture BatchPrint
    /// Méthode: Modifier /V des annotations + NeedAppearances=true pour que Adobe régénère l'affichage
    /// </summary>
    public static class PdfFormFillerService
    {
        // Mapping des noms de champs vers leurs index dans les annotations de page 1
        // Index 4=Texte1, 5=Texte2, 6=Texte3, 7=Texte4, 8=Texte5
        private static readonly Dictionary<string, int> FieldIndexes = new Dictionary<string, int>
        {
            { "Texte1", 4 },
            { "Texte2", 5 },
            { "Texte3", 6 },
            { "Texte4", 7 },
            { "Texte5", 8 }
        };

        // Événement pour le logging (sera connecté au journal de l'UI)
        public static event Action<string, string> OnLog;

        /// <summary>
        /// Remplit les champs de quantités dans le PDF 02-Machines.pdf
        /// </summary>
        public static bool FillQuantityFields(string pdfPath,
                                              int totalCsvQty,
                                              int totalCsvTags,
                                              int totalPdfQty,
                                              int totalPdfTags)
        {
            string? tempPath = null;

            try
            {
                if (string.IsNullOrEmpty(pdfPath))
                {
                    Log("Error", "[-] PdfFormFiller: Chemin PDF vide ou null");
                    return false;
                }

                if (!File.Exists(pdfPath))
                {
                    Log("Error", $"[-] PdfFormFiller: Fichier PDF introuvable: {pdfPath}");
                    return false;
                }

                Log("PdfAnalysis", $"[>] PdfFormFiller: Remplissage des quantites dans {Path.GetFileName(pdfPath)}");

                // Créer un fichier temporaire
                string? directory = Path.GetDirectoryName(pdfPath);
                if (string.IsNullOrEmpty(directory))
                {
                    Log("Error", "[-] PdfFormFiller: Impossible d'obtenir le repertoire du PDF");
                    return false;
                }
                tempPath = Path.Combine(directory, $"temp_{Guid.NewGuid():N}.pdf");

                // Copier l'original vers temp pour travailler dessus
                File.Copy(pdfPath, tempPath, true);

                int filledCount = 0;
                int totalPages = 0;

                // Ouvrir le PDF temp en lecture, écrire vers l'original
                using (var reader = new PdfReader(tempPath))
                using (var fs = new FileStream(pdfPath, FileMode.Create, FileAccess.Write))
                using (var stamper = new PdfStamper(reader, fs))
                {
                    totalPages = reader.NumberOfPages;
                    Log("PdfAnalysis", $"[i] PDF contient {totalPages} pages");

                    // Obtenir les annotations de la page 1 (couverture)
                    var page1 = reader.GetPageN(1);
                    var annots = page1.GetAsArray(PdfName.ANNOTS);

                    if (annots == null || annots.Size < 9)
                    {
                        Log("PdfAnalysis", "[!] PDF n'a pas assez d'annotations sur la page 1");
                        return false;
                    }

                    // Valeurs à remplir
                    var values = new Dictionary<int, string>
                    {
                        { 4, totalCsvQty.ToString(CultureInfo.InvariantCulture) },
                        { 5, totalCsvQty.ToString(CultureInfo.InvariantCulture) },
                        { 6, totalCsvTags.ToString(CultureInfo.InvariantCulture) },
                        { 7, totalCsvTags.ToString(CultureInfo.InvariantCulture) },
                        { 8, totalPages.ToString(CultureInfo.InvariantCulture) }
                    };

                    // Modifier chaque annotation
                    foreach (var kvp in values)
                    {
                        int idx = kvp.Key;
                        string val = kvp.Value;

                        try
                        {
                            var annRef = annots.GetAsIndirectObject(idx);
                            var ann = PdfReader.GetPdfObject(annRef);

                            if (ann is PdfDictionary annDict)
                            {
                                // Mettre la nouvelle valeur
                                annDict.Put(PdfName.V, new PdfString(val));

                                // Supprimer l'apparence existante pour forcer la régénération
                                annDict.Remove(PdfName.AP);

                                var fieldName = annDict.Get(PdfName.T)?.ToString();
                                Log("PdfAnalysis", $"    [+] {fieldName} = {val}");
                                filledCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Log("PdfAnalysis", $"[!] Erreur annotation index {idx}: {ex.Message}");
                        }
                    }

                    // IMPORTANT: Mettre NeedAppearances=true pour que Adobe régénère les apparences
                    var acroForm = reader.Catalog.GetAsDict(PdfName.ACROFORM);
                    if (acroForm != null)
                    {
                        acroForm.Put(PdfName.NEEDAPPEARANCES, PdfBoolean.PDFTRUE);
                        Log("PdfAnalysis", "[i] NeedAppearances = true");
                    }
                    else
                    {
                        // Créer l'AcroForm s'il n'existe pas
                        acroForm = new PdfDictionary();
                        acroForm.Put(PdfName.NEEDAPPEARANCES, PdfBoolean.PDFTRUE);
                        reader.Catalog.Put(PdfName.ACROFORM, acroForm);
                        Log("PdfAnalysis", "[i] AcroForm cree avec NeedAppearances = true");
                    }
                }

                Log("PdfAnalysis", $"[+] PdfFormFiller: {filledCount}/5 champs remplis avec succes");
                return filledCount > 0;
            }
            catch (Exception ex)
            {
                Log("Error", $"[-] PdfFormFiller: Erreur - {ex.Message}");
                return false;
            }
            finally
            {
                // Nettoyer le fichier temporaire
                if (!string.IsNullOrEmpty(tempPath) && File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch
                    {
                        // Ignorer
                    }
                }
            }
        }

        private static void Log(string category, string message)
        {
            OnLog?.Invoke(category, message);
        }
    }
}
