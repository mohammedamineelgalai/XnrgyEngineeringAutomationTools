using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using XnrgyEngineeringAutomationTools.Modules.ACP.Models;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Modules.ACP.Services
{
    /// <summary>
    /// Service pour lire et parser les fichiers Excel ACP
    /// Utilise EPPlus pour lire les fichiers .xlsm
    /// </summary>
    public class ACPExcelService
    {
        /// <summary>
        /// Lit un fichier Excel ACP et extrait la structure Unité → Modules → Points Critiques
        /// </summary>
        public ACPDataModel? ReadACPFromExcel(string excelFilePath)
        {
            try
            {
                if (!File.Exists(excelFilePath))
                {
                    Logger.Log($"[ACPExcel] Fichier non trouvé: {excelFilePath}", Logger.LogLevel.ERROR);
                    return null;
                }

                Logger.Log($"[ACPExcel] Lecture du fichier ACP: {excelFilePath}", Logger.LogLevel.INFO);

                // Note: EPPlus 4.x n'a pas de LicenseContext (seulement EPPlus 5+)
                // Le projet utilise EPPlus 4.5.3.3 compatible .NET Framework 4.8

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var workbook = package.Workbook;
                    
                    // Extraire le numéro de projet et référence depuis le nom du fichier
                    // Format: "10516-01_ACP_Rev-05.xlsm"
                    string fileName = Path.GetFileNameWithoutExtension(excelFilePath);
                    var parts = fileName.Split('_');
                    string projectRef = parts.Length > 0 ? parts[0] : "";  // "10516-01"
                    
                    var projectParts = projectRef.Split('-');
                    string projectNumber = projectParts.Length > 0 ? projectParts[0] : "";
                    string reference = projectParts.Length > 1 ? projectParts[1] : "";

                    var acpData = new ACPDataModel
                    {
                        UnitId = projectRef,
                        ProjectNumber = projectNumber,
                        Reference = reference,
                        UnitName = $"Unité {projectRef}",
                        CreatedDate = File.GetCreationTime(excelFilePath),
                        LastModifiedDate = File.GetLastWriteTime(excelFilePath),
                        Modules = new Dictionary<string, ACPModule>()
                    };

                    // Scanner toutes les feuilles du workbook
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        string sheetName = worksheet.Name;

                        // Ignorer les feuilles système
                        if (sheetName.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase) ||
                            sheetName.Equals("DATA", StringComparison.OrdinalIgnoreCase) ||
                            sheetName.Equals("Unit Infos", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        // Détecter les feuilles de module (format: "M01", "M02", etc.)
                        if (sheetName.StartsWith("M", StringComparison.OrdinalIgnoreCase) &&
                            sheetName.Length >= 2 &&
                            char.IsDigit(sheetName[1]))
                        {
                            string moduleId = sheetName.Substring(0, Math.Min(3, sheetName.Length));  // "M01", "M02", etc.
                            
                            var module = ParseModuleSheet(worksheet, moduleId);
                            if (module != null && module.CriticalPoints.Count > 0)
                            {
                                acpData.Modules[moduleId] = module;
                                Logger.Log($"[ACPExcel] Module {moduleId} trouvé avec {module.CriticalPoints.Count} points critiques", Logger.LogLevel.INFO);
                            }
                        }
                        // Feuille "Communication" contient des points critiques communs à tous les modules
                        else if (sheetName.Equals("Communication", StringComparison.OrdinalIgnoreCase))
                        {
                            var communicationPoints = ParseCommunicationSheet(worksheet);
                            
                            // Ajouter ces points à tous les modules existants
                            foreach (var module in acpData.Modules.Values)
                            {
                                foreach (var point in communicationPoints)
                                {
                                    // Vérifier si le point existe déjà
                                    if (!module.CriticalPoints.Any(p => p.Id == point.Id))
                                    {
                                        module.CriticalPoints.Add(point);
                                    }
                                }
                            }
                        }
                        // Feuille "Concepteur Lead" contient des points critiques pour le lead
                        else if (sheetName.Equals("Concepteur Lead", StringComparison.OrdinalIgnoreCase))
                        {
                            var leadPoints = ParseConcepteurLeadSheet(worksheet);
                            
                            // Ajouter à tous les modules
                            foreach (var module in acpData.Modules.Values)
                            {
                                foreach (var point in leadPoints)
                                {
                                    if (!module.CriticalPoints.Any(p => p.Id == point.Id))
                                    {
                                        module.CriticalPoints.Add(point);
                                    }
                                }
                            }
                        }
                    }

                    Logger.Log($"[ACPExcel] ACP lu avec succès: {acpData.Modules.Count} modules trouvés", Logger.LogLevel.INFO);
                    return acpData;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPExcel.ReadACPFromExcel", ex, Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// Parse une feuille de module (M01, M02, etc.)
        /// </summary>
        private ACPModule? ParseModuleSheet(ExcelWorksheet worksheet, string moduleId)
        {
            try
            {
                var module = new ACPModule
                {
                    ModuleId = moduleId,
                    ModuleName = $"Module {moduleId}",
                    CriticalPoints = new List<CriticalPoint>()
                };

                int startRow = 1;
                int currentPointId = 1;
                int endRow = worksheet.Dimension?.End.Row ?? 0;

                // Scanner les lignes pour trouver les points critiques
                for (int row = startRow; row <= endRow; row++)
                {
                    // Lire les colonnes B à Q (comme dans le fichier Excel montré)
                    string? category = worksheet.Cells[row, 2]?.Text?.Trim();
                    string? title = worksheet.Cells[row, 3]?.Text?.Trim();
                    string? description = worksheet.Cells[row, 4]?.Text?.Trim();
                    string? notes = "";

                    // Combiner toutes les colonnes suivantes comme notes/description étendue
                    var noteParts = new List<string>();
                    for (int col = 5; col <= 17; col++)  // Colonnes E à Q
                    {
                        string? cellText = worksheet.Cells[row, col]?.Text?.Trim();
                        if (!string.IsNullOrWhiteSpace(cellText))
                        {
                            noteParts.Add(cellText);
                        }
                    }
                    
                    if (noteParts.Count > 0)
                    {
                        notes = string.Join(" ", noteParts);
                    }

                    // Si on a au moins un titre ou une description, créer un point critique
                    if (!string.IsNullOrWhiteSpace(title) || !string.IsNullOrWhiteSpace(description))
                    {
                        var point = new CriticalPoint
                        {
                            Id = currentPointId++,
                            Category = category ?? "Général",
                            Title = title ?? "",
                            Description = description ?? "",
                            Notes = notes,
                            Priority = DeterminePriority(title, description, notes)
                        };

                        module.CriticalPoints.Add(point);
                    }
                }

                return module.CriticalPoints.Count > 0 ? module : null;
            }
            catch (Exception ex)
            {
                Logger.LogException($"ACPExcel.ParseModuleSheet({moduleId})", ex, Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// Parse la feuille "Communication"
        /// </summary>
        private List<CriticalPoint> ParseCommunicationSheet(ExcelWorksheet worksheet)
        {
            var points = new List<CriticalPoint>();
            int currentPointId = 1000;  // IDs élevés pour différencier des points module
            int endRow = worksheet.Dimension?.End.Row ?? 0;

            try
            {
                for (int row = 15; row <= endRow; row++)
                {
                    string? text = worksheet.Cells[row, 2]?.Text?.Trim();  // Colonne B
                    
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    // Détecter les catégories (lignes avec formatage spécial ou en gras)
                    bool isCategory = false;
                    try
                    {
                        var cell = worksheet.Cells[row, 2];
                        isCategory = cell.Style.Font.Bold || 
                                    text.All(c => char.IsUpper(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c));
                    }
                    catch { }

                    if (!isCategory && text.Length > 10)  // Ignorer les catégories courtes
                    {
                        var point = new CriticalPoint
                        {
                            Id = currentPointId++,
                            Category = "Communication",
                            Title = text.Length > 100 ? text.Substring(0, 100) : text,
                            Description = text,
                            Notes = "",
                            Priority = "Normal"
                        };

                        points.Add(point);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPExcel.ParseCommunicationSheet", ex, Logger.LogLevel.ERROR);
            }

            return points;
        }

        /// <summary>
        /// Parse la feuille "Concepteur Lead"
        /// </summary>
        private List<CriticalPoint> ParseConcepteurLeadSheet(ExcelWorksheet worksheet)
        {
            var points = new List<CriticalPoint>();
            int currentPointId = 2000;  // IDs élevés pour différencier
            int endRow = worksheet.Dimension?.End.Row ?? 0;

            try
            {
                // Logique similaire à ParseCommunicationSheet
                for (int row = 1; row <= endRow; row++)
                {
                    string? text = worksheet.Cells[row, 2]?.Text?.Trim();
                    
                    if (string.IsNullOrWhiteSpace(text) || text.Length < 10) continue;

                    var point = new CriticalPoint
                    {
                        Id = currentPointId++,
                        Category = "Concepteur Lead",
                        Title = text.Length > 100 ? text.Substring(0, 100) : text,
                        Description = text,
                        Priority = "Haute"  // Les points du lead sont prioritaires
                    };

                    points.Add(point);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPExcel.ParseConcepteurLeadSheet", ex, Logger.LogLevel.ERROR);
            }

            return points;
        }

        /// <summary>
        /// Détermine la priorité d'un point critique basé sur son contenu
        /// </summary>
        private string DeterminePriority(string? title, string? description, string? notes)
        {
            string combined = $"{title} {description} {notes}".ToLowerInvariant();

            // Mots-clés pour priorité haute
            if (combined.Contains("requis") || combined.Contains("obligatoire") || 
                combined.Contains("attention") || combined.Contains("important") ||
                combined.Contains("critique") || combined.Contains("vérifier"))
            {
                return "Haute";
            }

            // Mots-clés pour priorité basse
            if (combined.Contains("optionnel") || combined.Contains("si nécessaire") ||
                combined.Contains("peut-être"))
            {
                return "Basse";
            }

            return "Normal";
        }

        /// <summary>
        /// Recherche tous les fichiers ACP dans un dossier de projet
        /// </summary>
        public List<string> FindACPFiles(string projectFolderPath)
        {
            var acpFiles = new List<string>();

            try
            {
                if (!Directory.Exists(projectFolderPath))
                    return acpFiles;

                // Rechercher fichiers ACP (format: *_ACP_*.xlsm ou *_ACP_*.xlsx)
                var files = Directory.GetFiles(projectFolderPath, "*_ACP_*.xlsm", SearchOption.TopDirectoryOnly)
                    .Concat(Directory.GetFiles(projectFolderPath, "*_ACP_*.xlsx", SearchOption.TopDirectoryOnly));

                acpFiles.AddRange(files);

                Logger.Log($"[ACPExcel] {acpFiles.Count} fichier(s) ACP trouvé(s) dans {projectFolderPath}", Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPExcel.FindACPFiles", ex, Logger.LogLevel.ERROR);
            }

            return acpFiles;
        }

        /// <summary>
        /// Recherche tous les fichiers ACP dans C:\Vault\Engineering\Projects et C:\Engineering\Projects
        /// </summary>
        public List<ACPUnitInfo> FindAllACPFiles()
        {
            var units = new List<ACPUnitInfo>();

            try
            {
                // Scanner les deux emplacements possibles
                var basePaths = new List<string>
                {
                    @"C:\Engineering\Projects",       // Emplacement principal
                    @"C:\Vault\Engineering\Projects"  // Emplacement Vault
                };

                foreach (var basePath in basePaths)
                {
                    if (!Directory.Exists(basePath))
                    {
                        Logger.Log($"[ACPExcel] Dossier non trouve: {basePath}", Logger.LogLevel.DEBUG);
                        continue;
                    }

                    Logger.Log($"[ACPExcel] Scan du dossier: {basePath}", Logger.LogLevel.INFO);

                    // Scanner tous les projets (dossiers numériques)
                    var projectDirs = Directory.GetDirectories(basePath)
                        .Where(d => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(d), @"^\d+$"))
                        .OrderBy(d => d);

                    foreach (var projectDir in projectDirs)
                    {
                        string projectNumber = Path.GetFileName(projectDir);

                        // Scanner les références (REF01, REF02, etc.)
                        var refDirs = Directory.GetDirectories(projectDir)
                            .Where(d => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(d), @"^REF\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                            .OrderBy(d => d);

                        foreach (var refDir in refDirs)
                        {
                            string reference = Path.GetFileName(refDir).Substring(3);  // "REF01" -> "01"

                            // Chercher fichiers ACP dans le dossier REF
                            var acpFiles = FindACPFiles(refDir);

                            foreach (var acpFile in acpFiles)
                            {
                                // Extraire le numéro de projet et référence depuis le nom du fichier
                                string fileName = Path.GetFileNameWithoutExtension(acpFile);
                                var parts = fileName.Split('_');
                                string unitId = parts.Length > 0 ? parts[0] : $"{projectNumber}-{reference}";

                                // Éviter les doublons
                                if (!units.Any(u => u.FilePath.Equals(acpFile, StringComparison.OrdinalIgnoreCase)))
                                {
                                    units.Add(new ACPUnitInfo
                                    {
                                        UnitId = unitId,
                                        ProjectNumber = projectNumber,
                                        Reference = reference,
                                        FilePath = acpFile,
                                        LastModified = File.GetLastWriteTime(acpFile)
                                    });
                                }
                            }
                        }
                    }
                }

                Logger.Log($"[ACPExcel] {units.Count} unite(s) ACP trouvee(s) au total", Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPExcel.FindAllACPFiles", ex, Logger.LogLevel.ERROR);
            }

            return units;
        }
    }

    /// <summary>
    /// Informations sur une unité ACP trouvée
    /// </summary>
    public class ACPUnitInfo
    {
        public string UnitId { get; set; } = "";
        public string ProjectNumber { get; set; } = "";
        public string Reference { get; set; } = "";
        public string FilePath { get; set; } = "";
        public DateTime LastModified { get; set; }
        public string DisplayName => $"{UnitId} ({Path.GetFileName(FilePath)})";
    }
}


