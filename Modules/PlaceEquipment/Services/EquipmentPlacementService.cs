using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ACW = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Models;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Modules.PlaceEquipment.Services
{
    /// <summary>
    /// Service de placement d'equipements dans les modules XNRGY
    /// Gere le telechargement depuis Vault, Copy Design et insertion dans le top assembly
    /// 
    /// ARCHITECTURE EXTENSIBLE:
    /// Pour ajouter un nouvel equipement, simplement ajouter une entree dans AvailableEquipment:
    /// - Equipement simple: new EquipmentItem { Name, DisplayName, ProjectFileName, AssemblyFileName, VaultPath }
    /// - Equipement IPT: ajouter PrimaryFileType = PrimaryFileType.Part
    /// - Equipement avec variantes: ajouter Variants = new List<EquipmentVariant> { ... }
    /// - Equipement avec dessins multiples: ajouter AlternateDrawings = new List<string> { ... }
    /// </summary>
    public class EquipmentPlacementService
    {
        private readonly Action<string, string> _logCallback;
        private readonly VaultSdkService? _vaultService;

        /// <summary>
        /// Liste complete des equipements disponibles avec leurs fichiers .ipj et .iam/.ipt
        /// 
        /// STRUCTURE STANDARD (email Benoit - Janvier 2026):
        /// Vault: $/Engineering/Library/Equipment/[NomEquipement]
        /// Fichiers: [NomEquipement].ipj, [NomEquipement].iam, [NomEquipement].idw
        /// 
        /// POUR AJOUTER UN NOUVEL EQUIPEMENT:
        /// 1. Equipement standard (.iam):
        ///    new EquipmentItem { 
        ///        Name = "Nom_Dossier", DisplayName = "Nom Affichage", 
        ///        ProjectFileName = "Nom.ipj", AssemblyFileName = "Nom.iam",
        ///        VaultPath = "$/Engineering/Library/Equipment/Nom_Dossier" 
        ///    }
        /// 
        /// 2. Equipement avec piece principale (.ipt):
        ///    new EquipmentItem { 
        ///        Name = "Nom_Dossier", DisplayName = "Nom Affichage",
        ///        ProjectFileName = "Nom.ipj", PrimaryFileName = "Nom.ipt",
        ///        PrimaryFileType = PrimaryFileType.Part,
        ///        VaultPath = "$/Engineering/Library/Equipment/Nom_Dossier"
        ///    }
        /// 
        /// 3. Equipement avec variantes (ex: Infinitum):
        ///    new EquipmentItem {
        ///        Name = "Parent_Folder", DisplayName = "Equipement Parent",
        ///        ProjectFileName = "Parent.ipj",
        ///        Variants = new List<EquipmentVariant> {
        ///            new EquipmentVariant { Name = "Variante1", SubFolder = "Variante1", ... },
        ///            new EquipmentVariant { Name = "Variante2", SubFolder = "Variante2", ... }
        ///        }
        ///    }
        /// </summary>
        public static readonly List<EquipmentItem> AvailableEquipment = new List<EquipmentItem>
        {
            // ══════════════════════════════════════════════════════════════════
            // EQUIPEMENTS STANDARDS (.iam comme fichier principal)
            // Structure: Dossier/Equipement.ipj + Equipement.iam + Equipement.idw
            // ══════════════════════════════════════════════════════════════════
            
            new EquipmentItem { 
                Name = "Angular_Filter", 
                DisplayName = "Angular Filter", 
                ProjectFileName = "Angular_Filter.ipj", 
                AssemblyFileName = "Angular_Filter.iam",
                VaultPath = "$/Engineering/Library/Equipment/Angular_Filter" 
            },
            
            new EquipmentItem { 
                Name = "Bell_Mouth", 
                DisplayName = "Bell Mouth", 
                ProjectFileName = "Bell_Mouth.ipj", 
                AssemblyFileName = "Bell_Mouth.iam",
                VaultPath = "$/Engineering/Library/Equipment/Bell_Mouth" 
            },
            
            new EquipmentItem { 
                Name = "Blank_Test", 
                DisplayName = "Blank Test", 
                ProjectFileName = "Blank_Test.ipj", 
                AssemblyFileName = "Blank_Test.iam",
                VaultPath = "$/Engineering/Library/Equipment/Blank_Test" 
            },
            
            new EquipmentItem { 
                Name = "Catwalk", 
                DisplayName = "Catwalk", 
                ProjectFileName = "Catwalk.ipj", 
                AssemblyFileName = "Catwalk.iam",
                VaultPath = "$/Engineering/Library/Equipment/Catwalk" 
            },
            
            new EquipmentItem { 
                Name = "Circular_Opening_Support", 
                DisplayName = "Circular Opening Support", 
                ProjectFileName = "Circular_Opening_Support.ipj", 
                AssemblyFileName = "Circular_Opening_Support.iam",
                VaultPath = "$/Engineering/Library/Equipment/Circular_Opening_Support" 
            },
            
            new EquipmentItem { 
                Name = "Cooling_Coil", 
                DisplayName = "Cooling Coil", 
                ProjectFileName = "Cooling_Coil.ipj", 
                AssemblyFileName = "Cooling_Coil.iam",
                VaultPath = "$/Engineering/Library/Equipment/Cooling_Coil" 
            },
            
            new EquipmentItem { 
                Name = "Cooling_Coil_Double", 
                DisplayName = "Cooling Coil Double", 
                ProjectFileName = "Cooling_Coil_Double.ipj", 
                AssemblyFileName = "Cooling_Coil_Double.iam",
                VaultPath = "$/Engineering/Library/Equipment/Cooling_Coil_Double" 
            },
            
            // DAMPER: Equipement avec dessins alternatifs (Floor vs Wall/Roof)
            new EquipmentItem { 
                Name = "Damper", 
                DisplayName = "Damper", 
                ProjectFileName = "Damper.ipj", 
                AssemblyFileName = "Damper.iam",
                VaultPath = "$/Engineering/Library/Equipment/Damper",
                AlternateDrawings = new List<string> { "Floor_Damper.idw", "Wall_and_Roof_Damper.idw" }
            },
            
            new EquipmentItem { 
                Name = "Fan_Cube", 
                DisplayName = "Fan Cube", 
                ProjectFileName = "Fan_Cube_Assy.ipj", 
                AssemblyFileName = "Fan_Cube_Assy.iam",
                VaultPath = "$/Engineering/Library/Equipment/Fan_Cube" 
            },
            
            new EquipmentItem { 
                Name = "Floor_Pan_Only", 
                DisplayName = "Floor Pan Only", 
                ProjectFileName = "Floor_Pan.ipj", 
                AssemblyFileName = "Floor_Pan.iam",
                VaultPath = "$/Engineering/Library/Equipment/Floor_Pan_Only" 
            },
            
            new EquipmentItem { 
                Name = "Front_Filter", 
                DisplayName = "Front Filter", 
                ProjectFileName = "Front_Filter.ipj", 
                AssemblyFileName = "Front_Filter.iam",
                VaultPath = "$/Engineering/Library/Equipment/Front_Filter" 
            },
            
            new EquipmentItem { 
                Name = "Heat_Wheel", 
                DisplayName = "Heat Wheel", 
                ProjectFileName = "Heat_Wheel.ipj", 
                AssemblyFileName = "Heat_Wheel.iam",
                VaultPath = "$/Engineering/Library/Equipment/Heat_Wheel" 
            },
            
            new EquipmentItem { 
                Name = "Heating_Coil", 
                DisplayName = "Heating Coil", 
                ProjectFileName = "Heating_Coil.ipj", 
                AssemblyFileName = "Heating_Coil.iam",
                VaultPath = "$/Engineering/Library/Equipment/Heating_Coil" 
            },
            
            new EquipmentItem { 
                Name = "Heating_Coil_Double", 
                DisplayName = "Heating Coil Double", 
                ProjectFileName = "Heating_Coil_Double.ipj", 
                AssemblyFileName = "Heating_Coil_Double.iam",
                VaultPath = "$/Engineering/Library/Equipment/Heating_Coil_Double" 
            },
            
            new EquipmentItem { 
                Name = "Hood", 
                DisplayName = "Hood", 
                ProjectFileName = "Hood.ipj", 
                AssemblyFileName = "Hood.iam",
                VaultPath = "$/Engineering/Library/Equipment/Hood" 
            },
            
            new EquipmentItem { 
                Name = "Humidifier", 
                DisplayName = "Humidifier", 
                ProjectFileName = "Humidifier.ipj", 
                AssemblyFileName = "Humidifier.iam",
                VaultPath = "$/Engineering/Library/Equipment/Humidifier" 
            },
            
            new EquipmentItem { 
                Name = "Isolator_Plate", 
                DisplayName = "Isolator Plate", 
                ProjectFileName = "Isolator_Plate.ipj", 
                AssemblyFileName = "Isolator_Plate.iam",
                VaultPath = "$/Engineering/Library/Equipment/Isolator_Plate" 
            },
            
            new EquipmentItem { 
                Name = "Outlet_Guard", 
                DisplayName = "Outlet Guard", 
                ProjectFileName = "Outlet_Guard.ipj", 
                AssemblyFileName = "Outlet_Guard.iam",
                VaultPath = "$/Engineering/Library/Equipment/Outlet_Guard" 
            },
            
            new EquipmentItem { 
                Name = "Rehausse", 
                DisplayName = "Rehausse", 
                ProjectFileName = "Rehausse.ipj", 
                AssemblyFileName = "Rehausse.iam",
                VaultPath = "$/Engineering/Library/Equipment/Rehausse" 
            },
            
            new EquipmentItem { 
                Name = "Removable_Panel", 
                DisplayName = "Removable Panel", 
                ProjectFileName = "Removable_Panel.ipj", 
                AssemblyFileName = "Removable_Panel.iam",
                VaultPath = "$/Engineering/Library/Equipment/Removable_Panel" 
            },
            
            new EquipmentItem { 
                Name = "Silencer", 
                DisplayName = "Silencer", 
                ProjectFileName = "Silencer.ipj", 
                AssemblyFileName = "Silencer.iam",
                VaultPath = "$/Engineering/Library/Equipment/Silencer" 
            },
            
            new EquipmentItem { 
                Name = "Transition", 
                DisplayName = "Transition", 
                ProjectFileName = "Transition.ipj", 
                AssemblyFileName = "Transition.iam",
                VaultPath = "$/Engineering/Library/Equipment/Transition" 
            },
            
            new EquipmentItem { 
                Name = "VicWest", 
                DisplayName = "VicWest", 
                ProjectFileName = "VicWest.ipj", 
                AssemblyFileName = "VicWest.iam",
                VaultPath = "$/Engineering/Library/Equipment/VicWest" 
            },
            
            new EquipmentItem { 
                Name = "Xnrgy_Door", 
                DisplayName = "XNRGY Door", 
                ProjectFileName = "Xnrgy_Door.ipj", 
                AssemblyFileName = "Xnrgy_Door.iam",
                VaultPath = "$/Engineering/Library/Equipment/Xnrgy_Door" 
            },
            
            // ══════════════════════════════════════════════════════════════════
            // EQUIPEMENTS AVEC SOUS-DOSSIERS (Dwyer_Gage)
            // Structure: Parent/SubFolder/Equipement.ipj + .iam
            // ══════════════════════════════════════════════════════════════════
            
            new EquipmentItem { 
                Name = "Dwyer_2000", 
                DisplayName = "Dwyer 2000", 
                ProjectFileName = "Dwyer_2000.ipj", 
                AssemblyFileName = "Dwyer_2000.iam",
                VaultPath = "$/Engineering/Library/Equipment/Dwyer_Gage/Dwyer_2000" 
            },
            
            new EquipmentItem { 
                Name = "Dwyer_3000", 
                DisplayName = "Dwyer 3000", 
                ProjectFileName = "Dwyer_3000.ipj", 
                AssemblyFileName = "Dwyer_3000.iam",
                VaultPath = "$/Engineering/Library/Equipment/Dwyer_Gage/Dwyer_3000" 
            },
            
            // ══════════════════════════════════════════════════════════════════
            // EQUIPEMENTS INFINITUM (Moteurs avec variantes)
            // Chaque variante a son propre IPJ et fichier principal (.iam ou .ipt)
            // L'utilisateur doit choisir quelle variante utiliser
            // ══════════════════════════════════════════════════════════════════
            
            new EquipmentItem { 
                Name = "Infinitum", 
                DisplayName = "Infinitum Motor", 
                ProjectFileName = "Infinitum.ipj",  // IPJ parent (peut ne pas exister)
                VaultPath = "$/Engineering/Library/Equipment/Infinitum",
                Variants = new List<EquipmentVariant>
                {
                    new EquipmentVariant { 
                        Name = "7.5Hp_1800RPM", 
                        DisplayName = "7.5 HP @ 1800 RPM (IES180)", 
                        ProjectFileName = "7.5Hp_1800RPM_IES180.ipj",
                        PrimaryFileName = "7.5Hp_1800RPM_IES180.ipt",
                        FileType = PrimaryFileType.Part,
                        SubFolder = "7.5Hp_1800RPM"
                    },
                    new EquipmentVariant { 
                        Name = "7.5Hp_3600RPM", 
                        DisplayName = "7.5 HP @ 3600 RPM", 
                        ProjectFileName = "7.5Hp_3600RPM.ipj",
                        PrimaryFileName = "7.5Hp_3600RPM.iam",
                        FileType = PrimaryFileType.Assembly,
                        SubFolder = "7.5Hp_3600RPM"
                    },
                    new EquipmentVariant { 
                        Name = "10Hp_1800RPM_Gen2_1piecesVFD", 
                        DisplayName = "10 HP @ 1800 RPM Gen2 (1pc VFD)", 
                        ProjectFileName = "10Hp_1800RPM_Gen2_1piecesVFD.ipj",
                        PrimaryFileName = "10Hp_1800RPM_Gen2_1piecesVFD.ipt",
                        FileType = PrimaryFileType.Part,
                        SubFolder = "10Hp_1800RPM_Gen2_1piecesVFD"
                    },
                    new EquipmentVariant { 
                        Name = "10Hp_1800RPM_Gen2_2piecesVFD", 
                        DisplayName = "10 HP @ 1800 RPM Gen2 (2pc VFD)", 
                        ProjectFileName = "10Hp_1800RPM_Gen2_2piecesVFD.ipj",
                        PrimaryFileName = "10Hp_1800RPM_Gen2_2piecesVFD.ipt",
                        FileType = PrimaryFileType.Part,
                        SubFolder = "10Hp_1800RPM_Gen2_2piecesVFD"
                    },
                    new EquipmentVariant { 
                        Name = "10Hp_2400RPM", 
                        DisplayName = "10 HP @ 2400 RPM", 
                        ProjectFileName = "10Hp_2400RPM.ipj",
                        PrimaryFileName = "10Hp_2400RPM.ipt",
                        FileType = PrimaryFileType.Part,
                        SubFolder = "10Hp_2400RPM"
                    },
                    new EquipmentVariant { 
                        Name = "IES150", 
                        DisplayName = "IES150", 
                        ProjectFileName = "IES150.ipj",
                        PrimaryFileName = "IES150.iam",
                        FileType = PrimaryFileType.Assembly,
                        SubFolder = "IES150"
                    }
                }
            }
        };

        public EquipmentPlacementService(VaultSdkService? vaultService, Action<string, string>? logCallback = null)
        {
            _vaultService = vaultService;
            _logCallback = logCallback ?? ((msg, level) => { });
        }

        /// <summary>
        /// Obtient la liste de tous les équipements disponibles
        /// </summary>
        public ObservableCollection<EquipmentItem> GetAvailableEquipment()
        {
            var equipment = new ObservableCollection<EquipmentItem>();
            foreach (var item in AvailableEquipment.OrderBy(e => e.DisplayName))
            {
                // Définir le chemin local temporaire
                item.LocalTempPath = Path.Combine(@"C:\Vault\Engineering\Library\Equipment", item.Name);
                equipment.Add(item);
            }
            return equipment;
        }

        /// <summary>
        /// Télécharge un équipement depuis Vault vers le répertoire temporaire local
        /// Réutilise la logique de CreateModule pour le téléchargement
        /// </summary>
        public async Task<bool> DownloadEquipmentFromVaultAsync(EquipmentItem equipment, Action<int, string>? progressCallback = null)
        {
            if (_vaultService == null || !_vaultService.IsConnected)
            {
                Log("Service Vault non disponible ou non connecté", "ERROR");
                return false;
            }

            try
            {
                Log($"Téléchargement équipement depuis Vault: {equipment.VaultPath}", "INFO");
                progressCallback?.Invoke(0, "Connexion au dossier Vault...");
                
                var connection = _vaultService.Connection;
                if (connection == null)
                {
                    Log("Connexion Vault perdue", "ERROR");
                    return false;
                }

                // Créer le répertoire temporaire s'il n'existe pas
                var tempPath = Path.Combine(@"C:\Vault\Engineering\Library\Equipment", equipment.Name);
                if (Directory.Exists(tempPath))
                {
                    // Nettoyer le répertoire existant
                    try
                    {
                        Directory.Delete(tempPath, true);
                    }
                    catch { }
                }
                Directory.CreateDirectory(tempPath);
                Log($"Répertoire créé: {tempPath}", "DEBUG");

                // Obtenir le dossier Vault
                progressCallback?.Invoke(5, "Connexion au dossier Vault...");
                var vaultFolder = connection.WebServiceManager.DocumentService.GetFolderByPath(equipment.VaultPath);
                if (vaultFolder == null)
                {
                    Log($"Dossier Vault non trouvé: {equipment.VaultPath}", "ERROR");
                    return false;
                }
                Log($"Dossier Vault trouvé: {equipment.VaultPath}", "INFO");

                // Obtenir TOUS les fichiers RÉCURSIVEMENT
                progressCallback?.Invoke(10, "Énumération récursive des fichiers...");
                var allFiles = new List<ACW.File>();
                var allFolders = new List<ACW.Folder>();
                await Task.Run(() => GetAllFilesRecursive(connection, vaultFolder, allFiles, allFolders));

                if (allFiles.Count == 0)
                {
                    Log("Aucun fichier trouvé dans l'équipement Vault", "ERROR");
                    return false;
                }
                Log($"{allFiles.Count} fichiers trouvés au total (récursif)", "SUCCESS");

                // Obtenir le working folder
                var workingFolderObj = connection.WorkingFoldersManager.GetWorkingFolder("$");
                if (workingFolderObj == null || string.IsNullOrEmpty(workingFolderObj.FullPath))
                {
                    Log("Working folder non configuré dans Vault", "ERROR");
                    return false;
                }

                var workingFolder = workingFolderObj.FullPath;
                var relativePath = equipment.VaultPath.TrimStart('$', '/').Replace("/", "\\");
                var localFolder = Path.Combine(workingFolder, relativePath);

                // Préparer le téléchargement batch
                progressCallback?.Invoke(15, $"Préparation de {allFiles.Count} fichiers...");
                var downloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection, false);

                foreach (var file in allFiles)
                {
                    try
                    {
                        var fileIteration = new VDF.Vault.Currency.Entities.FileIteration(connection, file);
                        downloadSettings.AddFileToAcquire(fileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                    }
                    catch (Exception fileEx)
                    {
                        Log($"Erreur préparation {file.Name}: {fileEx.Message}", "WARNING");
                    }
                }

                // Télécharger tous les fichiers en UNE SEULE opération batch
                progressCallback?.Invoke(20, $"Téléchargement batch de {allFiles.Count} fichiers...");
                var downloadResult = await Task.Run(() => connection.FileManager.AcquireFiles(downloadSettings));

                if (downloadResult?.FileResults == null || !downloadResult.FileResults.Any())
                {
                    Log("Aucun fichier téléchargé", "ERROR");
                    return false;
                }

                var fileResultsList = downloadResult.FileResults.ToList();
                int successCount = fileResultsList.Count(r => r.LocalPath?.FullPath != null && File.Exists(r.LocalPath.FullPath));
                Log($"{successCount}/{fileResultsList.Count} fichiers téléchargés", "SUCCESS");

                // Copier vers le dossier temporaire
                progressCallback?.Invoke(70, "Copie vers dossier temporaire...");
                
                if (Directory.Exists(localFolder))
                {
                    CopyDirectory(localFolder, tempPath);
                }
                else
                {
                    // Copier chaque fichier téléchargé en préservant la structure
                    foreach (var result in fileResultsList)
                    {
                        if (result?.LocalPath?.FullPath == null) continue;

                        var localFilePath = result.LocalPath.FullPath;
                        if (!File.Exists(localFilePath)) continue;

                        string relativeFilePath;
                        if (localFilePath.StartsWith(workingFolder, StringComparison.OrdinalIgnoreCase))
                        {
                            relativeFilePath = localFilePath.Substring(workingFolder.Length).TrimStart('\\', '/');
                            var projectRelativePath = relativePath.TrimStart('\\', '/');
                            if (!string.IsNullOrEmpty(projectRelativePath) && relativeFilePath.StartsWith(projectRelativePath, StringComparison.OrdinalIgnoreCase))
                            {
                                relativeFilePath = relativeFilePath.Substring(projectRelativePath.Length).TrimStart('\\', '/');
                            }
                        }
                        else
                        {
                            relativeFilePath = Path.GetFileName(localFilePath);
                        }

                        var destFilePath = Path.Combine(tempPath, relativeFilePath);
                        var destDir = Path.GetDirectoryName(destFilePath);
                        if (!string.IsNullOrEmpty(destDir))
                        {
                            Directory.CreateDirectory(destDir);
                        }

                        try
                        {
                            File.Copy(localFilePath, destFilePath, true);
                        }
                        catch (Exception copyEx)
                        {
                            Log($"Erreur copie {Path.GetFileName(localFilePath)}: {copyEx.Message}", "WARNING");
                        }
                    }
                }

                var copiedFiles = Directory.GetFiles(tempPath, "*.*", SearchOption.AllDirectories);
                if (copiedFiles.Length == 0)
                {
                    Log("Aucun fichier copié vers le dossier temporaire", "ERROR");
                    return false;
                }

                Log($"{copiedFiles.Length} fichiers copiés vers le dossier temporaire", "SUCCESS");
                equipment.LocalTempPath = tempPath;
                progressCallback?.Invoke(100, "Téléchargement terminé");
                return true;
            }
            catch (Exception ex)
            {
                Log($"Erreur téléchargement équipement: {ex.Message}", "ERROR");
                return false;
            }
        }

        /// <summary>
        /// Obtient tous les fichiers récursivement depuis un dossier Vault
        /// </summary>
        private void GetAllFilesRecursive(VDF.Vault.Currency.Connections.Connection connection, ACW.Folder folder, List<ACW.File> allFiles, List<ACW.Folder> allFolders)
        {
            try
            {
                allFolders.Add(folder);
                var files = connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, false);
                if (files != null && files.Length > 0)
                {
                    allFiles.AddRange(files);
                }

                var subFolders = connection.WebServiceManager.DocumentService.GetFoldersByParentId(folder.Id, false);
                if (subFolders != null && subFolders.Length > 0)
                {
                    foreach (var subFolder in subFolders)
                    {
                        GetAllFilesRecursive(connection, subFolder, allFiles, allFolders);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Erreur énumération {folder.FullName}: {ex.Message}", "WARNING");
            }
        }

        /// <summary>
        /// Copie récursivement un répertoire
        /// </summary>
        private void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);

            foreach (var file in Directory.GetFiles(sourceDir))
            {
                var fileName = Path.GetFileName(file);
                var destFile = Path.Combine(destDir, fileName);
                File.Copy(file, destFile, true);
            }

            foreach (var dir in Directory.GetDirectories(sourceDir))
            {
                var dirName = Path.GetFileName(dir);
                var destSubDir = Path.Combine(destDir, dirName);
                CopyDirectory(dir, destSubDir);
            }
        }

        /// <summary>
        /// Détecte le module actif dans Inventor et extrait Project/Reference/Module
        /// Parse le chemin: C:\Vault\Engineering\Projects\{Project}\REF{Reference}\M{Module}\...
        /// </summary>
        public (string Project, string Reference, string Module, string TopAssemblyPath, string ProjectPath, string ProjectFile) DetectActiveModule(InventorService inventorService)
        {
            try
            {
                Log("Détection du module actif dans Inventor...", "INFO");
                
                if (!inventorService.IsConnected)
                {
                    Log("Inventor non connecté", "ERROR");
                    return (string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }

                // Obtenir le document actif
                var activeDocPath = inventorService.GetActiveDocumentPath();
                if (string.IsNullOrWhiteSpace(activeDocPath))
                {
                    Log("Aucun document actif dans Inventor", "ERROR");
                    return (string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }

                Log($"Document actif: {Path.GetFileName(activeDocPath)}", "INFO");

                // Parser le chemin pour extraire Project/Reference/Module
                // Format: C:\Vault\Engineering\Projects\{Project}\REF{Reference}\M{Module}\...
                var pattern = @"C:\\Vault\\Engineering\\Projects\\([^\\]+)\\REF(\d+)\\M(\d+)";
                var match = Regex.Match(activeDocPath, pattern, RegexOptions.IgnoreCase);
                
                if (!match.Success)
                {
                    Log($"Format de chemin non reconnu: {activeDocPath}", "ERROR");
                    return (string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }

                var project = match.Groups[1].Value;
                var reference = match.Groups[2].Value;
                var module = match.Groups[3].Value;

                Log($"Projet détecté: {project}, Ref: {reference}, Module: {module}", "SUCCESS");

                // Construire le chemin du projet
                var projectPath = Path.Combine(@"C:\Vault\Engineering\Projects", project, $"REF{reference}", $"M{module}");
                
                // Trouver le fichier .ipj du projet
                string? projectFile = null;
                if (Directory.Exists(projectPath))
                {
                    var ipjFiles = Directory.GetFiles(projectPath, "*.ipj", SearchOption.TopDirectoryOnly);
                    if (ipjFiles.Length > 0)
                    {
                        // Chercher le fichier principal (pattern XXXXX-XX-XX_2026.ipj)
                        var mainIpj = ipjFiles.FirstOrDefault(f =>
                        {
                            var fileName = Path.GetFileName(f);
                            return Regex.IsMatch(fileName, @"^\d{5}-\d{2}-\d{2}_2026\.ipj$", RegexOptions.IgnoreCase);
                        });
                        
                        projectFile = mainIpj ?? ipjFiles[0];
                        Log($"Fichier projet trouvé: {Path.GetFileName(projectFile)}", "INFO");
                    }
                }

                return (project, reference, module, activeDocPath, projectPath, projectFile ?? string.Empty);
            }
            catch (Exception ex)
            {
                Log($"Erreur détection module: {ex.Message}", "ERROR");
                return (string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
            }
        }

        private void Log(string message, string level)
        {
            _logCallback?.Invoke(message, level);
        }
    }
}

