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
    /// Service de placement d'équipements dans les modules XNRGY
    /// Gère le téléchargement depuis Vault, Copy Design et insertion dans le top assembly
    /// </summary>
    public class EquipmentPlacementService
    {
        private readonly Action<string, string> _logCallback;
        private readonly VaultSdkService? _vaultService;

        /// <summary>
        /// Liste complète des équipements disponibles avec leurs fichiers .ipj et .iam
        /// </summary>
        public static readonly List<EquipmentItem> AvailableEquipment = new List<EquipmentItem>
        {
            new EquipmentItem { Name = "AngularFilter", DisplayName = "Angular Filter", ProjectFileName = "Angular Filter.ipj", AssemblyFileName = "Angular Filter.iam", VaultPath = "$/Engineering/Library/Equipment/AngularFilter" },
            new EquipmentItem { Name = "Bellmouth", DisplayName = "Bellmouth", ProjectFileName = "BellMouth.ipj", AssemblyFileName = "BellMouth.iam", VaultPath = "$/Engineering/Library/Equipment/Bellmouth" },
            new EquipmentItem { Name = "BlankTest", DisplayName = "Blank Test", ProjectFileName = "Blank Test.ipj", AssemblyFileName = "Blank Test.iam", VaultPath = "$/Engineering/Library/Equipment/BlankTest" },
            new EquipmentItem { Name = "CatWalk", DisplayName = "Cat Walk", ProjectFileName = "Cat_Walk.ipj", AssemblyFileName = "Handrail.iam", VaultPath = "$/Engineering/Library/Equipment/CatWalk" },
            new EquipmentItem { Name = "CircularOpeningSupport", DisplayName = "Circular Opening Support", ProjectFileName = "Circular_Opening_Support.ipj", AssemblyFileName = "Circular_Opening_Support.iam", VaultPath = "$/Engineering/Library/Equipment/CircularOpeningSupport" },
            new EquipmentItem { Name = "CoolingCoil", DisplayName = "Cooling Coil", ProjectFileName = "CoolingCoil.ipj", AssemblyFileName = "CoolingCoil.iam", VaultPath = "$/Engineering/Library/Equipment/CoolingCoil" },
            new EquipmentItem { Name = "CoolingCoilDouble", DisplayName = "Cooling Coil Double", ProjectFileName = "Cooling Coil Double.ipj", AssemblyFileName = "Cooling Coil Double.iam", VaultPath = "$/Engineering/Library/Equipment/CoolingCoilDouble" },
            new EquipmentItem { Name = "Damper", DisplayName = "Damper", ProjectFileName = "Damper.ipj", AssemblyFileName = "Damper.iam", VaultPath = "$/Engineering/Library/Equipment/Damper" },
            new EquipmentItem { Name = "Dwyer_2000", DisplayName = "Dwyer 2000", ProjectFileName = "Dwyer_2000.ipj", AssemblyFileName = "Dwyer_2000.iam", VaultPath = "$/Engineering/Library/Equipment/Dwyer_Gage/Dwyer_2000" },
            new EquipmentItem { Name = "Dwyer_3000", DisplayName = "Dwyer 3000", ProjectFileName = "Dwyer_3000.ipj", AssemblyFileName = "Dwyer_3000.iam", VaultPath = "$/Engineering/Library/Equipment/Dwyer_Gage/Dwyer_3000" },
            new EquipmentItem { Name = "FanCube", DisplayName = "Fan Cube", ProjectFileName = "Fan Cube Assy.ipj", AssemblyFileName = "Fan Cube Assy.iam", VaultPath = "$/Engineering/Library/Equipment/FanCube" },
            new EquipmentItem { Name = "FloorPanOnly", DisplayName = "Floor Pan Only", ProjectFileName = "Floor Pan.ipj", AssemblyFileName = "Floor Pan.iam", VaultPath = "$/Engineering/Library/Equipment/FloorPanOnly" },
            new EquipmentItem { Name = "FrontFilter", DisplayName = "Front Filter", ProjectFileName = "Front Filter.ipj", AssemblyFileName = "Front Filter.iam", VaultPath = "$/Engineering/Library/Equipment/FrontFilter" },
            new EquipmentItem { Name = "HeatingCoil", DisplayName = "Heating Coil", ProjectFileName = "HeatingCoil.ipj", AssemblyFileName = "HeatingCoil.iam", VaultPath = "$/Engineering/Library/Equipment/HeatingCoil" },
            new EquipmentItem { Name = "HeatingCoilDouble", DisplayName = "Heating Coil Double", ProjectFileName = "Heating Coil Double.ipj", AssemblyFileName = "Heating Coil Double.iam", VaultPath = "$/Engineering/Library/Equipment/HeatingCoilDouble" },
            new EquipmentItem { Name = "HeatWheel", DisplayName = "Heat Wheel", ProjectFileName = "HeatWheel.ipj", AssemblyFileName = "HeatWheel.iam", VaultPath = "$/Engineering/Library/Equipment/HeatWheel" },
            new EquipmentItem { Name = "Humidifier", DisplayName = "Humidifier", ProjectFileName = "Humidifier.ipj", AssemblyFileName = "Humidifier.iam", VaultPath = "$/Engineering/Library/Equipment/Humidifier" },
            new EquipmentItem { Name = "IsolatorPlate", DisplayName = "Isolator Plate", ProjectFileName = "PlateAssy.ipj", AssemblyFileName = "PlateAssy.iam", VaultPath = "$/Engineering/Library/Equipment/IsolatorPlate" },
            new EquipmentItem { Name = "OutletGuard", DisplayName = "Outlet Guard", ProjectFileName = "GuardAssy.ipj", AssemblyFileName = "GuardAssy.iam", VaultPath = "$/Engineering/Library/Equipment/OutletGuard" },
            new EquipmentItem { Name = "Rehausse", DisplayName = "Rehausse", ProjectFileName = "Rehausse.ipj", AssemblyFileName = "Rehausse.iam", VaultPath = "$/Engineering/Library/Equipment/Rehausse" },
            new EquipmentItem { Name = "RemovablePanel", DisplayName = "Removable Panel", ProjectFileName = "Removable Panel.ipj", AssemblyFileName = "Removable Panel.iam", VaultPath = "$/Engineering/Library/Equipment/Removable Panel" },
            new EquipmentItem { Name = "Silencer", DisplayName = "Silencer", ProjectFileName = "Silencer.ipj", AssemblyFileName = "Silencer.iam", VaultPath = "$/Engineering/Library/Equipment/Silencer" },
            new EquipmentItem { Name = "Transition", DisplayName = "Transition", ProjectFileName = "Transition.ipj", AssemblyFileName = "Transition.iam", VaultPath = "$/Engineering/Library/Equipment/Transition" },
            new EquipmentItem { Name = "VicWest", DisplayName = "Vic West", ProjectFileName = "VicWest.ipj", AssemblyFileName = "VicWest.idw", VaultPath = "$/Engineering/Library/Equipment/VicWest" },
            new EquipmentItem { Name = "XnHoodAssy", DisplayName = "XNRGY Hood Assembly", ProjectFileName = "HoodAssembly.ipj", AssemblyFileName = "HoodAssembly.iam", VaultPath = "$/Engineering/Library/Equipment/XnHoodAssy" },
            new EquipmentItem { Name = "XnrgyDoor", DisplayName = "XNRGY Door", ProjectFileName = "Xnrgy_Door.ipj", AssemblyFileName = "Xnrgy_Door.iam", VaultPath = "$/Engineering/Library/Equipment/XnrgyDoor" }
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

