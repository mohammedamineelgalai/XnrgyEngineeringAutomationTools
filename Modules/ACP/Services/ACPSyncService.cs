using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using XnrgyEngineeringAutomationTools.Modules.ACP.Models;
using XnrgyEngineeringAutomationTools.Services;

namespace XnrgyEngineeringAutomationTools.Modules.ACP.Services
{
    /// <summary>
    /// Service de synchronisation bidirectionnelle ACP avec Vault
    /// - Upload/écraser les validations toutes les 4-5 minutes
    /// - Téléchargement des validations des autres utilisateurs
    /// - Structure : Unité → Modules → Points Critiques
    /// </summary>
    public class ACPSyncService
    {
        private readonly VaultSdkService _vaultService;
        private readonly string _localDataFolder;
        private readonly int _syncIntervalMinutes = 4;
        private System.Threading.Timer? _syncTimer;
        private bool _isSyncing = false;
        private readonly object _syncLock = new object();

        private const string VAULT_ACP_FOLDER = "$/Engineering/Inventor_Standards/Automation_Standard/ACP_Data";

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        public event Action<string, string>? SyncStatusChanged;
        public event Action<string>? SyncCompleted;

        public ACPSyncService(VaultSdkService vaultService)
        {
            _vaultService = vaultService ?? throw new ArgumentNullException(nameof(vaultService));
            
            _localDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XnrgyEngineeringAutomationTools",
                "ACP"
            );

            if (!Directory.Exists(_localDataFolder))
            {
                Directory.CreateDirectory(_localDataFolder);
            }
        }

        /// <summary>
        /// Démarre la synchronisation automatique toutes les 4-5 minutes
        /// </summary>
        public void StartAutoSync()
        {
            if (_syncTimer != null)
            {
                StopAutoSync();
            }

            _syncTimer = new System.Threading.Timer(
                async _ => await PerformSyncAllUnitsAsync(),
                null,
                TimeSpan.Zero,
                TimeSpan.FromMinutes(_syncIntervalMinutes)
            );

            OnSyncStatusChanged("INFO", "Synchronisation automatique démarrée (toutes les 4 minutes)");
            Logger.Log("[ACPSync] Synchronisation automatique démarrée", Logger.LogLevel.INFO);
        }

        public void StopAutoSync()
        {
            _syncTimer?.Dispose();
            _syncTimer = null;
            OnSyncStatusChanged("INFO", "Synchronisation automatique arrêtée");
            Logger.Log("[ACPSync] Synchronisation automatique arrêtée", Logger.LogLevel.INFO);
        }

        /// <summary>
        /// Synchronise une unité ACP spécifique (bidirectionnel)
        /// </summary>
        public async Task<bool> SyncUnitAsync(string unitId, ACPDataModel? localData = null)
        {
            if (!_vaultService.IsConnected)
            {
                OnSyncStatusChanged("ERROR", "Connexion Vault requise pour la synchronisation");
                return false;
            }

            lock (_syncLock)
            {
                if (_isSyncing)
                {
                    OnSyncStatusChanged("WARN", "Synchronisation déjà en cours, veuillez patienter...");
                    return false;
                }
                _isSyncing = true;
            }

            try
            {
                OnSyncStatusChanged("INFO", $"Synchronisation de l'unité {unitId}...");

                // Étape 1: Télécharger la version Vault (si existe)
                ACPDataModel? vaultData = await DownloadFromVaultAsync(unitId);

                // Étape 2: Charger les données locales
                if (localData == null)
                {
                    localData = LoadLocalData(unitId);
                }

                // Étape 3: Fusionner (dernier modifié gagne)
                ACPDataModel mergedData = MergeData(localData, vaultData, unitId);

                // Étape 4: Upload vers Vault
                bool uploadSuccess = await UploadToVaultAsync(mergedData);

                if (uploadSuccess)
                {
                    SaveLocalCache(mergedData);
                    OnSyncStatusChanged("SUCCESS", $"Unité {unitId} synchronisée avec succès");
                    Logger.Log($"[ACPSync] Unité {unitId} synchronisée (Version: {mergedData.Version})", Logger.LogLevel.INFO);
                    return true;
                }
                else
                {
                    OnSyncStatusChanged("ERROR", $"Échec de l'upload de l'unité {unitId}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPSync.SyncUnitAsync", ex, Logger.LogLevel.ERROR);
                OnSyncStatusChanged("ERROR", $"Erreur de synchronisation: {ex.Message}");
                return false;
            }
            finally
            {
                lock (_syncLock)
                {
                    _isSyncing = false;
                }
            }
        }

        /// <summary>
        /// Synchronise la validation d'un module spécifique
        /// </summary>
        public async Task<bool> SyncModuleValidationAsync(string unitId, string moduleId, ModuleValidationData validationData)
        {
            try
            {
                // Charger l'unité complète
                var unitData = LoadLocalData(unitId);
                if (unitData == null)
                {
                    // Télécharger depuis Vault
                    unitData = await DownloadFromVaultAsync(unitId);
                    if (unitData == null)
                    {
                        OnSyncStatusChanged("ERROR", $"Unité {unitId} non trouvée");
                        return false;
                    }
                }

                // Mettre à jour les validations du module
                if (unitData.Modules.ContainsKey(moduleId))
                {
                    var module = unitData.Modules[moduleId];
                    
                    // Appliquer les validations aux points critiques
                    foreach (var validation in validationData.ValidatedPoints)
                    {
                        var point = module.CriticalPoints.FirstOrDefault(p => p.Id == validation.Key);
                        if (point != null)
                        {
                            var valid = validation.Value;
                            point.IsValidated = valid.IsValidated;
                            point.ValidatedBy = valid.ValidatedBy;
                            point.ValidatedDate = valid.ValidatedDate;
                            point.ValidationComment = valid.Comment;
                            point.IsApproved = valid.IsApproved;
                            point.ApprovedBy = valid.ApprovedBy;
                            point.ApprovedDate = valid.ApprovedDate;
                            point.ApprovalComment = valid.ApprovalComment;
                        }
                    }

                    // Mettre à jour le statut du module
                    module.Status = validationData.ModuleStatus;
                    module.ValidatedBy = validationData.LastValidatedBy;
                    module.ValidatedDate = validationData.LastValidatedDate;
                }

                // Synchroniser l'unité complète
                unitData.LastModifiedDate = DateTime.UtcNow;
                unitData.LastModifiedBy = _vaultService.UserName ?? "Unknown";
                unitData.Version++;

                return await SyncUnitAsync(unitId, unitData);
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPSync.SyncModuleValidationAsync", ex, Logger.LogLevel.ERROR);
                return false;
            }
        }

        private async Task PerformSyncAllUnitsAsync()
        {
            try
            {
                Logger.Log("[ACPSync] Démarrage synchronisation automatique de toutes les unités...", Logger.LogLevel.INFO);
                OnSyncStatusChanged("INFO", "Synchronisation automatique en cours...");

                var localFiles = Directory.GetFiles(_localDataFolder, "*.json", SearchOption.TopDirectoryOnly);

                int successCount = 0;
                int failCount = 0;

                foreach (var filePath in localFiles)
                {
                    try
                    {
                        var localData = LoadFromFile(filePath);
                        if (localData != null)
                        {
                            bool success = await SyncUnitAsync(localData.UnitId, localData);
                            if (success) successCount++;
                            else failCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[ACPSync] Erreur synchronisation {Path.GetFileName(filePath)}: {ex.Message}", Logger.LogLevel.ERROR);
                        failCount++;
                    }
                }

                OnSyncStatusChanged("SUCCESS", 
                    $"Synchronisation terminée: {successCount} réussi(s), {failCount} échec(s)");
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPSync.PerformSyncAllUnitsAsync", ex, Logger.LogLevel.ERROR);
                OnSyncStatusChanged("ERROR", $"Erreur synchronisation automatique: {ex.Message}");
            }
        }

        private async Task<ACPDataModel?> DownloadFromVaultAsync(string unitId)
        {
            try
            {
                string fileName = $"ACP_{unitId}.json";
                string vaultFilePath = $"{VAULT_ACP_FOLDER}/{fileName}";

                var vaultFile = await Task.Run(() => _vaultService.FindFileByPath(vaultFilePath));

                if (vaultFile == null)
                {
                    Logger.Log($"[ACPSync] Fichier Vault non trouvé: {vaultFilePath}", Logger.LogLevel.DEBUG);
                    return null;
                }

                // Télécharger le fichier vers un emplacement temporaire
                string tempFilePath = Path.Combine(_localDataFolder, $"temp_{fileName}");
                bool downloaded = await Task.Run(() => _vaultService.AcquireFile(vaultFile, tempFilePath, false));

                if (downloaded && File.Exists(tempFilePath))
                {
                    string json = File.ReadAllText(tempFilePath, Encoding.UTF8);
                    var data = JsonSerializer.Deserialize<ACPDataModel>(json, JsonOptions);
                    
                    // Supprimer le fichier temporaire
                    try { File.Delete(tempFilePath); } catch { }
                    
                    Logger.Log($"[ACPSync] Téléchargé depuis Vault: {fileName} (Version: {data?.Version ?? 0})", 
                        Logger.LogLevel.INFO);
                    
                    return data;
                }

                return null;
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPSync.DownloadFromVaultAsync", ex, Logger.LogLevel.ERROR);
                return null;
            }
        }

        private async Task<bool> UploadToVaultAsync(ACPDataModel data)
        {
            try
            {
                string fileName = $"ACP_{data.UnitId}.json";
                string vaultFolderPath = VAULT_ACP_FOLDER;

                string json = JsonSerializer.Serialize(data, JsonOptions);
                byte[] jsonBytes = Encoding.UTF8.GetBytes(json);

                string tempFilePath = Path.Combine(_localDataFolder, $"upload_{fileName}");
                File.WriteAllText(tempFilePath, json, Encoding.UTF8);

                if (!await EnsureVaultFolderExistsAsync(vaultFolderPath))
                {
                    Logger.Log($"[ACPSync] Impossible de créer le dossier Vault: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                string vaultFilePath = $"{vaultFolderPath}/{fileName}";

                bool success = await Task.Run(() =>
                {
                    try
                    {
                        // Utiliser UploadFileWithoutLifecycle pour les fichiers JSON (pas de lifecycle requis)
                        return _vaultService.UploadFileWithoutLifecycle(
                            tempFilePath,
                            vaultFolderPath,
                            null, null, null
                        );
                    }
                    catch (Exception ex)
                    {
                        Logger.LogException("ACPSync.UploadToVaultAsync", ex, Logger.LogLevel.ERROR);
                        return false;
                    }
                });

                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }

                if (success)
                {
                    Logger.Log($"[ACPSync] Upload réussi vers Vault: {fileName} (Version: {data.Version})", 
                        Logger.LogLevel.INFO);
                }

                return success;
            }
            catch (Exception ex)
            {
                Logger.LogException("ACPSync.UploadToVaultAsync", ex, Logger.LogLevel.ERROR);
                return false;
            }
        }

        private ACPDataModel MergeData(ACPDataModel? local, ACPDataModel? vault, string unitId)
        {
            if (local == null && vault == null)
            {
                return new ACPDataModel
                {
                    UnitId = unitId,
                    LastModifiedBy = _vaultService.UserName ?? "Unknown",
                    LastModifiedDate = DateTime.UtcNow,
                    Version = 1,
                    Modules = new Dictionary<string, ACPModule>()
                };
            }

            if (vault == null)
            {
                local!.Version++;
                local.LastModifiedDate = DateTime.UtcNow;
                local.LastModifiedBy = _vaultService.UserName ?? "Unknown";
                return local;
            }

            if (local == null)
            {
                return vault;
            }

            // Comparer les dates (dernier modifié gagne)
            if (vault.LastModifiedDate > local.LastModifiedDate)
            {
                // Vault plus récent : utiliser Vault comme base
                // Mais fusionner les nouvelles validations locales
                var merged = new ACPDataModel
                {
                    UnitId = unitId,
                    ProjectNumber = vault.ProjectNumber,
                    Reference = vault.Reference,
                    UnitName = vault.UnitName,
                    CreatedBy = vault.CreatedBy,
                    CreatedDate = vault.CreatedDate,
                    LastModifiedBy = vault.LastModifiedBy,
                    LastModifiedDate = vault.LastModifiedDate,
                    Version = vault.Version,
                    Modules = new Dictionary<string, ACPModule>(vault.Modules)
                };

                // Fusionner les validations locales (par module)
                foreach (var localModule in local.Modules)
                {
                    if (merged.Modules.ContainsKey(localModule.Key))
                    {
                        var vaultModule = merged.Modules[localModule.Key];
                        
                        // Fusionner les validations des points critiques
                        foreach (var localPoint in localModule.Value.CriticalPoints)
                        {
                            var vaultPoint = vaultModule.CriticalPoints.FirstOrDefault(p => p.Id == localPoint.Id);
                            
                            if (vaultPoint == null)
                            {
                                vaultModule.CriticalPoints.Add(localPoint);
                            }
                            else if (localPoint.ValidatedDate.HasValue && 
                                    (!vaultPoint.ValidatedDate.HasValue || localPoint.ValidatedDate > vaultPoint.ValidatedDate))
                            {
                                // La validation locale est plus récente
                                vaultPoint.IsValidated = localPoint.IsValidated;
                                vaultPoint.ValidatedBy = localPoint.ValidatedBy;
                                vaultPoint.ValidatedDate = localPoint.ValidatedDate;
                                vaultPoint.ValidationComment = localPoint.ValidationComment;
                            }
                        }
                    }
                    else
                    {
                        merged.Modules[localModule.Key] = localModule.Value;
                    }
                }

                merged.Version++;
                merged.LastModifiedDate = DateTime.UtcNow;
                merged.LastModifiedBy = _vaultService.UserName ?? "Unknown";
                return merged;
            }
            else
            {
                // Local plus récent : utiliser local comme base
                local.Version = Math.Max(local.Version, vault.Version) + 1;
                local.LastModifiedDate = DateTime.UtcNow;
                local.LastModifiedBy = _vaultService.UserName ?? "Unknown";
                return local;
            }
        }

        private ACPDataModel? LoadLocalData(string unitId)
        {
            string filePath = Path.Combine(_localDataFolder, $"ACP_{unitId}.json");
            return LoadFromFile(filePath);
        }

        private ACPDataModel? LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            try
            {
                string json = File.ReadAllText(filePath, Encoding.UTF8);
                return JsonSerializer.Deserialize<ACPDataModel>(json, JsonOptions);
            }
            catch (Exception ex)
            {
                Logger.Log($"[ACPSync] Erreur chargement {filePath}: {ex.Message}", Logger.LogLevel.ERROR);
                return null;
            }
        }

        private void SaveLocalCache(ACPDataModel data)
        {
            try
            {
                string fileName = $"ACP_{data.UnitId}.json";
                string filePath = Path.Combine(_localDataFolder, fileName);
                string json = JsonSerializer.Serialize(data, JsonOptions);
                File.WriteAllText(filePath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Logger.Log($"[ACPSync] Erreur sauvegarde cache: {ex.Message}", Logger.LogLevel.ERROR);
            }
        }

        private async Task<bool> EnsureVaultFolderExistsAsync(string vaultPath)
        {
            return await Task.Run(async () =>
            {
                try
                {
                    bool exists = await _vaultService.GetFolderAsync(vaultPath);
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.LogException("ACPSync.EnsureVaultFolderExistsAsync", ex, Logger.LogLevel.ERROR);
                    return true;
                }
            });
        }

        private void OnSyncStatusChanged(string level, string message)
        {
            SyncStatusChanged?.Invoke(level, message);
        }

        public void Dispose()
        {
            StopAutoSync();
        }
    }
}

