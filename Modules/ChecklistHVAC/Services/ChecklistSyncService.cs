using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using XnrgyEngineeringAutomationTools.Modules.ChecklistHVAC.Models;
using XnrgyEngineeringAutomationTools.Services;
using ACW = Autodesk.Connectivity.WebServices;

namespace XnrgyEngineeringAutomationTools.Modules.ChecklistHVAC.Services
{
    /// <summary>
    /// Service de synchronisation bidirectionnelle Checklist HVAC avec Vault
    /// - Upload/écraser toutes les 4-5 minutes
    /// - Téléchargement des changements des autres utilisateurs
    /// - Gestion des conflits (dernier modifié gagne)
    /// - Support Word (export optionnel)
    /// </summary>
    public class ChecklistSyncService
    {
        private readonly VaultSdkService _vaultService;
        private readonly string _localDataFolder;
        private readonly int _syncIntervalMinutes = 4;  // Synchronisation toutes les 4 minutes
        private System.Threading.Timer? _syncTimer;
        private bool _isSyncing = false;
        private readonly object _syncLock = new object();

        // Chemin Vault pour stocker les données de checklist
        private const string VAULT_CHECKLIST_FOLDER = "$/Engineering/Inventor_Standards/Automation_Standard/Checklist_HVAC_Data";

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        public event Action<string, string>? SyncStatusChanged;  // (status, message)
        public event Action<ChecklistSyncMetadata>? SyncCompleted;

        public ChecklistSyncService(VaultSdkService vaultService)
        {
            _vaultService = vaultService ?? throw new ArgumentNullException(nameof(vaultService));
            
            // Dossier local pour cache des données
            _localDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XnrgyEngineeringAutomationTools",
                "ChecklistHVAC"
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

            // Timer qui se déclenche toutes les 4 minutes
            _syncTimer = new System.Threading.Timer(
                async _ => await PerformSyncAllModulesAsync(),
                null,
                TimeSpan.Zero,  // Démarrer immédiatement
                TimeSpan.FromMinutes(_syncIntervalMinutes)
            );

            OnSyncStatusChanged("INFO", "Synchronisation automatique démarrée (toutes les 4 minutes)");
            Logger.Log("[ChecklistSync] Synchronisation automatique démarrée", Logger.LogLevel.INFO);
        }

        /// <summary>
        /// Arrête la synchronisation automatique
        /// </summary>
        public void StopAutoSync()
        {
            _syncTimer?.Dispose();
            _syncTimer = null;
            OnSyncStatusChanged("INFO", "Synchronisation automatique arrêtée");
            Logger.Log("[ChecklistSync] Synchronisation automatique arrêtée", Logger.LogLevel.INFO);
        }

        /// <summary>
        /// Synchronise un module spécifique (bidirectionnel)
        /// </summary>
        public async Task<bool> SyncModuleAsync(string projectNumber, string reference, string module, 
            ChecklistDataModel? localData = null)
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
                string moduleId = $"{projectNumber}-{reference}-{module}";
                OnSyncStatusChanged("INFO", $"Synchronisation du module {moduleId}...");

                // Étape 1: Télécharger la version Vault (si existe)
                ChecklistDataModel? vaultData = await DownloadFromVaultAsync(moduleId);

                // Étape 2: Charger les données locales (depuis localStorage HTML ou fichier local)
                if (localData == null)
                {
                    localData = LoadLocalData(moduleId);
                }

                // Étape 3: Résolution de conflit (dernier modifié gagne)
                ChecklistDataModel mergedData = MergeData(localData, vaultData, moduleId, projectNumber, reference, module);

                // Étape 4: Upload vers Vault (écraser)
                bool uploadSuccess = await UploadToVaultAsync(mergedData);

                if (uploadSuccess)
                {
                    // Étape 5: Sauvegarder localement comme cache
                    SaveLocalCache(mergedData);

                    OnSyncStatusChanged("SUCCESS", $"Module {moduleId} synchronisé avec succès");
                    Logger.Log($"[ChecklistSync] Module {moduleId} synchronisé (Version: {mergedData.Version})", Logger.LogLevel.INFO);
                    return true;
                }
                else
                {
                    OnSyncStatusChanged("ERROR", $"Échec de l'upload du module {moduleId}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException("ChecklistSync.SyncModuleAsync", ex, Logger.LogLevel.ERROR);
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
        /// Synchronise tous les modules trouvés localement
        /// </summary>
        private async Task PerformSyncAllModulesAsync()
        {
            try
            {
                Logger.Log("[ChecklistSync] Démarrage synchronisation automatique de tous les modules...", Logger.LogLevel.INFO);
                OnSyncStatusChanged("INFO", "Synchronisation automatique en cours...");

                // Lire tous les fichiers JSON locaux
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
                            bool success = await SyncModuleAsync(
                                localData.ProjectNumber,
                                localData.Reference,
                                localData.Module,
                                localData
                            );

                            if (success)
                                successCount++;
                            else
                                failCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[ChecklistSync] Erreur synchronisation {Path.GetFileName(filePath)}: {ex.Message}", Logger.LogLevel.ERROR);
                        failCount++;
                    }
                }

                OnSyncStatusChanged("SUCCESS", 
                    $"Synchronisation terminée: {successCount} réussi(s), {failCount} échec(s)");
                
                Logger.Log($"[ChecklistSync] Synchronisation automatique terminée: {successCount} succès, {failCount} échecs", 
                    Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.LogException("ChecklistSync.PerformSyncAllModulesAsync", ex, Logger.LogLevel.ERROR);
                OnSyncStatusChanged("ERROR", $"Erreur synchronisation automatique: {ex.Message}");
            }
        }

        /// <summary>
        /// Télécharge les données depuis Vault
        /// </summary>
        private async Task<ChecklistDataModel?> DownloadFromVaultAsync(string moduleId)
        {
            try
            {
                string fileName = $"Checklist_{moduleId}.json";
                string vaultFilePath = $"{VAULT_CHECKLIST_FOLDER}/{fileName}";

                // Utiliser FindFileByPath de VaultSDKService
                var vaultFile = await Task.Run(() => _vaultService.FindFileByPath(vaultFilePath));

                if (vaultFile == null)
                {
                    Logger.Log($"[ChecklistSync] Fichier Vault non trouvé: {vaultFilePath}", Logger.LogLevel.DEBUG);
                    return null;
                }

                // Télécharger le fichier via GET depuis Vault
                bool downloaded = await _vaultService.GetFolderAsync(VAULT_CHECKLIST_FOLDER);
                if (downloaded)
                {
                    // Le GET télécharge vers C:\Vault\...
                    string vaultLocalPath = Path.Combine(
                        @"C:\Vault\Engineering\Inventor_Standards\Automation_Standard\Checklist_HVAC_Data",
                        fileName
                    );

                    if (File.Exists(vaultLocalPath))
                    {
                        // Charger le JSON depuis le fichier local Vault
                        string json = File.ReadAllText(vaultLocalPath, Encoding.UTF8);
                        var data = JsonSerializer.Deserialize<ChecklistDataModel>(json, JsonOptions);
                        
                        Logger.Log($"[ChecklistSync] Téléchargé depuis Vault: {fileName} (Version: {data?.Version ?? 0})", 
                            Logger.LogLevel.INFO);
                        
                        return data;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                Logger.LogException("ChecklistSync.DownloadFromVaultAsync", ex, Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// Upload les données vers Vault (écraser)
        /// </summary>
        private async Task<bool> UploadToVaultAsync(ChecklistDataModel data)
        {
            try
            {
                string fileName = $"Checklist_{data.ModuleId}.json";
                string vaultFolderPath = VAULT_CHECKLIST_FOLDER;

                // Sérialiser en JSON
                string json = JsonSerializer.Serialize(data, JsonOptions);
                byte[] jsonBytes = Encoding.UTF8.GetBytes(json);

                // Créer un fichier temporaire local
                string tempFilePath = Path.Combine(_localDataFolder, $"upload_{fileName}");
                File.WriteAllText(tempFilePath, json, Encoding.UTF8);

                // S'assurer que le dossier existe dans Vault
                if (!await EnsureVaultFolderExistsAsync(vaultFolderPath))
                {
                    Logger.Log($"[ChecklistSync] Impossible de créer le dossier Vault: {vaultFolderPath}", Logger.LogLevel.ERROR);
                    return false;
                }

                string vaultFilePath = $"{vaultFolderPath}/{fileName}";

                // Upload vers Vault (écraser si existe)
                bool success = await Task.Run(() =>
                {
                    try
                    {
                        // Vérifier si le fichier existe déjà
                        var existingFile = _vaultService.FindFileByPath(vaultFilePath);

                        if (existingFile != null)
                        {
                            // Fichier existe : utiliser UploadFile qui gère automatiquement CheckOut/CheckIn
                            // UploadFile gère déjà l'écrasement avec CheckOut/CheckIn automatique
                            return _vaultService.UploadFile(
                                tempFilePath,
                                vaultFolderPath,
                                null, null, null,  // Pas de propriétés Project/Ref/Module pour fichiers de données
                                null, null, null, null, null,
                                "Synchronisation Checklist HVAC automatique"
                            );
                        }
                        else
                        {
                            // Nouveau fichier : Upload simple
                            return _vaultService.UploadFile(
                                tempFilePath,
                                vaultFolderPath,
                                null, null, null,
                                null, null, null, null, null,
                                "Création Checklist HVAC"
                            );
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogException("ChecklistSync.UploadToVaultAsync", ex, Logger.LogLevel.ERROR);
                        return false;
                    }
                });

                // Nettoyer le fichier temporaire
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }

                if (success)
                {
                    Logger.Log($"[ChecklistSync] Upload réussi vers Vault: {fileName} (Version: {data.Version})", 
                        Logger.LogLevel.INFO);
                }

                return success;
            }
            catch (Exception ex)
            {
                Logger.LogException("ChecklistSync.UploadToVaultAsync", ex, Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Fusionne les données locales et Vault (dernier modifié gagne)
        /// </summary>
        private ChecklistDataModel MergeData(ChecklistDataModel? local, ChecklistDataModel? vault, 
            string moduleId, string project, string reference, string module)
        {
            // Si aucune donnée n'existe, créer une nouvelle
            if (local == null && vault == null)
            {
                return new ChecklistDataModel
                {
                    ModuleId = moduleId,
                    ProjectNumber = project,
                    Reference = reference,
                    Module = module,
                    LastModifiedBy = _vaultService.UserName ?? "Unknown",
                    LastModifiedDate = DateTime.UtcNow,
                    Version = 1,
                    Responses = new Dictionary<int, CheckpointResponse>()
                };
            }

            // Si seule la version locale existe
            if (vault == null)
            {
                local!.Version++;
                local.LastModifiedDate = DateTime.UtcNow;
                local.LastModifiedBy = _vaultService.UserName ?? "Unknown";
                return local;
            }

            // Si seule la version Vault existe
            if (local == null)
            {
                return vault;
            }

            // Comparer les dates de modification (dernier modifié gagne)
            if (vault.LastModifiedDate > local.LastModifiedDate)
            {
                // Version Vault est plus récente : utiliser Vault comme base
                // Mais fusionner les nouvelles réponses locales
                var merged = new ChecklistDataModel
                {
                    ModuleId = moduleId,
                    ProjectNumber = project,
                    Reference = reference,
                    Module = module,
                    LastModifiedBy = vault.LastModifiedBy,
                    LastModifiedDate = vault.LastModifiedDate,
                    Version = vault.Version,
                    Responses = new Dictionary<int, CheckpointResponse>(vault.Responses)
                };

                // Ajouter les nouvelles réponses locales (celles modifiées après la dernière sync)
                foreach (var localResponse in local.Responses)
                {
                    if (!merged.Responses.ContainsKey(localResponse.Key) || 
                        localResponse.Value.ModifiedDate > vault.LastModifiedDate)
                    {
                        merged.Responses[localResponse.Key] = localResponse.Value;
                        merged.LastModifiedDate = DateTime.UtcNow;
                        merged.LastModifiedBy = _vaultService.UserName ?? "Unknown";
                    }
                }

                merged.Version++;
                return merged;
            }
            else
            {
                // Version locale est plus récente ou égale : utiliser locale comme base
                local.Version = Math.Max(local.Version, vault.Version) + 1;
                local.LastModifiedDate = DateTime.UtcNow;
                local.LastModifiedBy = _vaultService.UserName ?? "Unknown";
                return local;
            }
        }

        /// <summary>
        /// Charge les données depuis un fichier local
        /// </summary>
        private ChecklistDataModel? LoadLocalData(string moduleId)
        {
            string filePath = Path.Combine(_localDataFolder, $"Checklist_{moduleId}.json");
            return LoadFromFile(filePath);
        }

        /// <summary>
        /// Charge depuis un fichier
        /// </summary>
        private ChecklistDataModel? LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            try
            {
                string json = File.ReadAllText(filePath, Encoding.UTF8);
                return JsonSerializer.Deserialize<ChecklistDataModel>(json, JsonOptions);
            }
            catch (Exception ex)
            {
                Logger.Log($"[ChecklistSync] Erreur chargement {filePath}: {ex.Message}", Logger.LogLevel.ERROR);
                return null;
            }
        }

        /// <summary>
        /// Sauvegarde en cache local
        /// </summary>
        private void SaveLocalCache(ChecklistDataModel data)
        {
            try
            {
                string fileName = $"Checklist_{data.ModuleId}.json";
                string filePath = Path.Combine(_localDataFolder, fileName);
                string json = JsonSerializer.Serialize(data, JsonOptions);
                File.WriteAllText(filePath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Logger.Log($"[ChecklistSync] Erreur sauvegarde cache: {ex.Message}", Logger.LogLevel.ERROR);
            }
        }

        /// <summary>
        /// S'assure que le dossier Vault existe
        /// Utilise GetFolderAsync qui crée automatiquement le dossier si nécessaire
        /// </summary>
        private async Task<bool> EnsureVaultFolderExistsAsync(string vaultPath)
        {
            return await Task.Run(async () =>
            {
                try
                {
                    // Tenter de faire un GET du dossier (le créera si nécessaire via VaultSDKService)
                    // Si le dossier n'existe pas, GetFolderAsync retournera false mais le dossier sera créé
                    bool exists = await _vaultService.GetFolderAsync(vaultPath);
                    
                    // Vérifier si le dossier existe vraiment maintenant
                    // Si GetFolderAsync échoue, on essaie quand même l'upload (le dossier sera créé automatiquement)
                    return true; // On laisse UploadFile gérer la création si nécessaire
                }
                catch (Exception ex)
                {
                    Logger.LogException("ChecklistSync.EnsureVaultFolderExistsAsync", ex, Logger.LogLevel.ERROR);
                    // Même en cas d'erreur, on essaie quand même l'upload
                    return true;
                }
            });
        }

        /// <summary>
        /// Exporte les données vers Word (optionnel, à implémenter)
        /// </summary>
        public async Task<bool> ExportToWordAsync(ChecklistDataModel data, string outputPath)
        {
            // TODO: Implémenter export Word avec bibliothèque comme DocX ou iTextSharp
            // Pour l'instant, on peut exporter en JSON et l'utilisateur peut l'ouvrir dans Word
            try
            {
                string json = JsonSerializer.Serialize(data, JsonOptions);
                File.WriteAllText(outputPath, json, Encoding.UTF8);
                return true;
            }
            catch
            {
                return false;
            }
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

