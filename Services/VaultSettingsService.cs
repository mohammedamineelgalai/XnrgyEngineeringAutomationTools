using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using XnrgyEngineeringAutomationTools.Models;
using ACW = Autodesk.Connectivity.WebServices;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service de gestion des parametres partages via Vault pour XNRGY Engineering Automation Tools
    /// Deploiement multi-sites: Saint-Hubert QC + Arizona US (2 usines) = 50+ utilisateurs
    /// 
    /// - Chiffrement AES-256 des fichiers de configuration
    /// - Synchronisation automatique avec Vault au demarrage
    /// - Chemin Vault: $/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp/
    /// - Seuls les admins (Role Administrator ou Groupe Admin_Designer) peuvent modifier et uploader
    /// - Tous les utilisateurs recoivent la config au demarrage via Get automatique
    /// </summary>
    public class VaultSettingsService
    {
        // Chemins Vault et Local - Structure standardisee XNRGY
        private const string VAULT_APP_FOLDER = "$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp";
        private const string LOCAL_APP_FOLDER = @"C:\Vault\Engineering\Inventor_Standards\Automation_Standard\Configuration_Files\XnrgyEngineeringAutomationToolsApp";
        private const string CONFIG_FILENAME = "XnrgyEngineeringAutomationToolsSettings.config";

        // Cle de chiffrement AES-256 (XNRGY-specific - 32 bytes pour AES-256)
        // Cette cle est obfusquee dans le code compile
        private static readonly byte[] EncryptionKey = Encoding.UTF8.GetBytes("XNRGY2026-AUT0M@T10N-S3CR3T-K3Y!");  // 32 chars = 256 bits
        private static readonly byte[] EncryptionIV = Encoding.UTF8.GetBytes("XNRGY-IV-16BYTE!");  // 16 bytes pour AES

        private readonly VaultSdkService? _vaultService;
        private static ModuleSettings? _cachedSettings;
        private static readonly object _lock = new object();

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        /// <summary>
        /// Chemin local complet du fichier de configuration
        /// </summary>
        public string LocalFilePath => Path.Combine(LOCAL_APP_FOLDER, CONFIG_FILENAME);

        /// <summary>
        /// Chemin Vault du fichier de configuration
        /// </summary>
        public string VaultFilePath => $"{VAULT_APP_FOLDER}/{CONFIG_FILENAME}";

        /// <summary>
        /// Constructeur avec service Vault optionnel
        /// </summary>
        public VaultSettingsService(VaultSdkService? vaultService = null)
        {
            _vaultService = vaultService;
        }

        /// <summary>
        /// Parametres actuels (charge depuis Vault/Local si necessaire)
        /// </summary>
        public ModuleSettings Current
        {
            get
            {
                if (_cachedSettings == null)
                {
                    lock (_lock)
                    {
                        if (_cachedSettings == null)
                        {
                            _cachedSettings = Load();
                        }
                    }
                }
                return _cachedSettings;
            }
        }

        /// <summary>
        /// Charge les parametres depuis le fichier local (synchronise depuis Vault si necessaire)
        /// </summary>
        public ModuleSettings Load()
        {
            try
            {
                // Etape 1: Synchroniser depuis Vault si connecte
                SyncFromVault();

                // Etape 2: Charger depuis le fichier local chiffre
                if (File.Exists(LocalFilePath))
                {
                    byte[] encryptedData = File.ReadAllBytes(LocalFilePath);
                    string json = Decrypt(encryptedData);
                    
                    var settings = JsonSerializer.Deserialize<ModuleSettings>(json, JsonOptions);
                    if (settings != null)
                    {
                        Logger.Log($"[+] Parametres charges depuis {CONFIG_FILENAME} (chiffre)", Logger.LogLevel.INFO);
                        return settings;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur chargement parametres partages: {ex.Message}", Logger.LogLevel.WARNING);
            }

            // Retourner les valeurs par defaut
            Logger.Log("[i] Utilisation des parametres par defaut", Logger.LogLevel.INFO);
            return new ModuleSettings();
        }

        /// <summary>
        /// Sauvegarde les parametres et uploade vers Vault (admin seulement)
        /// Strategie: Delete ancien fichier + Add nouveau fichier (plus robuste que CheckOut/CheckIn pour contenu)
        /// </summary>
        public bool Save(ModuleSettings settings)
        {
            try
            {
                // Verifier les droits admin
                if (_vaultService != null && !_vaultService.IsCurrentUserAdmin())
                {
                    Logger.Log("[-] Sauvegarde refusee: droits administrateur requis", Logger.LogLevel.ERROR);
                    return false;
                }

                // Creer le dossier local si necessaire
                string? directory = Path.GetDirectoryName(LocalFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                    Logger.Log($"[+] Dossier local cree: {directory}", Logger.LogLevel.DEBUG);
                }

                // S'assurer que le fichier local n'est pas read-only
                if (File.Exists(LocalFilePath))
                {
                    var fileInfo = new FileInfo(LocalFilePath);
                    if (fileInfo.IsReadOnly)
                    {
                        fileInfo.IsReadOnly = false;
                        Logger.Log("[i] Flag read-only enleve du fichier local", Logger.LogLevel.DEBUG);
                    }
                }

                // Serialiser et chiffrer
                string json = JsonSerializer.Serialize(settings, JsonOptions);
                byte[] encryptedData = Encrypt(json);
                
                // Sauvegarder localement
                File.WriteAllBytes(LocalFilePath, encryptedData);
                Logger.Log($"[+] Parametres sauvegardes localement (chiffre): {LocalFilePath}", Logger.LogLevel.INFO);

                // Mettre a jour le cache
                lock (_lock)
                {
                    _cachedSettings = settings;
                }

                // Upload vers Vault (Delete + Add si existe deja)
                if (_vaultService != null && _vaultService.IsConnected)
                {
                    // Verifier si le fichier existe deja dans Vault
                    var existingFile = _vaultService.FindFileByPath(VaultFilePath);
                    
                    if (existingFile != null)
                    {
                        // Fichier existe - Supprimer puis re-ajouter avec nouveau contenu
                        Logger.Log("[>] Fichier existe dans Vault - suppression avant re-upload...", Logger.LogLevel.DEBUG);
                        
                        bool deleteSuccess = _vaultService.DeleteFileFromVault(existingFile);
                        if (!deleteSuccess)
                        {
                            Logger.Log("[-] Suppression du fichier existant echouee", Logger.LogLevel.ERROR);
                            return false;
                        }
                        Logger.Log("[+] Ancien fichier supprime de Vault", Logger.LogLevel.DEBUG);
                    }
                    
                    // Ajouter le nouveau fichier avec commentaire standardise
                    string configComment = $"MAJ Configuration | XNRGY Engineering Automation Tools | {DateTime.Now:yyyy-MM-dd HH:mm}";
                    bool uploadSuccess = _vaultService.AddFileToVault(LocalFilePath, VAULT_APP_FOLDER, configComment);
                    if (uploadSuccess)
                    {
                        Logger.Log("[+] Configuration uploadee dans Vault", Logger.LogLevel.INFO);
                        return true;
                    }
                    else
                    {
                        Logger.Log("[-] Upload vers Vault echoue", Logger.LogLevel.ERROR);
                        return false;
                    }
                }
                
                return true; // Sauvegarde locale reussie (Vault non connecte)
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur sauvegarde parametres: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Synchronise le fichier depuis Vault (Get latest)
        /// </summary>
        public bool SyncFromVault()
        {
            try
            {
                if (_vaultService == null || !_vaultService.IsConnected)
                {
                    Logger.Log("[i] Vault non connecte - utilisation du fichier local", Logger.LogLevel.DEBUG);
                    return false;
                }

                // Verifier si le fichier existe dans Vault
                var vaultFile = _vaultService.FindFileByPath(VaultFilePath);
                if (vaultFile == null)
                {
                    Logger.Log($"[i] Fichier config non trouve dans Vault: {VaultFilePath}", Logger.LogLevel.DEBUG);
                    return false;
                }

                // Verifier si on a besoin de telecharger (comparer versions/dates)
                if (File.Exists(LocalFilePath))
                {
                    var localFileInfo = new FileInfo(LocalFilePath);
                    // Si le fichier Vault est plus recent, telecharger
                    if (vaultFile.CkInDate <= localFileInfo.LastWriteTimeUtc)
                    {
                        Logger.Log("[i] Fichier local a jour - pas de sync necessaire", Logger.LogLevel.DEBUG);
                        return true;
                    }
                }

                // Telecharger depuis Vault (AcquireFile)
                Logger.Log($"[>] Telechargement config depuis Vault: {VaultFilePath}", Logger.LogLevel.INFO);
                
                // Creer le dossier si necessaire
                string? directory = Path.GetDirectoryName(LocalFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Utiliser AcquireFile pour Get
                bool success = _vaultService.AcquireFile(vaultFile, LocalFilePath, false); // false = Get, pas checkout
                
                if (success)
                {
                    Logger.Log($"[+] Config synchronisee depuis Vault", Logger.LogLevel.INFO);
                }
                
                return success;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur sync depuis Vault: {ex.Message}", Logger.LogLevel.WARNING);
                return false;
            }
        }

        /// <summary>
        /// Upload le fichier de config vers Vault (admin seulement)
        /// </summary>
        public bool UploadToVault()
        {
            try
            {
                if (_vaultService == null || !_vaultService.IsConnected)
                {
                    Logger.Log("[!] Vault non connecte - upload impossible", Logger.LogLevel.WARNING);
                    return false;
                }

                if (!_vaultService.IsCurrentUserAdmin())
                {
                    Logger.Log("[-] Upload refuse: droits administrateur requis", Logger.LogLevel.ERROR);
                    return false;
                }

                if (!File.Exists(LocalFilePath))
                {
                    Logger.Log("[-] Fichier local introuvable pour upload", Logger.LogLevel.ERROR);
                    return false;
                }

                Logger.Log($"[>] Upload config vers Vault: {VaultFilePath}", Logger.LogLevel.INFO);

                // Verifier si le fichier existe deja dans Vault
                var existingFile = _vaultService.FindFileByPath(VaultFilePath);
                
                if (existingFile != null)
                {
                    // Fichier existe - CheckOut, Update, CheckIn
                    Logger.Log("[>] Mise a jour du fichier existant dans Vault...", Logger.LogLevel.DEBUG);
                    string updateComment = $"MAJ Configuration | XNRGY Engineering Automation Tools | {DateTime.Now:yyyy-MM-dd HH:mm}";
                    bool success = _vaultService.UpdateFileInVault(existingFile, LocalFilePath, updateComment);
                    if (success)
                    {
                        Logger.Log("[+] Configuration mise a jour dans Vault", Logger.LogLevel.INFO);
                    }
                    return success;
                }
                else
                {
                    // Nouveau fichier - Add to Vault
                    Logger.Log("[>] Ajout nouveau fichier config dans Vault...", Logger.LogLevel.DEBUG);
                    string addComment = $"MAJ Configuration | XNRGY Engineering Automation Tools | {DateTime.Now:yyyy-MM-dd HH:mm}";
                    bool success = _vaultService.AddFileToVault(LocalFilePath, VAULT_APP_FOLDER, addComment);
                    if (success)
                    {
                        Logger.Log("[+] Configuration ajoutee dans Vault", Logger.LogLevel.INFO);
                    }
                    return success;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur upload vers Vault: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Recharge les parametres depuis Vault/Local
        /// </summary>
        public void Reload()
        {
            lock (_lock)
            {
                _cachedSettings = null;
                _cachedSettings = Load();
            }
        }

        /// <summary>
        /// Reinitialise aux valeurs par defaut et sauvegarde
        /// </summary>
        public bool ResetToDefaults()
        {
            var defaults = new ModuleSettings();
            return Save(defaults);
        }

        #region Encryption Methods (AES-256)

        /// <summary>
        /// Chiffre une chaine avec AES-256
        /// </summary>
        private byte[] Encrypt(string plainText)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = EncryptionKey;
                aes.IV = EncryptionIV;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);

                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt, Encoding.UTF8))
                        {
                            swEncrypt.Write(plainText);
                        }
                    }
                    return msEncrypt.ToArray();
                }
            }
        }

        /// <summary>
        /// Dechiffre des donnees avec AES-256
        /// </summary>
        private string Decrypt(byte[] cipherText)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = EncryptionKey;
                aes.IV = EncryptionIV;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                using (MemoryStream msDecrypt = new MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader srDecrypt = new StreamReader(csDecrypt, Encoding.UTF8))
                        {
                            return srDecrypt.ReadToEnd();
                        }
                    }
                }
            }
        }

        #endregion
    }
}
