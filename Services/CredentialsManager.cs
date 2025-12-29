using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Gestionnaire sécurisé des credentials Vault
    /// Stocke les identifiants dans un fichier séparé du code source
    /// avec chiffrement basique pour protection
    /// </summary>
    public class CredentialsManager
    {
        // Fichier stocké dans AppData pour ne pas être perdu en recompilant
        private static readonly string AppDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "XnrgyEngineeringAutomationTools");
        
        private static readonly string CredentialsFile = Path.Combine(AppDataFolder, "vault-credentials.dat");
        
        // Clé de chiffrement simple (non sécurisé pour production, mais protège contre lecture directe)
        private static readonly byte[] EntropyBytes = Encoding.UTF8.GetBytes("XNRGY-VAT-2025-MAE");

        /// <summary>
        /// Modèle pour les credentials sauvegardés
        /// </summary>
        public class VaultCredentials
        {
            public string Server { get; set; } = "VAULTPOC";
            public string VaultName { get; set; } = "TestXNRGY";
            public string Username { get; set; } = "";
            public string Password { get; set; } = "";
            public bool SaveCredentials { get; set; } = false;
            public DateTime LastSaved { get; set; } = DateTime.MinValue;
        }

        /// <summary>
        /// Charge les credentials sauvegardés
        /// </summary>
        public static VaultCredentials Load()
        {
            try
            {
                if (!File.Exists(CredentialsFile))
                {
                    Logger.Log("[i] Aucun fichier credentials trouvé, utilisation valeurs par défaut", Logger.LogLevel.DEBUG);
                    return new VaultCredentials();
                }

                // Lire et déchiffrer
                byte[] encryptedData = File.ReadAllBytes(CredentialsFile);
                byte[] decryptedData = ProtectedData.Unprotect(encryptedData, EntropyBytes, DataProtectionScope.CurrentUser);
                string json = Encoding.UTF8.GetString(decryptedData);

                var credentials = JsonSerializer.Deserialize<VaultCredentials>(json);
                if (credentials != null)
                {
                    Logger.Log($"[+] Credentials chargés (Serveur: {credentials.Server}, Vault: {credentials.VaultName})", Logger.LogLevel.DEBUG);
                    return credentials;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur chargement credentials: {ex.Message}", Logger.LogLevel.DEBUG);
            }

            return new VaultCredentials();
        }

        /// <summary>
        /// Sauvegarde les credentials de façon sécurisée
        /// </summary>
        public static bool Save(VaultCredentials credentials)
        {
            try
            {
                // Créer le dossier si nécessaire
                if (!Directory.Exists(AppDataFolder))
                {
                    Directory.CreateDirectory(AppDataFolder);
                    Logger.Log($"[i] Dossier créé: {AppDataFolder}", Logger.LogLevel.DEBUG);
                }

                credentials.LastSaved = DateTime.Now;

                // Sérialiser et chiffrer
                string json = JsonSerializer.Serialize(credentials, new JsonSerializerOptions { WriteIndented = true });
                byte[] data = Encoding.UTF8.GetBytes(json);
                byte[] encryptedData = ProtectedData.Protect(data, EntropyBytes, DataProtectionScope.CurrentUser);

                File.WriteAllBytes(CredentialsFile, encryptedData);
                
                Logger.Log($"[+] Credentials sauvegardés dans {CredentialsFile}", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur sauvegarde credentials: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Efface les credentials sauvegardés
        /// </summary>
        public static bool Clear()
        {
            try
            {
                if (File.Exists(CredentialsFile))
                {
                    File.Delete(CredentialsFile);
                    Logger.Log("[i] Credentials effacés", Logger.LogLevel.INFO);
                }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur effacement credentials: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Vérifie si des credentials sont sauvegardés
        /// </summary>
        public static bool HasSavedCredentials()
        {
            return File.Exists(CredentialsFile);
        }

        /// <summary>
        /// Obtient le chemin du fichier credentials (pour info/debug)
        /// </summary>
        public static string GetCredentialsFilePath()
        {
            return CredentialsFile;
        }
    }
}
