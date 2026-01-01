using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Gestionnaire des preferences utilisateur (theme, parametres UI)
    /// Stocke les preferences dans AppData de facon securisee
    /// Compatible avec les standards Windows pour applications professionnelles
    /// </summary>
    public static class UserPreferencesManager
    {
        // Fichier stocke dans AppData Local
        private static readonly string AppDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "XnrgyEngineeringAutomationTools");
        
        private static readonly string PreferencesFile = Path.Combine(AppDataFolder, "user-preferences.json");
        
        // Cache en memoire pour acces rapide
        private static UserPreferences? _cachedPreferences;

        /// <summary>
        /// Modele pour les preferences utilisateur
        /// </summary>
        public class UserPreferences
        {
            // Theme
            public bool IsDarkTheme { get; set; } = true;
            
            // Parametres fenetre principale
            public double WindowWidth { get; set; } = 0;
            public double WindowHeight { get; set; } = 0;
            public bool IsMaximized { get; set; } = false;
            
            // Parametres generaux
            public bool AutoConnectVault { get; set; } = true;
            public bool AutoConnectInventor { get; set; } = true;
            public bool ShowStartupChecklist { get; set; } = true;
            
            // Metadonnees
            public string AppVersion { get; set; } = "1.0.0";
            public DateTime LastSaved { get; set; } = DateTime.MinValue;
            public string LastUser { get; set; } = "";
        }

        /// <summary>
        /// Charge les preferences utilisateur
        /// </summary>
        public static UserPreferences Load()
        {
            // Retourner le cache si disponible
            if (_cachedPreferences != null)
            {
                return _cachedPreferences;
            }

            try
            {
                if (!File.Exists(PreferencesFile))
                {
                    Logger.Log("[i] Aucun fichier preferences trouve, utilisation valeurs par defaut", Logger.LogLevel.DEBUG);
                    _cachedPreferences = new UserPreferences();
                    return _cachedPreferences;
                }

                // Lire le fichier JSON (non chiffre car pas de donnees sensibles)
                string json = File.ReadAllText(PreferencesFile, Encoding.UTF8);
                var preferences = JsonSerializer.Deserialize<UserPreferences>(json);
                
                if (preferences != null)
                {
                    Logger.Log($"[+] Preferences chargees (Theme: {(preferences.IsDarkTheme ? "Sombre" : "Clair")})", Logger.LogLevel.DEBUG);
                    _cachedPreferences = preferences;
                    return preferences;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur chargement preferences: {ex.Message}", Logger.LogLevel.DEBUG);
            }

            _cachedPreferences = new UserPreferences();
            return _cachedPreferences;
        }

        /// <summary>
        /// Sauvegarde les preferences utilisateur
        /// </summary>
        public static bool Save(UserPreferences preferences)
        {
            try
            {
                // Creer le dossier si necessaire
                if (!Directory.Exists(AppDataFolder))
                {
                    Directory.CreateDirectory(AppDataFolder);
                    Logger.Log($"[i] Dossier cree: {AppDataFolder}", Logger.LogLevel.DEBUG);
                }

                preferences.LastSaved = DateTime.Now;
                preferences.LastUser = Environment.UserName;

                // Serialiser en JSON lisible
                var options = new JsonSerializerOptions 
                { 
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                string json = JsonSerializer.Serialize(preferences, options);
                
                File.WriteAllText(PreferencesFile, json, Encoding.UTF8);
                
                // Mettre a jour le cache
                _cachedPreferences = preferences;
                
                Logger.Log($"[+] Preferences sauvegardees (Theme: {(preferences.IsDarkTheme ? "Sombre" : "Clair")})", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur sauvegarde preferences: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Sauvegarde uniquement le theme (raccourci)
        /// </summary>
        public static bool SaveTheme(bool isDarkTheme)
        {
            var prefs = Load();
            prefs.IsDarkTheme = isDarkTheme;
            return Save(prefs);
        }

        /// <summary>
        /// Charge uniquement le theme (raccourci)
        /// </summary>
        public static bool LoadTheme()
        {
            return Load().IsDarkTheme;
        }

        /// <summary>
        /// Reinitialise les preferences par defaut
        /// </summary>
        public static bool Reset()
        {
            try
            {
                if (File.Exists(PreferencesFile))
                {
                    File.Delete(PreferencesFile);
                }
                _cachedPreferences = null;
                Logger.Log("[i] Preferences reinitialisees", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur reinitialisation preferences: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Obtient le chemin du fichier preferences (pour info/debug)
        /// </summary>
        public static string GetPreferencesFilePath()
        {
            return PreferencesFile;
        }

        /// <summary>
        /// Verifie si le fichier preferences existe
        /// </summary>
        public static bool HasSavedPreferences()
        {
            return File.Exists(PreferencesFile);
        }

        /// <summary>
        /// Invalide le cache (force rechargement au prochain Load)
        /// </summary>
        public static void InvalidateCache()
        {
            _cachedPreferences = null;
        }
    }
}
