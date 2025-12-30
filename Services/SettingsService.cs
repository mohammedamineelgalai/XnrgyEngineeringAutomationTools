using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using XnrgyEngineeringAutomationTools.Models;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service de gestion des paramètres de l'application
    /// Charge et sauvegarde les configurations depuis/vers un fichier JSON
    /// </summary>
    public static class SettingsService
    {
        private static readonly string SettingsFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "modulesettings.json");

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        private static ModuleSettings _currentSettings;
        private static readonly object _lock = new object();

        /// <summary>
        /// Paramètres actuels de l'application
        /// </summary>
        public static ModuleSettings Current
        {
            get
            {
                if (_currentSettings == null)
                {
                    lock (_lock)
                    {
                        if (_currentSettings == null)
                        {
                            _currentSettings = Load();
                        }
                    }
                }
                return _currentSettings;
            }
        }

        /// <summary>
        /// Charge les paramètres depuis le fichier JSON
        /// </summary>
        public static ModuleSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    string json = File.ReadAllText(SettingsFilePath);
                    var settings = JsonSerializer.Deserialize<ModuleSettings>(json, JsonOptions);
                    if (settings != null)
                    {
                        Logger.Log("[+] Parametres charges depuis modulesettings.json", Logger.LogLevel.INFO);
                        return settings;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur chargement parametres: {ex.Message}", Logger.LogLevel.WARNING);
            }

            // Retourner les valeurs par défaut
            Logger.Log("[i] Utilisation des parametres par defaut", Logger.LogLevel.INFO);
            return new ModuleSettings();
        }

        /// <summary>
        /// Sauvegarde les paramètres dans le fichier JSON
        /// </summary>
        public static bool Save(ModuleSettings settings)
        {
            try
            {
                string json = JsonSerializer.Serialize(settings, JsonOptions);
                File.WriteAllText(SettingsFilePath, json);
                
                // Mettre à jour le cache
                lock (_lock)
                {
                    _currentSettings = settings;
                }
                
                Logger.Log("[+] Parametres sauvegardes dans modulesettings.json", Logger.LogLevel.INFO);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur sauvegarde parametres: {ex.Message}", Logger.LogLevel.ERROR);
                return false;
            }
        }

        /// <summary>
        /// Recharge les paramètres depuis le fichier
        /// </summary>
        public static void Reload()
        {
            lock (_lock)
            {
                _currentSettings = Load();
            }
        }

        /// <summary>
        /// Réinitialise les paramètres aux valeurs par défaut
        /// </summary>
        public static void ResetToDefaults()
        {
            var defaults = new ModuleSettings();
            Save(defaults);
        }
    }

    /// <summary>
    /// Conteneur de tous les paramètres des modules
    /// </summary>
    public class ModuleSettings
    {
        /// <summary>
        /// Paramètres du module "Créer Module"
        /// </summary>
        public CreateModuleSettings CreateModule { get; set; } = new CreateModuleSettings();

        // Futurs modules:
        // public VaultUploadSettings VaultUpload { get; set; } = new VaultUploadSettings();
        // public SmartToolsSettings SmartTools { get; set; } = new SmartToolsSettings();
    }
}
