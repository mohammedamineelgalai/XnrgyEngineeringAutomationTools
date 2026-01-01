using System.Windows.Media;

namespace XnrgyEngineeringAutomationTools.Services
{
    /// <summary>
    /// Service centralise pour les couleurs des journaux dans toute l'application.
    /// Garantit une uniformite visuelle entre tous les modules.
    /// 
    /// PALETTE STANDARD XNRGY:
    /// - SUCCESS : Vert SpringGreen brillant (#00FF7F)
    /// - ERROR   : Rouge vif (#FF4444)
    /// - WARNING : Jaune or brillant (#FFD700)
    /// - INFO    : Blanc pur (#FFFFFF)
    /// - DEBUG   : Cyan brillant (#00FFFF)
    /// - TIMESTAMP : Gris moyen (#888888)
    /// </summary>
    public static class JournalColorService
    {
        // ====================================================================
        // COULEURS STANDARD - NE PAS MODIFIER SANS RAISON
        // ====================================================================
        
        /// <summary>Vert SpringGreen brillant pour les succes</summary>
        public static readonly Color SuccessColor = (Color)ColorConverter.ConvertFromString("#00FF7F");
        
        /// <summary>Rouge vif pour les erreurs</summary>
        public static readonly Color ErrorColor = (Color)ColorConverter.ConvertFromString("#FF4444");
        
        /// <summary>Jaune or brillant pour les avertissements</summary>
        public static readonly Color WarningColor = (Color)ColorConverter.ConvertFromString("#FFD700");
        
        /// <summary>Blanc pur pour les informations</summary>
        public static readonly Color InfoColor = (Color)ColorConverter.ConvertFromString("#FFFFFF");
        
        /// <summary>Cyan clair brillant pour les messages de debug</summary>
        public static readonly Color DebugColor = (Color)ColorConverter.ConvertFromString("#00FFFF");
        
        /// <summary>Gris moyen pour les timestamps</summary>
        public static readonly Color TimestampColor = (Color)ColorConverter.ConvertFromString("#888888");

        // ====================================================================
        // BRUSHES PRE-CREES POUR PERFORMANCE
        // ====================================================================
        
        private static SolidColorBrush? _successBrush;
        private static SolidColorBrush? _errorBrush;
        private static SolidColorBrush? _warningBrush;
        private static SolidColorBrush? _infoBrush;
        private static SolidColorBrush? _debugBrush;
        private static SolidColorBrush? _timestampBrush;

        /// <summary>Brush vert pour succes</summary>
        public static SolidColorBrush SuccessBrush => _successBrush ??= new SolidColorBrush(SuccessColor);
        
        /// <summary>Brush rouge pour erreurs</summary>
        public static SolidColorBrush ErrorBrush => _errorBrush ??= new SolidColorBrush(ErrorColor);
        
        /// <summary>Brush jaune pour avertissements</summary>
        public static SolidColorBrush WarningBrush => _warningBrush ??= new SolidColorBrush(WarningColor);
        
        /// <summary>Brush blanc pour infos</summary>
        public static SolidColorBrush InfoBrush => _infoBrush ??= new SolidColorBrush(InfoColor);
        
        /// <summary>Brush gris pour debug</summary>
        public static SolidColorBrush DebugBrush => _debugBrush ??= new SolidColorBrush(DebugColor);
        
        /// <summary>Brush gris pour timestamps</summary>
        public static SolidColorBrush TimestampBrush => _timestampBrush ??= new SolidColorBrush(TimestampColor);

        // ====================================================================
        // ENUM LOG LEVEL STANDARD
        // ====================================================================
        
        /// <summary>Niveaux de log standard pour les journaux UI</summary>
        public enum LogLevel
        {
            INFO,
            SUCCESS,
            WARNING,
            ERROR,
            DEBUG
        }

        // ====================================================================
        // METHODES UTILITAIRES
        // ====================================================================

        /// <summary>
        /// Retourne le brush correspondant au niveau de log
        /// </summary>
        public static SolidColorBrush GetBrushForLevel(LogLevel level)
        {
            return level switch
            {
                LogLevel.SUCCESS => SuccessBrush,
                LogLevel.ERROR => ErrorBrush,
                LogLevel.WARNING => WarningBrush,
                LogLevel.DEBUG => DebugBrush,
                _ => InfoBrush
            };
        }

        /// <summary>
        /// Retourne la couleur correspondant au niveau de log
        /// </summary>
        public static Color GetColorForLevel(LogLevel level)
        {
            return level switch
            {
                LogLevel.SUCCESS => SuccessColor,
                LogLevel.ERROR => ErrorColor,
                LogLevel.WARNING => WarningColor,
                LogLevel.DEBUG => DebugColor,
                _ => InfoColor
            };
        }

        /// <summary>
        /// Retourne le code hex de la couleur pour le niveau de log
        /// </summary>
        public static string GetHexColorForLevel(LogLevel level)
        {
            return level switch
            {
                LogLevel.SUCCESS => "#00FF7F",
                LogLevel.ERROR => "#FF4444",
                LogLevel.WARNING => "#FFD700",
                LogLevel.DEBUG => "#00FFFF",
                _ => "#FFFFFF"
            };
        }

        /// <summary>
        /// Retourne le prefixe standard pour le niveau de log
        /// </summary>
        public static string GetPrefixForLevel(LogLevel level)
        {
            return level switch
            {
                LogLevel.SUCCESS => "[+]",
                LogLevel.ERROR => "[-]",
                LogLevel.WARNING => "[!]",
                LogLevel.DEBUG => "[~]",
                _ => "[>]"
            };
        }

        /// <summary>
        /// Convertit le LogLevel interne vers Logger.LogLevel pour ecriture fichier
        /// </summary>
        public static Logger.LogLevel ToFileLogLevel(LogLevel level)
        {
            return level switch
            {
                LogLevel.SUCCESS => Logger.LogLevel.INFO,
                LogLevel.ERROR => Logger.LogLevel.ERROR,
                LogLevel.WARNING => Logger.LogLevel.WARNING,
                LogLevel.DEBUG => Logger.LogLevel.DEBUG,
                _ => Logger.LogLevel.INFO
            };
        }
    }
}
