// ============================================================================
// DXFVerifierInfoWindow.xaml.cs
// DXF Verifier Info Window - Modern Info Dialog
// Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
// ============================================================================

using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Views
{
    /// <summary>
    /// Fenetre d'information moderne pour DXF Verifier
    /// Migration depuis VB.NET MainForm.vb ShowModernInfoDialog()
    /// </summary>
    public partial class DXFVerifierInfoWindow : Window
    {
        public DXFVerifierInfoWindow()
        {
            InitializeComponent();
            BuildContent();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BuildContent()
        {
            // Ajouter les sections d'information
            AddSection("OBJECTIF PRINCIPAL",
                "Automatiser la verification des Cut Lists PDF par rapport aux fichiers CSV.\n\n" +
                "[+] Eliminer les 30-40 minutes de verification manuelle par module\n" +
                "[+] Precision validee ~97% sur projets multi-modules reels\n" +
                "[+] Integration native templates Excel XNRGY sans formation requise",
                Color.FromRgb(144, 238, 144)); // Vert clair #90EE90

            AddSeparator();

            AddSection("UTILISATION SIMPLE",
                "1. Ouvrir le PDF Cut List dans n'importe quel lecteur\n" +
                "   Adobe Reader, Acrobat, Foxit, Chrome, Edge...\n\n" +
                "2. Cliquer 'Detecter PDF' pour extraction automatique\n" +
                "   [+] Projet, reference et module extraits\n" +
                "   [+] Chemins fichiers construits intelligemment\n\n" +
                "3. Verifier le code couleur des fichiers\n" +
                "   [+] VERT = Fichier accessible  [-] ROUGE = Fichier manquant\n\n" +
                "4. Cliquer 'Verifier le PDF' pour analyse complete",
                Color.FromRgb(100, 149, 237)); // Bleu Cornflower #6495ED

            AddSeparator();

            AddSection("FONCTIONS ALTERNATIVES",
                "Selection manuelle si detection echoue :\n\n" +
                "'Choisir Projet' - Parcourir les projets disponibles\n" +
                "[+] Liste complete projets/references/modules\n" +
                "[+] Navigation visuelle structure dossiers\n" +
                "[+] Selection rapide projets connus\n\n" +
                "Boutons [...] - Selection manuelle fichiers\n" +
                "[+] CSV : Parcourir si chemin non standard\n" +
                "[+] Excel : Selectionner fichier existant\n" +
                "[+] PDF : Choisir autre PDF que 02-Machines\n\n" +
                "Chemin de base modifiable pour projets speciaux",
                Color.FromRgb(255, 183, 77)); // Orange clair #FFB74D

            AddSeparator();

            AddSection("EXTRACTION PDF AVANCEE",
                "Double strategie pour precision maximale :\n\n" +
                "Strategie 1 - Tableaux structures\n" +
                "[+] Detection colonnes TAG et QTY automatique\n" +
                "[+] Gestion variations mise en page Inventor\n\n" +
                "Strategie 2 - Ballons et cartouches\n" +
                "[+] Recuperation elements manques\n" +
                "[+] Extraction annotations et cartouches\n\n" +
                "Reconnaissance intelligente tags, quantites et materiaux",
                Color.FromRgb(186, 85, 211)); // Violet clair #BA55D3

            AddSeparator();

            AddSection("CODE COULEUR UNIVERSEL",
                "VERT : Correspondance parfaite (85% des cas)\n" +
                "[+] Tag + quantite + materiau corrects\n\n" +
                "JAUNE : Difference quantite ou materiau (11% des cas)\n" +
                "[!] Verification manuelle recommandee\n\n" +
                "BLEU : Quantite zero dans PDF (1% des cas)\n" +
                "[?] Cas speciaux production\n\n" +
                "ROUGE : Tag manquant (3% des cas)\n" +
                "[-] Investigation requise\n\n" +
                "Statistiques temps reel avec compteurs et pourcentages",
                Color.FromRgb(255, 165, 0)); // Orange #FFA500

            AddSeparator();

            AddSection("DETECTION AUTOMATIQUE AVANCEE",
                "Detection multi-niveau du PDF ouvert :\n\n" +
                "[+] Enumeration fenetres Windows API native\n" +
                "[+] Support Adobe Reader/Acrobat COM API\n" +
                "[+] Extraction chemins depuis titres fenetres\n" +
                "[+] Detection fichiers verrouilles (en cours)\n" +
                "[+] Reconstruction chemins structure XNRGY\n\n" +
                "Extraction automatique projet/reference/module\n" +
                "[+] Pattern 4-6 chiffres pour projets\n" +
                "[+] REF01, REF02 ou numeros simples\n" +
                "[+] M01, M02 pour modules",
                Color.FromRgb(138, 43, 226)); // Violet BlueViolet #8A2BE2

            AddSeparator();

            AddSection("INTEGRATION EXCEL",
                "Gestion automatique complete :\n\n" +
                "[+] Copie templates depuis C:\\Vault\\Engineering\\Library\\\n" +
                "[+] Renommage selon structure projet\n" +
                "[+] Mise a jour metadonnees automatique\n" +
                "[+] Fermeture Excel si verrouille\n" +
                "[+] Formatage couleurs metier\n" +
                "[+] Compatibilite workflow Check Lists existant\n\n" +
                "Cellules mises a jour automatiquement :\n" +
                "[+] C1 : Date analyse (YYYY-MM-DD)\n" +
                "[+] C2 : Numero projet\n" +
                "[+] C3 : Reference (numerique)\n" +
                "[+] C4 : Module (numerique)",
                Color.FromRgb(0, 206, 209)); // Turquoise #00CED1

            AddSeparator();

            AddSection("PERFORMANCES VALIDEES",
                "Metriques production XNRGY :\n\n" +
                "[+] Extraction : ~97% sur projets multi-modules reels\n" +
                "[+] Temps : 3-8 minutes vs 30-40 minutes manuellement\n" +
                "[+] Analyse PDF : <5 secondes pour fichiers 2-4 pages\n" +
                "[+] Memoire : <100MB compatible workstations CAD\n" +
                "[+] Detection automatique : >95% projets standards\n" +
                "[+] Adoption : 100% grace integration native\n" +
                "[+] ROI : 85% gain temps par module\n\n" +
                "Compatible : .NET 9, Windows 10/11",
                Color.FromRgb(147, 112, 219)); // Violet MediumPurple #9370DB

            AddSeparator();

            AddSection("DEVELOPPEUR ET SUPPORT",
                "Mohammed Amine Elgalai\n" +
                "Automation Engineer\n" +
                "XNRGY Climate Systems ULC\n\n" +
                "Expertise :\n" +
                "[+] Conception modules energetiques Inventor\n" +
                "[+] Developpement outils automatisation CAD\n" +
                "[+] Optimisation workflows production\n" +
                "[+] Formation equipes techniques\n\n" +
                "Contact :\n" +
                "[+] Email : mohammedamine.elgalai@xnrgy.com\n" +
                "[+] Teams : @Mohammed Amine Elgalai\n\n" +
                "Licence : Proprietaire XNRGY Climate Systems ULC",
                Color.FromRgb(144, 238, 144)); // Vert clair
        }

        private void AddSection(string title, string content, Color titleColor)
        {
            // Titre de section
            var titleBlock = new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(titleColor),
                Margin = new Thickness(0, 0, 0, 8)
            };
            ContentPanel.Children.Add(titleBlock);

            // Contenu de section - THEME SOMBRE
            var contentBlock = new TextBlock
            {
                Text = content,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(224, 224, 224)), // #E0E0E0 - Texte clair
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 15)
            };
            ContentPanel.Children.Add(contentBlock);
        }

        private void AddSeparator()
        {
            var separator = new Border
            {
                Height = 1,
                Background = new SolidColorBrush(Color.FromRgb(74, 127, 191)), // #4A7FBF - BorderColor
                Margin = new Thickness(0, 5, 0, 15)
            };
            ContentPanel.Children.Add(separator);
        }
    }
}
