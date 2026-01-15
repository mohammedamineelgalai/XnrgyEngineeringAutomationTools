// ============================================================================
// DXFVerifierWindow.xaml.cs
// DXF-CSV vs PDF Verifier v1.2 - WPF Migration
// Author: Mohammed Amine Elgalai
// XNRGY Climate Systems ULC
// ============================================================================
// [!] CRITICAL: This is a PRECISE migration from MainForm.vb (VB.NET WinForms)
// [!] DO NOT modify the logic - it was calibrated over 1 month
// ============================================================================

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media;
using Microsoft.Win32;
using XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Services;
using XnrgyEngineeringAutomationTools.Services;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Binding = System.Windows.Data.Binding;

namespace XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Views
{
    /// <summary>
    /// DXF-CSV vs PDF Verifier Window - XNRGY Engineering Automation Tools
    /// Precise migration from VB.NET WinForms MainForm.vb
    /// </summary>
    public partial class DXFVerifierWindow : Window
    {
        #region Win32 API Imports

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const int SW_MAXIMIZE = 3;
        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

        #endregion

        #region Private Fields

        private string _csvFilePath = string.Empty;
        private string _excelFilePath = string.Empty;
        private string _pdfFilePath = string.Empty;
        private string _logFilePath = string.Empty;
        private List<DxfItem> _csvItems = new List<DxfItem>();
        private List<DxfItem> _pdfItems = new List<DxfItem>();
        private ObservableCollection<VerificationResultItem> _results = new ObservableCollection<VerificationResultItem>();

        private const string BasePath = @"C:\Vault\Engineering\Projects";
        private readonly StringBuilder _logBuilder = new StringBuilder();
        
        // Service Vault herite de MainWindow
        private readonly VaultSdkService? _vaultService;
        
        // Chronometre pour la progression style Upload Module
        private DateTime _progressStartTime;
        private System.Windows.Threading.DispatcherTimer? _progressTimer;
        
        // Chronometre Stopwatch pour mesure precise du temps ecoule/estime
        private readonly Stopwatch _verificationStopwatch = new Stopwatch();

        #endregion

        #region Constructor

        /// <summary>
        /// Constructeur par defaut (sans service Vault)
        /// </summary>
        public DXFVerifierWindow() : this(null)
        {
        }

        /// <summary>
        /// Constructeur avec service Vault pour heritage du statut de connexion depuis MainWindow
        /// </summary>
        /// <param name="vaultService">Service Vault connecte (optionnel)</param>
        public DXFVerifierWindow(VaultSdkService? vaultService)
        {
            _vaultService = vaultService;
            
            InitializeComponent();
            
            // Subscribe to static service log events
            PdfAnalyzerService.OnLog += OnServiceLog;
            ExcelManagerService.OnLog += OnServiceLog;
            PdfFormFillerService.OnLog += OnServiceLog;
            
            // Bind results to DataGrid
            ResultDataGrid.ItemsSource = _results;
            
            // Initialize log file
            InitializeLogFile();
        }

        private void OnServiceLog(string category, string message)
        {
            // Utiliser Dispatcher pour thread-safety (appele depuis Task.Run)
            Dispatcher.BeginInvoke(new Action(() =>
            {
                LogMessage(message);
            }));
        }

        #endregion

        #region Window Events

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Charger les listes deroulantes des projets
            LoadProjectDropdowns();
            
            // Mettre a jour le statut Vault herite de MainWindow
            UpdateVaultConnectionStatusFromService();
            
            LogMessage("[+] DXF-CSV vs PDF Verifier v1.0 initialise");
            LogMessage("[i] Pret pour la detection de projet");
            BasePathTextBox.Text = BasePath;
            ShowStatus("Pret - Detectez un PDF ouvert ou selectionnez un projet manuellement", StatusType.Info);
        }

        /// <summary>
        /// Met a jour le statut Vault depuis le service herite de MainWindow
        /// </summary>
        private void UpdateVaultConnectionStatusFromService()
        {
            if (_vaultService != null && _vaultService.IsConnected)
            {
                UpdateVaultConnectionStatus(true, _vaultService.VaultName, _vaultService.UserName);
            }
            else
            {
                UpdateVaultConnectionStatus(false, null, null);
            }
        }

        /// <summary>
        /// Met a jour l'indicateur de connexion Vault dans l'en-tete (style SmartTools/CreateModule)
        /// </summary>
        private void UpdateVaultConnectionStatus(bool isConnected, string? vaultName, string? userName)
        {
            Dispatcher.Invoke(() =>
            {
                if (VaultStatusIndicator != null)
                {
                    VaultStatusIndicator.Fill = new SolidColorBrush(
                        isConnected ? (Color)ColorConverter.ConvertFromString("#107C10") : (Color)ColorConverter.ConvertFromString("#E81123"));
                }
                
                if (RunVaultName != null && RunUserName != null && RunStatus != null)
                {
                    RunVaultName.Text = isConnected ? $" Vault : {vaultName}  /  " : " Vault : --  /  ";
                    RunUserName.Text = isConnected ? $" Utilisateur : {userName}  /  " : " Utilisateur : --  /  ";
                    RunStatus.Text = isConnected ? " Statut : Connecte" : " Statut : Deconnecte";
                }
            });
        }

        /// <summary>
        /// Helper method to set a value in a ComboBox by finding the matching item
        /// </summary>
        private void SetComboBoxValue(System.Windows.Controls.ComboBox comboBox, string value)
        {
            if (string.IsNullOrEmpty(value) || comboBox.ItemsSource == null) return;
            
            foreach (var item in comboBox.ItemsSource)
            {
                if (item?.ToString() == value)
                {
                    comboBox.SelectedItem = item;
                    return;
                }
            }
            // If not found in list, try to add it temporarily
            // For non-editable ComboBox, we can't type a value, so leave unselected
        }

        private void LoadProjectDropdowns()
        {
            try
            {
                var basePath = BasePathTextBox.Text;
                if (!Directory.Exists(basePath)) return;

                // Charger les projets disponibles
                var projectDirs = Directory.GetDirectories(basePath)
                    .Select(d => Path.GetFileName(d))
                    .Where(name => Regex.IsMatch(name, @"^\d+$"))
                    .OrderByDescending(x => x)
                    .ToList();

                ProjectNumberComboBox.ItemsSource = projectDirs;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur chargement projets: {ex.Message}");
            }
        }

        private void ProjectNumberComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var projectNumber = ProjectNumberComboBox.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(projectNumber)) return;

                var projectPath = Path.Combine(BasePathTextBox.Text, projectNumber);
                if (!Directory.Exists(projectPath)) return;

                // Charger les references disponibles pour ce projet
                var refDirs = Directory.GetDirectories(projectPath)
                    .Select(d => Path.GetFileName(d))
                    .Where(name => name.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
                    .Select(name => name.Substring(3)) // Enlever "REF" prefix
                    .OrderByDescending(x => x)
                    .ToList();

                ReferenceComboBox.ItemsSource = refDirs;
                ReferenceComboBox.SelectedIndex = -1;
                ModuleNumberComboBox.ItemsSource = null;
                ModuleNumberComboBox.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur chargement references: {ex.Message}");
            }
        }

        private void ReferenceComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var projectNumber = ProjectNumberComboBox.SelectedItem?.ToString();
                var reference = ReferenceComboBox.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(projectNumber) || string.IsNullOrEmpty(reference)) return;

                var refPath = Path.Combine(BasePathTextBox.Text, projectNumber, $"REF{reference}");
                if (!Directory.Exists(refPath)) return;

                // Charger les modules disponibles pour cette reference
                var moduleDirs = Directory.GetDirectories(refPath)
                    .Select(d => Path.GetFileName(d))
                    .Where(name => name.StartsWith("M", StringComparison.OrdinalIgnoreCase) && name.Length <= 4)
                    .OrderBy(x => x)
                    .ToList();

                ModuleNumberComboBox.ItemsSource = moduleDirs;
                ModuleNumberComboBox.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur chargement modules: {ex.Message}");
            }
        }

        private void ModuleNumberComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var projectNumber = ProjectNumberComboBox.SelectedItem?.ToString();
                var reference = ReferenceComboBox.SelectedItem?.ToString();
                var module = ModuleNumberComboBox.SelectedItem?.ToString();
                
                if (string.IsNullOrEmpty(projectNumber) || string.IsNullOrEmpty(reference) || string.IsNullOrEmpty(module))
                    return;

                // Detecter automatiquement les fichiers pour ce module
                var moduleNum = module.StartsWith("M", StringComparison.OrdinalIgnoreCase) ? module.Substring(1) : module;
                var projectInfo = new ProjectPathInfo
                {
                    ProjectNumber = projectNumber,
                    Reference = reference,
                    ModuleNumber = moduleNum,
                    BasePath = Path.Combine(BasePathTextBox.Text, projectNumber, $"REF{reference}", module)
                };
                
                AutoDetectFiles(projectInfo);
                LogMessage($"[+] Module selectionne: {projectNumber} REF{reference} {module}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur selection module: {ex.Message}");
            }
        }

        #endregion

        #region PDF Detection

        private void DetectPdfButton_Click(object sender, RoutedEventArgs e)
        {
            LogMessage("[>] Detection du PDF ouvert dans Adobe Reader/Acrobat...");
            ShowStatus("Detection du PDF en cours...", StatusType.Info);
            
            try
            {
                var pdfPath = FindOpenPdfDirect();
                
                if (!string.IsNullOrEmpty(pdfPath))
                {
                    _pdfFilePath = pdfPath;
                    PdfPathLabel.Text = pdfPath;
                    PdfPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                    
                    LogMessage($"[+] PDF detecte: {pdfPath}");
                    
                    // Extract project info from path
                    var projectInfo = ExtractProjectInfoFromPath(pdfPath);
                    if (projectInfo != null)
                    {
                        SetComboBoxValue(ProjectNumberComboBox, projectInfo.ProjectNumber);
                        SetComboBoxValue(ReferenceComboBox, projectInfo.Reference);
                        SetComboBoxValue(ModuleNumberComboBox, projectInfo.ModuleNumber);
                        
                        LogMessage($"[+] Projet detecte: {projectInfo.ProjectNumber} REF{projectInfo.Reference} M{projectInfo.ModuleNumber}");
                        
                        // Auto-detect CSV and Excel files
                        AutoDetectFiles(projectInfo);
                    }
                    
                    ShowStatus($"PDF detecte: {Path.GetFileName(pdfPath)}", StatusType.Success);
                }
                else
                {
                    LogMessage("[!] Aucun PDF ouvert detecte dans Adobe Reader/Acrobat");
                    ShowStatus("Aucun PDF ouvert detecte - Selectionnez manuellement", StatusType.Warning);
                    MessageBox.Show("Aucun fichier PDF ouvert detecte dans Adobe Reader ou Acrobat.\n\n" +
                                  "Assurez-vous qu'un fichier PDF 03-CutList est ouvert.", 
                                  "Detection PDF", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[-] Erreur detection PDF: {ex.Message}");
                ShowStatus("Erreur lors de la detection du PDF", StatusType.Error);
            }
        }

        private string FindOpenPdfDirect()
        {
            string pdfPath = string.Empty;
            
            try
            {
                // Try Adobe Acrobat DC first
                var acrobatPath = TryGetAdobePathViaRegistry("AcroExch.App");
                if (!string.IsNullOrEmpty(acrobatPath))
                {
                    LogMessage($"[+] Adobe Acrobat detecte via COM");
                    pdfPath = GetPdfPathFromAdobeCom("AcroExch.App");
                    if (!string.IsNullOrEmpty(pdfPath)) return pdfPath;
                }
                
                // Try via process window title
                var processes = new[] { "AcroRd32", "Acrobat", "AcroRd64" };
                foreach (var procName in processes)
                {
                    var procs = Process.GetProcessesByName(procName);
                    foreach (var proc in procs)
                    {
                        try
                        {
                            var title = proc.MainWindowTitle;
                            if (!string.IsNullOrEmpty(title))
                            {
                                // Extract file path from window title
                                // Adobe typically shows: "filename.pdf - Adobe Acrobat Reader DC"
                                var match = Regex.Match(title, @"^(.+\.pdf)", RegexOptions.IgnoreCase);
                                if (match.Success)
                                {
                                    var fileName = match.Groups[1].Value.Trim();
                                    // Search for this file in common locations
                                    pdfPath = SearchForPdfFile(fileName);
                                    if (!string.IsNullOrEmpty(pdfPath))
                                    {
                                        LogMessage($"[+] PDF trouve via titre fenetre: {pdfPath}");
                                        return pdfPath;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[!] Erreur FindOpenPdfDirect: {ex.Message}");
            }
            
            return pdfPath;
        }

        private string TryGetAdobePathViaRegistry(string progId)
        {
            try
            {
                using (var key = Registry.ClassesRoot.OpenSubKey($@"{progId}\CLSID"))
                {
                    if (key != null)
                    {
                        return key.GetValue(null)?.ToString() ?? string.Empty;
                    }
                }
            }
            catch { }
            return string.Empty;
        }

        private string GetPdfPathFromAdobeCom(string progId)
        {
            try
            {
                var adobeType = Type.GetTypeFromProgID(progId);
                if (adobeType != null)
                {
                    dynamic app = Activator.CreateInstance(adobeType);
                    if (app != null)
                    {
                        try
                        {
                            dynamic doc = app.GetActiveDoc();
                            if (doc != null)
                            {
                                string path = doc.GetFileName();
                                Marshal.ReleaseComObject(doc);
                                Marshal.ReleaseComObject(app);
                                return path;
                            }
                        }
                        catch { }
                        Marshal.ReleaseComObject(app);
                    }
                }
            }
            catch { }
            return string.Empty;
        }

        private string SearchForPdfFile(string fileName)
        {
            // Search in Projects folder
            if (Directory.Exists(BasePath))
            {
                try
                {
                    var files = Directory.GetFiles(BasePath, fileName, SearchOption.AllDirectories);
                    if (files.Length > 0)
                    {
                        return files[0];
                    }
                }
                catch { }
            }
            
            // Check if it's already a full path
            if (File.Exists(fileName))
            {
                return fileName;
            }
            
            return string.Empty;
        }

        #endregion

        #region Project Info Extraction

        private ProjectPathInfo ExtractProjectInfoFromPath(string filePath)
        {
            try
            {
                // Pattern: C:\Vault\Engineering\Projects\12345\REF01\M01\...
                var match = Regex.Match(filePath, @"Projects[\\\/](\d+)[\\\/]REF(\d+)[\\\/]M(\d+)", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return new ProjectPathInfo
                    {
                        ProjectNumber = match.Groups[1].Value,
                        Reference = match.Groups[2].Value,
                        ModuleNumber = match.Groups[3].Value,
                        BasePath = Path.GetDirectoryName(filePath) ?? string.Empty
                    };
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[!] Erreur extraction info projet: {ex.Message}");
            }
            return null;
        }

        private void AutoDetectFiles(ProjectPathInfo projectInfo)
        {
            var projectFolder = $@"{BasePath}\{projectInfo.ProjectNumber}\REF{projectInfo.Reference}\M{projectInfo.ModuleNumber}";
            
            LogMessage($"[>] Recherche des fichiers dans: {projectFolder}");
            
            if (!Directory.Exists(projectFolder))
            {
                LogMessage($"[-] Dossier non trouve: {projectFolder}");
                return;
            }
            
            // Extraire le numero de reference sans le prefixe REF
            var refNumber = projectInfo.Reference;
            
            // === 1. DETECTION CSV ===
            // Chemin Vault 2026: 5_Exportation/Sheet_Metal_Nesting/Punch/[PROJECT]-[REF]-M[XX].csv
            var csvPath = Path.Combine(projectFolder, "5_Exportation", "Sheet_Metal_Nesting", "Punch", 
                $"{projectInfo.ProjectNumber}-{refNumber}-M{projectInfo.ModuleNumber}.csv");
            
            if (File.Exists(csvPath))
            {
                _csvFilePath = csvPath;
                CsvPathLabel.Text = _csvFilePath;
                CsvPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                LogMessage($"[+] CSV detecte: {Path.GetFileName(_csvFilePath)}");
            }
            else
            {
                // Fallback: recherche par pattern dans tout le dossier
                var csvPatterns = new[] { "*-DXF.csv", "*_DXF.csv", "*DXF*.csv", "*.csv" };
                foreach (var pattern in csvPatterns)
                {
                    try
                    {
                        var csvFiles = Directory.GetFiles(projectFolder, pattern, SearchOption.AllDirectories);
                        if (csvFiles.Length > 0)
                        {
                            _csvFilePath = csvFiles[0];
                            CsvPathLabel.Text = _csvFilePath;
                            CsvPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                            LogMessage($"[+] CSV detecte (fallback): {Path.GetFileName(_csvFilePath)}");
                            break;
                        }
                    }
                    catch { /* Ignorer les erreurs d'acces */ }
                }
            }
            
            if (string.IsNullOrEmpty(_csvFilePath))
            {
                CsvPathLabel.Text = "Non detecte";
                CsvPathLabel.Foreground = new SolidColorBrush(Color.FromRgb(255, 99, 71)); // Tomato
                LogMessage($"[-] CSV non trouve");
            }
            
            // === 2. DETECTION EXCEL ===
            // Chemin Vault 2026: 0-Documents/[PROJECT]-[REF]-M[XX]_Decompte de DXF_DXF Count.xlsx
            var excelPath = Path.Combine(projectFolder, "0-Documents", 
                $"{projectInfo.ProjectNumber}-{refNumber}-M{projectInfo.ModuleNumber}_Décompte de DXF_DXF Count.xlsx");
            
            if (File.Exists(excelPath))
            {
                _excelFilePath = excelPath;
                ExcelPathLabel.Text = _excelFilePath;
                ExcelPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                LogMessage($"[+] Excel detecte: {Path.GetFileName(_excelFilePath)}");
            }
            else
            {
                // Fallback: recherche par pattern
                var excelPatterns = new[] { "*DXF Count*.xlsx", "*_DXF*.xlsx", "*-DXF.xlsx" };
                foreach (var pattern in excelPatterns)
                {
                    try
                    {
                        var excelFiles = Directory.GetFiles(projectFolder, pattern, SearchOption.AllDirectories);
                        if (excelFiles.Length > 0)
                        {
                            _excelFilePath = excelFiles[0];
                            ExcelPathLabel.Text = _excelFilePath;
                            ExcelPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                            LogMessage($"[+] Excel detecte (fallback): {Path.GetFileName(_excelFilePath)}");
                            break;
                        }
                    }
                    catch { /* Ignorer les erreurs d'acces */ }
                }
            }
            
            if (string.IsNullOrEmpty(_excelFilePath))
            {
                ExcelPathLabel.Text = "Non detecte";
                ExcelPathLabel.Foreground = new SolidColorBrush(Color.FromRgb(255, 215, 0)); // Gold - Excel peut etre cree
                LogMessage($"[!] Excel non trouve - sera cree automatiquement");
            }
            
            // === 3. DETECTION PDF ===
            // Chemin Vault 2026: 6-Shop Drawing PDF/Production/BatchPrint/02-Machines.pdf
            var pdfPath = Path.Combine(projectFolder, "6-Shop Drawing PDF", "Production", "BatchPrint", "02-Machines.pdf");
            
            if (File.Exists(pdfPath))
            {
                _pdfFilePath = pdfPath;
                PdfPathLabel.Text = _pdfFilePath;
                PdfPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                LogMessage($"[+] PDF detecte: {Path.GetFileName(_pdfFilePath)}");
            }
            else
            {
                // Fallback: recherche par pattern
                var pdfPatterns = new[] { "*Machines*.pdf", "*CutList*.pdf", "*Cut_List*.pdf", "*.pdf" };
                foreach (var pattern in pdfPatterns)
                {
                    try
                    {
                        var pdfFiles = Directory.GetFiles(projectFolder, pattern, SearchOption.AllDirectories);
                        if (pdfFiles.Length > 0)
                        {
                            _pdfFilePath = pdfFiles[0];
                            PdfPathLabel.Text = _pdfFilePath;
                            PdfPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                            LogMessage($"[+] PDF detecte (fallback): {Path.GetFileName(_pdfFilePath)}");
                            break;
                        }
                    }
                    catch { /* Ignorer les erreurs d'acces */ }
                }
            }
            
            if (string.IsNullOrEmpty(_pdfFilePath))
            {
                PdfPathLabel.Text = "Non detecte";
                PdfPathLabel.Foreground = new SolidColorBrush(Color.FromRgb(255, 99, 71)); // Tomato
                LogMessage($"[-] PDF non trouve");
            }
            
            // Afficher le statut final
            var filesFound = new List<string>();
            if (!string.IsNullOrEmpty(_csvFilePath)) filesFound.Add("CSV");
            if (!string.IsNullOrEmpty(_excelFilePath)) filesFound.Add("Excel");
            if (!string.IsNullOrEmpty(_pdfFilePath)) filesFound.Add("PDF");
            
            if (filesFound.Count == 3)
            {
                ShowStatus($"[+] Tous les fichiers detectes - Pret pour verification", StatusType.Success);
            }
            else if (filesFound.Count > 0)
            {
                ShowStatus($"[!] Fichiers detectes: {string.Join(", ", filesFound)}", StatusType.Warning);
            }
            else
            {
                ShowStatus($"[-] Aucun fichier detecte dans le module", StatusType.Error);
            }
        }

        #endregion

        #region Verification Logic

        private async void VerifyButton_Click(object sender, RoutedEventArgs e)
        {
            // Validate inputs
            if (string.IsNullOrEmpty(_csvFilePath) && string.IsNullOrEmpty(_excelFilePath))
            {
                MessageBox.Show("Veuillez d'abord detecter ou selectionner un fichier CSV/Excel.", 
                              "Fichier manquant", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            if (string.IsNullOrEmpty(_pdfFilePath))
            {
                MessageBox.Show("Veuillez d'abord detecter ou selectionner un fichier PDF.", 
                              "Fichier manquant", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            VerifyButton.IsEnabled = false;
            ResetProgress();
            StartProgressTimer();
            _results.Clear();
            
            try
            {
                LogMessage("=".PadRight(60, '='));
                LogMessage("[>] DEBUT DE LA VERIFICATION");
                LogMessage("=".PadRight(60, '='));
                
                // Demarrer le chronometre pour le temps ecoule/estime
                StartTimer();
                
                // Step 1: Read CSV/Excel data
                UpdateTimer(10, 5, 1, "Lecture CSV/Excel");
                
                _csvItems.Clear();
                
                if (!string.IsNullOrEmpty(_csvFilePath) && File.Exists(_csvFilePath))
                {
                    LogMessage($"[>] Lecture CSV: {Path.GetFileName(_csvFilePath)}");
                    // Use static method from ExcelManagerService
                    _csvItems = ExcelManagerService.ReadCsvFile(_csvFilePath);
                    LogMessage($"[+] {_csvItems.Count} tags lus depuis CSV");
                }
                
                UpdateTimer(30, 5, 2, "Analyse PDF");
                
                // Step 2: Analyze PDF avec strategie CSV->PDF
                LogMessage($"[>] Analyse PDF: {Path.GetFileName(_pdfFilePath)}");
                
                _pdfItems.Clear();
                
                // STRATEGIE CSV->PDF: Construire la reference CSV pour corriger les quantites des tags echus
                Dictionary<string, PdfAnalyzerService.CsvRow>? csvReference = null;
                if (_csvItems.Count > 0)
                {
                    csvReference = new Dictionary<string, PdfAnalyzerService.CsvRow>(StringComparer.OrdinalIgnoreCase);
                    foreach (var csvItem in _csvItems)
                    {
                        var normalizedTag = NormalizeTag(csvItem.Tag);
                        if (!csvReference.ContainsKey(normalizedTag))
                        {
                            csvReference[normalizedTag] = new PdfAnalyzerService.CsvRow(csvItem.Tag, csvItem.Quantity, csvItem.Material);
                        }
                    }
                    LogMessage($"[i] Reference CSV preparee: {csvReference.Count} tags");
                }
                
                // Utiliser ExtractTablesFromPdfV2 avec algorithme de proximite X/Y valide a 100%
                // Note: Les logs sont automatiquement envoyes au journal via OnServiceLog (connecte dans constructeur)
                var pdfData = await Task.Run(() => PdfAnalyzerService.ExtractTablesFromPdfV2(_pdfFilePath, csvReference));
                
                // Convert PDF Dictionary<tag, qty> to List<DxfItem>
                foreach (var kvp in pdfData)
                {
                    _pdfItems.Add(new DxfItem
                    {
                        Tag = kvp.Key,
                        Quantity = kvp.Value,
                        Material = string.Empty
                    });
                }
                
                LogMessage($"[+] {_pdfItems.Count} tags extraits du PDF (strategie CSV->PDF)");
                LogMessage($"[i] Pages analysees: {PdfAnalyzerService.LastAnalyzedPageCount}");
                
                UpdateTimer(50, 5, 3, "Comparaison des resultats");
                
                // Step 3: Compare results
                CompareAndDisplayResults();
                
                UpdateTimer(70, 5, 4, "Statistiques");
                
                // Step 4: Update statistics
                UpdateStatistics();
                
                // Step 5: Write results to Excel report
                if (!string.IsNullOrEmpty(_excelFilePath))
                {
                    UpdateTimer(85, 5, 5, "Ecriture Excel");
                    LogMessage($"[>] Ecriture du rapport Excel: {Path.GetFileName(_excelFilePath)}");
                    
                    try
                    {
                        // Convert results to DxfItem list for ExcelManagerService
                        var dxfItemsForExcel = _csvItems.Select(csv =>
                        {
                            var result = _results.FirstOrDefault(r => r.Tag == csv.Tag);
                            return new DxfItem
                            {
                                Tag = csv.Tag,
                                Quantity = csv.Quantity,
                                Material = csv.Material,
                                FoundInPdf = result?.PdfQuantity > 0 || (result?.Status.Contains("[+]") ?? false),
                                PdfQuantity = result?.PdfQuantity ?? 0
                            };
                        }).ToList();
                        
                        ExcelManagerService.WriteToExcel(_excelFilePath, dxfItemsForExcel);
                        LogMessage($"[+] Rapport Excel sauvegarde: {dxfItemsForExcel.Count} elements");
                    }
                    catch (Exception exExcel)
                    {
                        LogMessage($"[!] Erreur ecriture Excel: {exExcel.Message}");
                    }
                }
                
                // Arret du timer et affichage final
                _verificationStopwatch.Stop();
                UpdateTimer(100, 5, 5, "Termine");
                
                // Step 6: Determine final status
                var mismatchCount = _results.Count(r => r.Status.Contains("[-]") || r.Status.Contains("[!]") || r.Status.Contains("[?]"));
                if (mismatchCount == 0)
                {
                    ShowStatus($"SUCCES: {_results.Count} tags verifies", StatusType.Success);
                    LogMessage($"[+] VERIFICATION REUSSIE: {_results.Count} tags verifies");
                }
                else
                {
                    ShowStatus($"{mismatchCount} differences sur {_results.Count} tags", StatusType.Warning);
                    LogMessage($"[!] VERIFICATION AVEC DIFFERENCES: {mismatchCount}/{_results.Count} tags avec problemes");
                }
                
                LogMessage("=".PadRight(60, '='));
                LogMessage("[+] FIN DE LA VERIFICATION");
                LogMessage("=".PadRight(60, '='));
                
                // Step 7: Open Excel report at the end (only Excel, not PDF)
                if (!string.IsNullOrEmpty(_excelFilePath) && File.Exists(_excelFilePath))
                {
                    LogMessage($"[>] Ouverture du rapport Excel...");
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = _excelFilePath,
                            UseShellExecute = true
                        });
                        LogMessage($"[+] Excel ouvert: {Path.GetFileName(_excelFilePath)}");
                    }
                    catch (Exception exOpen)
                    {
                        LogMessage($"[!] Erreur ouverture Excel: {exOpen.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[-] ERREUR VERIFICATION: {ex.Message}");
                ShowStatus($"Erreur: {ex.Message}", StatusType.Error);
                MessageBox.Show($"Erreur lors de la verification:\n{ex.Message}", 
                              "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                VerifyButton.IsEnabled = true;
            }
        }

        private void CompareAndDisplayResults()
        {
            _results.Clear();
            
            // Create lookup dictionary for PDF items by tag (pour detecter doublons)
            var pdfLookup = new Dictionary<string, DxfItem>(StringComparer.OrdinalIgnoreCase);
            var duplicateTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // Tags en double dans PDF
            
            foreach (var pdfItem in _pdfItems)
            {
                var normalizedTag = NormalizeTag(pdfItem.Tag);
                if (!pdfLookup.ContainsKey(normalizedTag))
                {
                    pdfLookup[normalizedTag] = pdfItem;
                }
                else
                {
                    // Tag en double detecte - marquer comme doublon
                    duplicateTags.Add(normalizedTag);
                    // Accumulate quantities for duplicate tags
                    pdfLookup[normalizedTag].Quantity += pdfItem.Quantity;
                }
            }
            
            // Log des doublons detectes
            if (duplicateTags.Count > 0)
            {
                LogMessage($"[!] {duplicateTags.Count} tags en double detectes dans le PDF");
            }
            
            // Compare each CSV item with PDF - Couleurs compatibles theme sombre (plus claires)
            foreach (var csvItem in _csvItems)
            {
                var normalizedTag = NormalizeTag(csvItem.Tag);
                var result = new VerificationResultItem
                {
                    Tag = csvItem.Tag,
                    Quantity = csvItem.Quantity,
                    Material = csvItem.Material
                };
                
                if (pdfLookup.TryGetValue(normalizedTag, out var pdfItem))
                {
                    result.PdfQuantity = pdfItem.Quantity;
                    
                    // Verifier si c'est un doublon
                    bool isDuplicate = duplicateTags.Contains(normalizedTag);
                    
                    if (isDuplicate)
                    {
                        // ROUGE VIF - Doublon detecte dans PDF (cumul affiche)
                        result.Status = $"❌ DOUBLON PDF: qty cumulee={pdfItem.Quantity} (CSV={csvItem.Quantity})";
                        result.StatusColor = new SolidColorBrush(Color.FromRgb(255, 80, 80)); // Rouge vif
                    }
                    else if (csvItem.Quantity == pdfItem.Quantity)
                    {
                        // VERT CLAIR - OK (compatible theme sombre)
                        result.Status = "✅ Tag trouvé, quantité OK";
                        result.StatusColor = new SolidColorBrush(Color.FromRgb(144, 238, 144)); // LightGreen
                    }
                    else if (pdfItem.Quantity == 0)
                    {
                        // BLEU CLAIR - Quantite 0
                        result.Status = "🔍 Tag trouvé mais quantité = 0 dans PDF";
                        result.StatusColor = new SolidColorBrush(Color.FromRgb(173, 216, 230)); // LightBlue
                    }
                    else
                    {
                        // ROUGE - Quantite differente (demande utilisateur: rouge au lieu d'orange)
                        result.Status = $"❌ Qty DIFFERENTE: CSV={csvItem.Quantity} vs PDF={pdfItem.Quantity}";
                        result.StatusColor = new SolidColorBrush(Color.FromRgb(255, 100, 100)); // Rouge clair
                    }
                }
                else
                {
                    // ROUGE FONCE - Manquant dans PDF
                    result.PdfQuantity = 0;
                    result.Status = "❌ Tag non trouvé dans le PDF";
                    result.StatusColor = new SolidColorBrush(Color.FromRgb(255, 150, 150)); // Rouge clair
                }
                
                _results.Add(result);
            }
            
            // Check for items in PDF not in CSV
            foreach (var pdfItem in _pdfItems)
            {
                var normalizedTag = NormalizeTag(pdfItem.Tag);
                var csvMatch = _csvItems.Any(c => NormalizeTag(c.Tag).Equals(normalizedTag, StringComparison.OrdinalIgnoreCase));
                
                if (!csvMatch)
                {
                    _results.Add(new VerificationResultItem
                    {
                        Tag = pdfItem.Tag,
                        Quantity = 0,
                        Material = string.Empty,
                        PdfQuantity = pdfItem.Quantity,
                        Status = "ℹ️ SUPPLÉMENT dans PDF (pas dans CSV)",
                        StatusColor = new SolidColorBrush(Color.FromRgb(135, 206, 250)) // LightSkyBlue
                    });
                }
            }
            
            // Compter les erreurs
            int erreurCount = _results.Count(r => r.Status.StartsWith("❌"));
            int okCount = _results.Count(r => r.Status.StartsWith("✅"));
            
            LogMessage($"[+] Comparaison terminee: {_results.Count} lignes ({okCount} OK, {erreurCount} erreurs)");
        }

        private string NormalizeTag(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return string.Empty;
            
            // Remove spaces and normalize separators
            return Regex.Replace(tag.Trim().ToUpperInvariant(), @"[-_\s]+", "-");
        }

        private void UpdateStatistics()
        {
            // CSV Statistics
            var totalCsvQty = _csvItems.Sum(i => i.Quantity);
            var totalCsvTags = _csvItems.Count;
            TotalCsvQtyLabel.Text = totalCsvQty.ToString();
            TotalCsvTagsLabel.Text = totalCsvTags.ToString();
            
            // PDF Statistics
            var totalPdfQty = _pdfItems.Sum(i => i.Quantity);
            var totalPdfTags = _pdfItems.Count;
            TotalPdfQtyLabel.Text = totalPdfQty.ToString();
            TotalPdfTagsLabel.Text = totalPdfTags.ToString();
            
            // Color based on match - Update TextBlock foreground color
            if (totalCsvQty == totalPdfQty)
            {
                TotalPdfQtyLabel.Foreground = new SolidColorBrush(Color.FromRgb(144, 238, 144)); // LightGreen
            }
            else
            {
                TotalPdfQtyLabel.Foreground = new SolidColorBrush(Color.FromRgb(255, 99, 71)); // Tomato
            }
            
            if (totalCsvTags == totalPdfTags)
            {
                TotalPdfTagsLabel.Foreground = new SolidColorBrush(Color.FromRgb(144, 238, 144)); // LightGreen
            }
            else
            {
                TotalPdfTagsLabel.Foreground = new SolidColorBrush(Color.FromRgb(255, 165, 0)); // Orange
            }
            
            // PDF Pages - Use static property from PdfAnalyzerService
            try
            {
                var pageCount = PdfAnalyzerService.LastAnalyzedPageCount;
                TotalPdfPagesLabel.Text = pageCount.ToString();
            }
            catch
            {
                TotalPdfPagesLabel.Text = "?";
            }
            
            LogMessage($"[i] Statistiques: CSV={totalCsvQty} pieces/{totalCsvTags} tags | PDF={totalPdfQty} pieces/{totalPdfTags} tags");
        }

        private int ParseInt(string value)
        {
            if (int.TryParse(value, out var result))
                return result;
            return 1;
        }

        #endregion

        #region Browse Buttons

        private void BrowsePathButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Selectionner le chemin de base des projets";
                dialog.SelectedPath = BasePath;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BasePathTextBox.Text = dialog.SelectedPath;
                    LogMessage($"[i] Chemin de base modifie: {dialog.SelectedPath}");
                }
            }
        }

        private void SelectProjectButton_Click(object sender, RoutedEventArgs e)
        {
            LogMessage("[>] Chargement de la liste des projets...");
            ShowStatus("Chargement des projets...", StatusType.Info);
            
            try
            {
                var basePath = BasePathTextBox.Text;
                if (!Directory.Exists(basePath))
                {
                    MessageBox.Show($"Le chemin de base n'existe pas:\n{basePath}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // Load available projects
                var projects = LoadAvailableProjects(basePath);
                
                if (projects.Count == 0)
                {
                    LogMessage("[!] Aucun projet trouve dans le chemin de base");
                    ShowStatus("Aucun projet trouve", StatusType.Warning);
                    MessageBox.Show("Aucun projet trouve dans le chemin de base.\n\n" +
                                  "Structure attendue: Projects\\[XXXXX]\\REF[XX]\\M[XX]",
                                  "Aucun projet", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                LogMessage($"[+] {projects.Count} modules trouves dans {projects.Select(p => p.ProjectNumber).Distinct().Count()} projets");
                
                // Utiliser la nouvelle fenetre de selection moderne
                var dialog = new ProjectSelectorWindow(projects)
                {
                    Owner = this
                };
                
                if (dialog.ShowDialog() == true && dialog.SelectedProject != null)
                {
                    var selectedProject = dialog.SelectedProject;
                    
                    // Extract reference number (remove REF prefix)
                    var refNumber = selectedProject.Reference;
                    if (refNumber.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
                        refNumber = refNumber.Substring(3);
                    
                    SetComboBoxValue(ProjectNumberComboBox, selectedProject.ProjectNumber);
                    SetComboBoxValue(ReferenceComboBox, refNumber);
                    SetComboBoxValue(ModuleNumberComboBox, selectedProject.ModuleNumber);
                    
                    LogMessage($"[+] Projet selectionne: {selectedProject.ProjectNumber} REF{refNumber} {selectedProject.ModuleNumber}");
                    
                    // Auto-detect files
                    var projectInfo = new ProjectPathInfo
                    {
                        ProjectNumber = selectedProject.ProjectNumber,
                        Reference = refNumber,
                        ModuleNumber = selectedProject.ModuleNumber.StartsWith("M") ? selectedProject.ModuleNumber.Substring(1) : selectedProject.ModuleNumber,
                        BasePath = selectedProject.ProjectPath
                    };
                    AutoDetectFiles(projectInfo);
                    
                    ShowStatus($"Projet selectionne: {selectedProject.ProjectNumber}", StatusType.Success);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[-] Erreur lors du chargement des projets: {ex.Message}");
                ShowStatus("Erreur lors du chargement des projets", StatusType.Error);
                MessageBox.Show($"Erreur: {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Load all available projects from the base path
        /// Structure: Projects/[XXXXX]/REF[XX]/M[XX]
        /// </summary>
        private List<ProjectPathInfo> LoadAvailableProjects(string basePath)
        {
            var projects = new List<ProjectPathInfo>();
            
            if (!Directory.Exists(basePath))
                return projects;
            
            try
            {
                // Get all project directories (4-6 digits)
                foreach (var projectDir in Directory.GetDirectories(basePath))
                {
                    var projectNumber = Path.GetFileName(projectDir);
                    if (!Regex.IsMatch(projectNumber, @"^\d{4,6}$"))
                        continue;
                    
                    try
                    {
                        // Get all reference directories (REF01, REF02, etc.)
                        foreach (var refDir in Directory.GetDirectories(projectDir))
                        {
                            var refName = Path.GetFileName(refDir);
                            if (!Regex.IsMatch(refName, @"^REF\d{1,2}$", RegexOptions.IgnoreCase))
                                continue;
                            
                            // Get all module directories (M01, M02, etc.)
                            foreach (var moduleDir in Directory.GetDirectories(refDir))
                            {
                                var moduleName = Path.GetFileName(moduleDir);
                                if (!Regex.IsMatch(moduleName, @"^M\d{1,2}$", RegexOptions.IgnoreCase))
                                    continue;
                                
                                projects.Add(new ProjectPathInfo
                                {
                                    ProjectNumber = projectNumber,
                                    Reference = refName,
                                    ModuleNumber = moduleName,
                                    ProjectPath = moduleDir,
                                    BasePath = moduleDir
                                });
                            }
                        }
                    }
                    catch
                    {
                        // Skip directories we can't access
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[!] Erreur scan projets: {ex.Message}");
            }
            
            return projects;
        }

        private void CsvBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Selectionner le fichier CSV/DXF",
                Filter = "Fichiers CSV|*.csv|Tous les fichiers|*.*",
                InitialDirectory = BasePath
            };
            
            if (dialog.ShowDialog() == true)
            {
                _csvFilePath = dialog.FileName;
                CsvPathLabel.Text = _csvFilePath;
                CsvPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                LogMessage($"[+] CSV selectionne: {dialog.FileName}");
            }
        }

        private void ExcelBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Selectionner le fichier Excel",
                Filter = "Fichiers Excel|*.xlsx;*.xls|Tous les fichiers|*.*",
                InitialDirectory = BasePath
            };
            
            if (dialog.ShowDialog() == true)
            {
                _excelFilePath = dialog.FileName;
                ExcelPathLabel.Text = _excelFilePath;
                ExcelPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                LogMessage($"[+] Excel selectionne: {dialog.FileName}");
            }
        }

        private void PdfBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Selectionner le fichier PDF",
                Filter = "Fichiers PDF|*.pdf|Tous les fichiers|*.*",
                InitialDirectory = BasePath
            };
            
            if (dialog.ShowDialog() == true)
            {
                _pdfFilePath = dialog.FileName;
                PdfPathLabel.Text = _pdfFilePath;
                PdfPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                LogMessage($"[+] PDF selectionne: {dialog.FileName}");
                
                // Try to extract project info
                var projectInfo = ExtractProjectInfoFromPath(_pdfFilePath);
                if (projectInfo != null)
                {
                    SetComboBoxValue(ProjectNumberComboBox, projectInfo.ProjectNumber);
                    SetComboBoxValue(ReferenceComboBox, projectInfo.Reference);
                    SetComboBoxValue(ModuleNumberComboBox, projectInfo.ModuleNumber);
                    AutoDetectFiles(projectInfo);
                }
            }
        }

        #endregion

        #region Info Button

        private void InfoButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var infoWindow = new DXFVerifierInfoWindow();
                infoWindow.Owner = this;
                infoWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                LogMessage($"[-] Erreur ouverture Info: {ex.Message}");
                MessageBox.Show($"Erreur: {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Logging

        private void InitializeLogFile()
        {
            try
            {
                var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                
                _logFilePath = Path.Combine(logDir, $"DXFVerifier_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                LogMessage($"[+] Log file: {_logFilePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing log: {ex.Message}");
            }
        }

        private void LogMessage(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var logLine = $"[{timestamp}] {message}";
            
            // Add to log builder
            _logBuilder.AppendLine(logLine);
            
            // Determiner le niveau de log et la couleur
            JournalColorService.LogLevel level = JournalColorService.LogLevel.INFO;
            if (message.StartsWith("[+]") || message.Contains("succes") || message.Contains("trouve") || message.Contains("OK"))
                level = JournalColorService.LogLevel.SUCCESS;
            else if (message.StartsWith("[-]") || message.Contains("erreur") || message.Contains("Error") || message.Contains("echoue"))
                level = JournalColorService.LogLevel.ERROR;
            else if (message.StartsWith("[!]") || message.Contains("attention") || message.Contains("Warning") || message.Contains("manquant"))
                level = JournalColorService.LogLevel.WARNING;
            else if (message.StartsWith("[~]") || message.StartsWith("[?]"))
                level = JournalColorService.LogLevel.DEBUG;
            
            // Update UI (must be on UI thread)
            Dispatcher.Invoke(() =>
            {
                // Creer le paragraph avec couleur
                var paragraph = new Paragraph();
                paragraph.Margin = new Thickness(0, 2, 0, 2);
                
                // Timestamp en gris
                var timestampRun = new Run($"[{timestamp}] ")
                {
                    Foreground = JournalColorService.TimestampBrush
                };
                paragraph.Inlines.Add(timestampRun);
                
                // Message avec couleur selon le niveau
                var messageRun = new Run(message)
                {
                    Foreground = JournalColorService.GetBrushForLevel(level)
                };
                paragraph.Inlines.Add(messageRun);
                
                LogTextBox.Document.Blocks.Add(paragraph);
                LogTextBox.ScrollToEnd();
            });
            
            // Write to file
            try
            {
                if (!string.IsNullOrEmpty(_logFilePath))
                {
                    File.AppendAllText(_logFilePath, logLine + Environment.NewLine);
                }
            }
            catch { }
        }

        private void ViewLogButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_logFilePath) && File.Exists(_logFilePath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _logFilePath,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Le fichier log n'existe pas encore.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ClearLogButton_Click(object sender, RoutedEventArgs e)
        {
            // Clear RichTextBox
            LogTextBox.Document.Blocks.Clear();
            _logBuilder.Clear();
            LogMessage("[i] Journal efface");
        }

        #endregion

        #region Status Display

        private enum StatusType
        {
            Info,
            Success,
            Warning,
            Error
        }

        /// <summary>
        /// Demarre le chronometre de progression
        /// </summary>
        private void StartProgressTimer()
        {
            _progressStartTime = DateTime.Now;
            
            if (_progressTimer == null)
            {
                _progressTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(500)
                };
                _progressTimer.Tick += (s, e) => UpdateElapsedTime();
            }
            _progressTimer.Start();
        }

        /// <summary>
        /// Arrete le chronometre de progression
        /// </summary>
        private void StopProgressTimer()
        {
            _progressTimer?.Stop();
        }

        /// <summary>
        /// Demarre le chronometre pour mesurer le temps ecoule
        /// </summary>
        private void StartTimer()
        {
            _verificationStopwatch.Restart();
            TxtProgressTimeElapsed.Text = "00:00";
            TxtProgressTimeEstimated.Text = "--:--";
            TxtProgressPercent.Text = "0%";
            TxtProgressStatus.Text = "Demarrage...";
            TxtProgressStep.Text = "";
            ProgressFillBorder.Width = 0;
        }

        /// <summary>
        /// Met a jour la barre de progression avec temps ecoule/estime
        /// </summary>
        private void UpdateTimer(int percentage, int totalSteps, int currentStep, string stepName)
        {
            Dispatcher.Invoke(() =>
            {
                // Calculer la largeur de remplissage
                var parentGrid = ProgressFillBorder.Parent as Grid;
                var totalWidth = parentGrid?.ActualWidth ?? 600;
                if (totalWidth <= 4) totalWidth = 600;
                
                var fillWidth = ((totalWidth - 4) * percentage) / 100.0;
                ProgressFillBorder.Width = Math.Max(0, fillWidth);
                
                // Temps ecoule
                var elapsed = _verificationStopwatch.Elapsed;
                TxtProgressTimeElapsed.Text = FormatTime(elapsed);
                
                // Temps estime (basé sur la progression)
                if (percentage > 0 && percentage < 100)
                {
                    var estimatedTotal = TimeSpan.FromTicks((long)(elapsed.Ticks / (percentage / 100.0)));
                    var remaining = estimatedTotal - elapsed;
                    TxtProgressTimeEstimated.Text = FormatTime(remaining);
                }
                else if (percentage >= 100)
                {
                    TxtProgressTimeEstimated.Text = "00:00";
                }
                
                // Pourcentage
                TxtProgressPercent.Text = $"{percentage}%";
                
                // Etape courante
                TxtProgressStep.Text = $"Etape {currentStep}/{totalSteps}: {stepName}";
            });
        }

        /// <summary>
        /// Formate un TimeSpan en mm:ss
        /// </summary>
        private string FormatTime(TimeSpan time)
        {
            return time.TotalMinutes >= 1 
                ? $"{(int)time.TotalMinutes:D2}:{time.Seconds:D2}" 
                : $"00:{(int)time.TotalSeconds:D2}";
        }

        /// <summary>
        /// Met a jour le temps ecoule affiche
        /// </summary>
        private void UpdateElapsedTime()
        {
            var elapsed = DateTime.Now - _progressStartTime;
            TxtProgressTimeElapsed.Text = elapsed.ToString(@"mm\:ss");
        }

        /// <summary>
        /// Reinitialise la barre de progression
        /// </summary>
        private void ResetProgress()
        {
            Dispatcher.Invoke(() =>
            {
                ProgressFillBorder.Width = 0;
                TxtProgressStatus.Text = "Pret";
                TxtProgressStep.Text = "";
                TxtProgressPercent.Text = "0%";
                TxtProgressTimeElapsed.Text = "00:00";
                TxtProgressTimeEstimated.Text = "00:00";
            });
        }

        private void ShowStatus(string message, StatusType type)
        {
            Dispatcher.Invoke(() =>
            {
                // Couleurs selon le type
                var color = type switch
                {
                    StatusType.Success => Colors.LightGreen,
                    StatusType.Warning => Color.FromRgb(255, 193, 7), // #FFC107 jaune
                    StatusType.Error => Color.FromRgb(220, 53, 69),   // #DC3545 rouge
                    _ => Colors.White
                };
                
                TxtProgressStatus.Foreground = new SolidColorBrush(color);
                TxtProgressStatus.Text = message;
            });
        }

        #endregion
    }

    #region Support Classes

    /// <summary>
    /// Project path information extracted from file path
    /// </summary>
    public class ProjectPathInfo
    {
        public string ProjectNumber { get; set; } = string.Empty;
        public string Reference { get; set; } = string.Empty;
        public string ModuleNumber { get; set; } = string.Empty;
        public string BasePath { get; set; } = string.Empty;
        public string ProjectPath { get; set; } = string.Empty;
    }

    /// <summary>
    /// Verification result item for DataGrid display
    /// </summary>
    public class VerificationResultItem : INotifyPropertyChanged
    {
        private string _tag = string.Empty;
        private int _quantity;
        private string _material = string.Empty;
        private int _pdfQuantity;
        private string _status = string.Empty;
        private Brush _statusColor = Brushes.White;

        public string Tag
        {
            get => _tag;
            set { _tag = value; OnPropertyChanged(nameof(Tag)); }
        }

        public int Quantity
        {
            get => _quantity;
            set { _quantity = value; OnPropertyChanged(nameof(Quantity)); }
        }

        public string Material
        {
            get => _material;
            set { _material = value; OnPropertyChanged(nameof(Material)); }
        }

        public int PdfQuantity
        {
            get => _pdfQuantity;
            set { _pdfQuantity = value; OnPropertyChanged(nameof(PdfQuantity)); }
        }

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(nameof(Status)); }
        }

        public Brush StatusColor
        {
            get => _statusColor;
            set { _statusColor = value; OnPropertyChanged(nameof(StatusColor)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    #endregion
}
