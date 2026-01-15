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
using System.Windows.Forms;
using System.Windows.Media;
using Microsoft.Win32;
using XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Services;
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

        #endregion

        #region Constructor

        public DXFVerifierWindow()
        {
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
            LogMessage($"[{category}] {message}");
        }

        #endregion

        #region Window Events

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Charger le logo XNRGY
            LoadLogo();
            
            LogMessage("[+] DXF-CSV vs PDF Verifier v1.0 initialise");
            LogMessage("[i] Pret pour la detection de projet");
            BasePathTextBox.Text = BasePath;
            ShowStatus("Pret - Detectez un PDF ouvert ou selectionnez un projet manuellement", StatusType.Info);
        }

        private void LoadLogo()
        {
            try
            {
                string logoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "xnrgy_logo.png");
                if (System.IO.File.Exists(logoPath))
                {
                    LogoImage.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri(logoPath, UriKind.Absolute));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur chargement logo: {ex.Message}");
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
                        ProjectNumberTextBox.Text = projectInfo.ProjectNumber;
                        ReferenceTextBox.Text = projectInfo.Reference;
                        ModuleNumberTextBox.Text = projectInfo.ModuleNumber;
                        
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
            
            // Detect CSV/DXF file
            var csvPatterns = new[] { "*-DXF.csv", "*_DXF.csv", "*DXF*.csv" };
            foreach (var pattern in csvPatterns)
            {
                var csvFiles = Directory.GetFiles(projectFolder, pattern, SearchOption.AllDirectories);
                if (csvFiles.Length > 0)
                {
                    _csvFilePath = csvFiles[0];
                    CsvPathLabel.Text = _csvFilePath;
                    CsvPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                    LogMessage($"[+] CSV detecte: {Path.GetFileName(_csvFilePath)}");
                    break;
                }
            }
            
            // Detect Excel file
            var excelPatterns = new[] { "*-DXF.xlsx", "*_DXF.xlsx", "*DXF*.xlsx" };
            foreach (var pattern in excelPatterns)
            {
                var excelFiles = Directory.GetFiles(projectFolder, pattern, SearchOption.AllDirectories);
                if (excelFiles.Length > 0)
                {
                    _excelFilePath = excelFiles[0];
                    ExcelPathLabel.Text = _excelFilePath;
                    ExcelPathLabel.Foreground = new SolidColorBrush(Colors.LightGreen);
                    LogMessage($"[+] Excel detecte: {Path.GetFileName(_excelFilePath)}");
                    break;
                }
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
            ProgressBar.Value = 0;
            _results.Clear();
            
            try
            {
                ShowStatus("Verification en cours...", StatusType.Info);
                LogMessage("=".PadRight(60, '='));
                LogMessage("[>] DEBUT DE LA VERIFICATION");
                LogMessage("=".PadRight(60, '='));
                
                // Step 1: Read CSV/Excel data
                ProgressBar.Value = 10;
                ShowStatus("Lecture des donnees CSV/Excel...", StatusType.Info);
                
                _csvItems.Clear();
                
                if (!string.IsNullOrEmpty(_csvFilePath) && File.Exists(_csvFilePath))
                {
                    LogMessage($"[>] Lecture CSV: {Path.GetFileName(_csvFilePath)}");
                    // Use static method from ExcelManagerService
                    _csvItems = ExcelManagerService.ReadCsvFile(_csvFilePath);
                    LogMessage($"[+] {_csvItems.Count} tags lus depuis CSV");
                }
                
                ProgressBar.Value = 30;
                
                // Step 2: Analyze PDF
                ShowStatus("Analyse du PDF...", StatusType.Info);
                LogMessage($"[>] Analyse PDF: {Path.GetFileName(_pdfFilePath)}");
                
                _pdfItems.Clear();
                // Use static method from PdfAnalyzerService - returns Dictionary<string, int>
                var pdfData = await Task.Run(() => PdfAnalyzerService.ExtractTablesFromPdf(_pdfFilePath));
                
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
                
                LogMessage($"[+] {_pdfItems.Count} tags extraits du PDF");
                
                ProgressBar.Value = 60;
                
                // Step 3: Compare results
                ShowStatus("Comparaison des resultats...", StatusType.Info);
                CompareAndDisplayResults();
                
                ProgressBar.Value = 90;
                
                // Step 4: Update statistics
                UpdateStatistics();
                
                ProgressBar.Value = 100;
                
                // Step 5: Determine final status
                var mismatchCount = _results.Count(r => r.Status.Contains("ERREUR") || r.Status.Contains("MANQUANT"));
                if (mismatchCount == 0)
                {
                    ShowStatus($"Verification terminee - SUCCES: Tous les {_results.Count} tags correspondent!", StatusType.Success);
                    LogMessage($"[+] VERIFICATION REUSSIE: {_results.Count} tags verifies");
                }
                else
                {
                    ShowStatus($"Verification terminee - {mismatchCount} differences trouvees sur {_results.Count} tags", StatusType.Warning);
                    LogMessage($"[!] VERIFICATION AVEC DIFFERENCES: {mismatchCount}/{_results.Count} tags avec problemes");
                }
                
                LogMessage("=".PadRight(60, '='));
                LogMessage("[+] FIN DE LA VERIFICATION");
                LogMessage("=".PadRight(60, '='));
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
            
            // Create lookup dictionary for PDF items by tag
            var pdfLookup = new Dictionary<string, DxfItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var pdfItem in _pdfItems)
            {
                var normalizedTag = NormalizeTag(pdfItem.Tag);
                if (!pdfLookup.ContainsKey(normalizedTag))
                {
                    pdfLookup[normalizedTag] = pdfItem;
                }
                else
                {
                    // Accumulate quantities for duplicate tags
                    pdfLookup[normalizedTag].Quantity += pdfItem.Quantity;
                }
            }
            
            // Compare each CSV item with PDF
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
                    
                    if (csvItem.Quantity == pdfItem.Quantity)
                    {
                        result.Status = "OK - Quantites identiques";
                        result.StatusColor = new SolidColorBrush(Color.FromRgb(144, 238, 144)); // LightGreen
                    }
                    else
                    {
                        result.Status = $"ERREUR - CSV:{csvItem.Quantity} vs PDF:{pdfItem.Quantity}";
                        result.StatusColor = new SolidColorBrush(Color.FromRgb(255, 99, 71)); // Tomato
                    }
                }
                else
                {
                    result.PdfQuantity = 0;
                    result.Status = "MANQUANT dans PDF";
                    result.StatusColor = new SolidColorBrush(Color.FromRgb(255, 165, 0)); // Orange
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
                        Status = "SUPPLEMENT dans PDF (pas dans CSV)",
                        StatusColor = new SolidColorBrush(Color.FromRgb(135, 206, 250)) // LightSkyBlue
                    });
                }
            }
            
            LogMessage($"[+] Comparaison terminee: {_results.Count} lignes");
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
                
                // Create selection dialog
                var dialog = new Window
                {
                    Title = "Selection du Projet - DXF Verifier",
                    Width = 800,
                    Height = 550,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this,
                    Background = new SolidColorBrush(Color.FromRgb(30, 30, 46)),
                    ResizeMode = ResizeMode.CanResize
                };
                
                var mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                
                // Create ListView
                var listView = new System.Windows.Controls.ListView
                {
                    Margin = new Thickness(15),
                    Background = new SolidColorBrush(Color.FromRgb(37, 37, 54)),
                    Foreground = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(74, 127, 191)),
                    BorderThickness = new Thickness(2)
                };
                
                // Create GridView for columns
                var gridView = new GridView();
                gridView.Columns.Add(new GridViewColumn { Header = "Projet", DisplayMemberBinding = new System.Windows.Data.Binding("ProjectNumber"), Width = 100 });
                gridView.Columns.Add(new GridViewColumn { Header = "Reference", DisplayMemberBinding = new System.Windows.Data.Binding("Reference"), Width = 100 });
                gridView.Columns.Add(new GridViewColumn { Header = "Module", DisplayMemberBinding = new System.Windows.Data.Binding("ModuleNumber"), Width = 100 });
                gridView.Columns.Add(new GridViewColumn { Header = "Chemin", DisplayMemberBinding = new System.Windows.Data.Binding("ProjectPath"), Width = 450 });
                listView.View = gridView;
                
                // Add items
                foreach (var project in projects.OrderByDescending(p => p.ProjectNumber).ThenBy(p => p.Reference).ThenBy(p => p.ModuleNumber))
                {
                    listView.Items.Add(project);
                }
                
                Grid.SetRow(listView, 0);
                mainGrid.Children.Add(listView);
                
                // Button panel
                var buttonPanel = new StackPanel
                {
                    Orientation = System.Windows.Controls.Orientation.Horizontal,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                    Margin = new Thickness(15, 10, 15, 15)
                };
                
                var btnSelect = new System.Windows.Controls.Button
                {
                    Content = "Selectionner",
                    Width = 120,
                    Height = 35,
                    Margin = new Thickness(0, 0, 10, 0),
                    Background = new SolidColorBrush(Color.FromRgb(16, 124, 16)),
                    Foreground = Brushes.White,
                    FontWeight = FontWeights.Bold,
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                
                var btnCancel = new System.Windows.Controls.Button
                {
                    Content = "Annuler",
                    Width = 100,
                    Height = 35,
                    Background = new SolidColorBrush(Color.FromRgb(45, 45, 68)),
                    Foreground = Brushes.White,
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                
                ProjectPathInfo selectedProject = null;
                
                btnSelect.Click += (s, args) =>
                {
                    if (listView.SelectedItem is ProjectPathInfo selected)
                    {
                        selectedProject = selected;
                        dialog.DialogResult = true;
                        dialog.Close();
                    }
                    else
                    {
                        MessageBox.Show("Veuillez selectionner un projet.", "Selection requise", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                };
                
                btnCancel.Click += (s, args) =>
                {
                    dialog.DialogResult = false;
                    dialog.Close();
                };
                
                // Double-click to select
                listView.MouseDoubleClick += (s, args) =>
                {
                    if (listView.SelectedItem is ProjectPathInfo selected)
                    {
                        selectedProject = selected;
                        dialog.DialogResult = true;
                        dialog.Close();
                    }
                };
                
                buttonPanel.Children.Add(btnSelect);
                buttonPanel.Children.Add(btnCancel);
                
                Grid.SetRow(buttonPanel, 1);
                mainGrid.Children.Add(buttonPanel);
                
                dialog.Content = mainGrid;
                
                if (dialog.ShowDialog() == true && selectedProject != null)
                {
                    // Extract reference number (remove REF prefix)
                    var refNumber = selectedProject.Reference;
                    if (refNumber.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
                        refNumber = refNumber.Substring(3);
                    
                    ProjectNumberTextBox.Text = selectedProject.ProjectNumber;
                    ReferenceTextBox.Text = refNumber;
                    ModuleNumberTextBox.Text = selectedProject.ModuleNumber;
                    
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
                    ProjectNumberTextBox.Text = projectInfo.ProjectNumber;
                    ReferenceTextBox.Text = projectInfo.Reference;
                    ModuleNumberTextBox.Text = projectInfo.ModuleNumber;
                    AutoDetectFiles(projectInfo);
                }
            }
        }

        #endregion

        #region Info Button

        private void InfoButton_Click(object sender, RoutedEventArgs e)
        {
            var info = new StringBuilder();
            info.AppendLine("DXF-CSV vs PDF Verifier v1.2");
            info.AppendLine("XNRGY Engineering Automation Tools");
            info.AppendLine("By Mohammed Amine Elgalai");
            info.AppendLine();
            info.AppendLine("=".PadRight(50, '='));
            info.AppendLine("FONCTIONNALITES:");
            info.AppendLine("-".PadRight(50, '-'));
            info.AppendLine("- Detection automatique du PDF ouvert dans Adobe");
            info.AppendLine("- Extraction des tags et quantites depuis CSV/DXF");
            info.AppendLine("- Analyse des tables PDF avec PdfPig");
            info.AppendLine("- Comparaison et verification des donnees");
            info.AppendLine("- Export vers Excel avec EPPlus");
            info.AppendLine();
            info.AppendLine("=".PadRight(50, '='));
            info.AppendLine("STRUCTURE DES DOSSIERS ATTENDUE:");
            info.AppendLine("-".PadRight(50, '-'));
            info.AppendLine(@"C:\Vault\Engineering\Projects\");
            info.AppendLine(@"    [PROJECT_NUMBER]\");
            info.AppendLine(@"        REF[XX]\");
            info.AppendLine(@"            M[XX]\");
            info.AppendLine(@"                03-CutList_[PROJECT]_REF[XX]_M[XX].pdf");
            info.AppendLine(@"                [PROJECT]_REF[XX]_M[XX]-DXF.csv");
            info.AppendLine(@"                [PROJECT]_REF[XX]_M[XX]-DXF.xlsx");
            info.AppendLine();
            info.AppendLine("=".PadRight(50, '='));
            info.AppendLine("PATTERNS DE TAG RECONNUS:");
            info.AppendLine("-".PadRight(50, '-'));
            info.AppendLine("- XX1234 (2-3 lettres + 1-4 chiffres)");
            info.AppendLine("- XX1234-5678 (avec suffixe)");
            info.AppendLine("- XX1234_5678 (avec underscore)");
            info.AppendLine();
            info.AppendLine("Version: 1.2.0 - Migrated to WPF");
            info.AppendLine("Build: " + DateTime.Now.ToString("yyyy-MM-dd"));
            
            MessageBox.Show(info.ToString(), "Information - DXF Verifier", MessageBoxButton.OK, MessageBoxImage.Information);
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
            
            // Update UI (must be on UI thread)
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText(logLine + Environment.NewLine);
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
            LogTextBox.Clear();
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

        private void ShowStatus(string message, StatusType type)
        {
            var prefix = type switch
            {
                StatusType.Success => "[+] ",
                StatusType.Warning => "[!] ",
                StatusType.Error => "[-] ",
                _ => "[i] "
            };
            
            StatusLabel.Text = prefix + message;
            
            StatusLabel.Foreground = type switch
            {
                StatusType.Success => new SolidColorBrush(Color.FromRgb(144, 238, 144)), // LightGreen
                StatusType.Warning => new SolidColorBrush(Color.FromRgb(255, 165, 0)),   // Orange
                StatusType.Error => new SolidColorBrush(Color.FromRgb(255, 99, 71)),     // Tomato
                _ => new SolidColorBrush(Color.FromRgb(224, 224, 224))                    // Light gray
            };
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
