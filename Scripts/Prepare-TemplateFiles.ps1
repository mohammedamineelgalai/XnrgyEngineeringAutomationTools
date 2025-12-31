<#
.SYNOPSIS
    Prepare les fichiers IPT d'un dossier template pour Vault
    
.DESCRIPTION
    Ce script ouvre chaque fichier .ipt dans Inventor, applique:
    - Vue Isometrique
    - Zoom All (Fit)
    - Sauvegarde
    - Fermeture
    
    Cela nettoie les fichiers avant de les mettre dans Vault.
    
.PARAMETER FolderPath
    Chemin du dossier contenant les fichiers IPT a preparer
    
.EXAMPLE
    .\Prepare-TemplateFiles.ps1 -FolderPath "C:\Vault\Engineering\Library\Templates\Module_Template"
    
.NOTES
    Auteur: Mohammed Amine Elgalai - XNRGY Climate Systems
    Date: 2025-12-30
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$FolderPath
)

# ============================================================================
# Configuration
# ============================================================================
$ErrorActionPreference = "Continue"
$script:FilesProcessed = 0
$script:FilesSuccess = 0
$script:FilesFailed = 0
$script:TotalFiles = 0

# ============================================================================
# Fonctions
# ============================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "SUCCESS" { "Green" }
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        "INFO"    { "Cyan" }
        default   { "White" }
    }
    
    $prefix = switch ($Level) {
        "SUCCESS" { "[+]" }
        "ERROR"   { "[-]" }
        "WARNING" { "[!]" }
        "INFO"    { "[i]" }
        default   { "[>]" }
    }
    
    Write-Host "[$timestamp] $prefix $Message" -ForegroundColor $color
}

function Get-InventorApplication {
    Write-Log "Connexion a Inventor..." "INFO"
    
    $invApp = $null
    
    # Methode 1: GetActiveObject standard
    try {
        $invApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Inventor.Application")
        if ($null -ne $invApp) {
            Write-Log "Connecte via GetActiveObject" "SUCCESS"
            return $invApp
        }
    }
    catch { }
    
    # Methode 2: Via Get-Process et EnvDTE (si Inventor est ouvert)
    try {
        $invProcess = Get-Process -Name "Inventor" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $invProcess) {
            Write-Log "Processus Inventor detecte (PID: $($invProcess.Id)). Tentative de connexion..." "INFO"
            
            # Attendre un peu et reessayer GetActiveObject
            Start-Sleep -Seconds 2
            $invApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Inventor.Application")
            if ($null -ne $invApp) {
                Write-Log "Connecte apres detection processus" "SUCCESS"
                return $invApp
            }
        }
    }
    catch { }
    
    # Methode 3: Demander a l'utilisateur
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "  INVENTOR NON DETECTE" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Le script ne peut pas se connecter a Inventor." -ForegroundColor White
    Write-Host ""
    Write-Host "  Solutions:" -ForegroundColor Cyan
    Write-Host "  1. Assurez-vous qu'Inventor est ouvert" -ForegroundColor White
    Write-Host "  2. Fermez et rouvrez Inventor (sans mode admin)" -ForegroundColor White
    Write-Host "  3. Executez ce script dans le MEME contexte qu'Inventor" -ForegroundColor White
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""
    
    $choice = Read-Host "Voulez-vous que le script demarre une NOUVELLE instance Inventor? (O/N)"
    
    if ($choice -match "^[OoYy]") {
        Write-Log "Demarrage d'une nouvelle instance Inventor..." "INFO"
        try {
            $invApp = New-Object -ComObject "Inventor.Application"
            $invApp.Visible = $true
            Start-Sleep -Seconds 5
            Write-Log "Inventor demarre (Version: $($invApp.SoftwareVersion.DisplayVersion))" "SUCCESS"
            return $invApp
        }
        catch {
            Write-Log "Impossible de demarrer Inventor: $_" "ERROR"
            return $null
        }
    }
    else {
        Write-Log "Operation annulee" "WARNING"
        return $null
    }
}

function Set-IsometricView {
    param($Document)
    
    try {
        # Pour les fichiers Part (.ipt), utiliser la methode ViewOrientationType sur la vue active
        $view = $Document.Views.Item(1)
        
        if ($view -ne $null) {
            # kIsoTopRightViewOrientation = 10734
            # Utiliser SetViewOrientation pour les parts
            try {
                $camera = $view.Camera
                $camera.ViewOrientationType = 10734  # kIsoTopRightViewOrientation
                $camera.Apply()
            }
            catch {
                # Alternative: utiliser la commande Home View
                try {
                    $view.GoHome()
                }
                catch {
                    # Ignorer si pas supporte
                }
            }
        }
        
        return $true
    }
    catch {
        # Ignorer les erreurs de vue - pas critique
        return $true
    }
}

function Set-ZoomAll {
    param($Document)
    
    try {
        $view = $Document.Views.Item(1)
        $view.Fit()
        return $true
    }
    catch {
        Write-Log "    Erreur Zoom All: $_" "WARNING"
        return $false
    }
}

function Process-IptFile {
    param(
        [System.Object]$InventorApp,
        [string]$FilePath
    )
    
    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $script:FilesProcessed++
    
    Write-Log "[$script:FilesProcessed/$script:TotalFiles] $fileName" "INFO"
    
    $doc = $null
    
    try {
        # Verifier si le fichier est read-only et l'enlever
        $fileInfo = Get-Item $FilePath
        if ($fileInfo.IsReadOnly) {
            $fileInfo.IsReadOnly = $false
        }
        
        # Ouvrir le fichier en mode visible
        $doc = $InventorApp.Documents.Open($FilePath, $true)  # $true = Visible
        
        if ($null -eq $doc) {
            Write-Log "    Echec ouverture" "ERROR"
            $script:FilesFailed++
            return
        }
        
        # Attendre un peu
        Start-Sleep -Milliseconds 300
        
        # Zoom Fit (All)
        try {
            $view = $doc.Views.Item(1)
            if ($null -ne $view) {
                $view.Fit()
            }
        }
        catch { }
        
        # Home View (vue par defaut/iso)
        try {
            $view = $doc.Views.Item(1)
            if ($null -ne $view) {
                $view.GoHome()
            }
        }
        catch { }
        
        # Sauvegarder
        $doc.Save()
        
        # Fermer
        $doc.Close($true)
        $doc = $null
        
        $script:FilesSuccess++
        Write-Log "    OK" "SUCCESS"
    }
    catch {
        Write-Log "    ERREUR: $_" "ERROR"
        $script:FilesFailed++
        
        # Tenter de fermer le document en cas d'erreur
        if ($null -ne $doc) {
            try {
                $doc.Close($true)
            }
            catch { }
        }
    }
}

# ============================================================================
# Main
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PREPARATION FICHIERS TEMPLATE POUR VAULT" -ForegroundColor Cyan
Write-Host "  XNRGY Climate Systems - Mohammed Amine Elgalai" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Demander le dossier si pas fourni
if ([string]::IsNullOrEmpty($FolderPath)) {
    Add-Type -AssemblyName System.Windows.Forms
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Selectionnez le dossier du template Module"
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    $folderBrowser.SelectedPath = "C:\Vault\Engineering\Library"
    
    $result = $folderBrowser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $FolderPath = $folderBrowser.SelectedPath
    }
    else {
        Write-Log "Operation annulee par l'utilisateur" "WARNING"
        exit
    }
}

# Verifier que le dossier existe
if (-not (Test-Path $FolderPath)) {
    Write-Log "Dossier introuvable: $FolderPath" "ERROR"
    exit 1
}

Write-Log "Dossier source: $FolderPath" "INFO"

# Trouver tous les fichiers .ipt recursivment
$iptFiles = Get-ChildItem -Path $FolderPath -Filter "*.ipt" -Recurse -File | 
            Where-Object { $_.Name -notlike "~*" -and $_.Name -notlike "*.bak*" }

$script:TotalFiles = $iptFiles.Count

if ($script:TotalFiles -eq 0) {
    Write-Log "Aucun fichier .ipt trouve dans le dossier" "WARNING"
    exit
}

Write-Log "Fichiers IPT trouves: $script:TotalFiles" "INFO"
Write-Host ""

# Confirmation
$confirmation = Read-Host "Voulez-vous continuer? (O/N)"
if ($confirmation -notmatch "^[OoYy]") {
    Write-Log "Operation annulee" "WARNING"
    exit
}

Write-Host ""

# Connexion a Inventor
$inventor = Get-InventorApplication
if ($null -eq $inventor) {
    Write-Log "Impossible de se connecter a Inventor" "ERROR"
    exit 1
}

Write-Host ""
Write-Log "Debut du traitement de $script:TotalFiles fichiers..." "INFO"
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

# Traiter chaque fichier
foreach ($file in $iptFiles) {
    Process-IptFile -InventorApp $inventor -FilePath $file.FullName
}

# Resume
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Log "Fichiers traites: $script:FilesProcessed" "INFO"
Write-Log "Succes: $script:FilesSuccess" "SUCCESS"

if ($script:FilesFailed -gt 0) {
    Write-Log "Echecs: $script:FilesFailed" "ERROR"
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  TERMINE!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Pause finale
Read-Host "Appuyez sur Entree pour fermer"
