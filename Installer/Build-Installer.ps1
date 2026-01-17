# ═══════════════════════════════════════════════════════════════════════════════
# Script: Build-Installer.ps1
# Description: Compile l'application et prepare le package d'installation complet
# Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
# Date: 2026-01-16
# ═══════════════════════════════════════════════════════════════════════════════

param(
    [switch]$SkipBuild,
    [switch]$CreateZip
)

$ErrorActionPreference = "Stop"

# Chemins
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$AppProject = Join-Path $ProjectRoot "XnrgyEngineeringAutomationTools.csproj"
$InstallerProject = Join-Path $ProjectRoot "Installer\XnrgyInstaller.csproj"
$AppBinRelease = Join-Path $ProjectRoot "bin\Release"
$InstallerBinRelease = Join-Path $ProjectRoot "Installer\bin\Release"
$FilesDir = Join-Path $InstallerBinRelease "Files"
$MSBuild = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  XNRGY Engineering Automation Tools - Build Installer Package" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# ETAPE 1: Build de l'application principale
# ─────────────────────────────────────────────────────────────────────────────
if (-not $SkipBuild) {
    Write-Host "[1/5] Compilation de l'application principale..." -ForegroundColor Yellow
    
    if (-not (Test-Path $MSBuild)) {
        $MSBuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    }
    if (-not (Test-Path $MSBuild)) {
        $MSBuild = "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
    }
    
    Push-Location $ProjectRoot
    & $MSBuild $AppProject /t:Rebuild /p:Configuration=Release /m /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[-] Erreur lors de la compilation de l'application" -ForegroundColor Red
        Pop-Location
        exit 1
    }
    Pop-Location
    Write-Host "[+] Application compilee avec succes" -ForegroundColor Green
} else {
    Write-Host "[1/5] Compilation ignoree (SkipBuild)" -ForegroundColor Gray
}

# ─────────────────────────────────────────────────────────────────────────────
# ETAPE 2: Preparation du dossier Files
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[2/5] Preparation du dossier Files..." -ForegroundColor Yellow

if (Test-Path $FilesDir) {
    Remove-Item $FilesDir -Recurse -Force
}
New-Item -ItemType Directory -Path $FilesDir -Force | Out-Null

# ─────────────────────────────────────────────────────────────────────────────
# ETAPE 3: Copie de tous les fichiers necessaires
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[3/5] Copie des fichiers de l'application..." -ForegroundColor Yellow

# Extensions a copier
$Extensions = @("*.exe", "*.dll", "*.config", "*.ico", "*.png", "*.json")

# Fichiers/dossiers a exclure
$Excludes = @(
    "*.pdb",
    "*.xml",
    "*.vshost.*",
    "*.manifest",
    "WebView2Loader.dll"  # Necessaire uniquement si WebView2 est utilise
)

$fileCount = 0
foreach ($ext in $Extensions) {
    $files = Get-ChildItem -Path $AppBinRelease -Filter $ext -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        $exclude = $false
        foreach ($pattern in $Excludes) {
            if ($file.Name -like $pattern) {
                $exclude = $true
                break
            }
        }
        if (-not $exclude) {
            Copy-Item $file.FullName -Destination $FilesDir -Force
            $fileCount++
        }
    }
}

# Copier l'icone depuis Resources si pas dans bin
$IconSource = Join-Path $ProjectRoot "Resources\XnrgyEngineeringAutomationTools.ico"
$IconDest = Join-Path $FilesDir "XnrgyEngineeringAutomationTools.ico"
if ((Test-Path $IconSource) -and -not (Test-Path $IconDest)) {
    Copy-Item $IconSource -Destination $IconDest -Force
    $fileCount++
}

# Creer le dossier Logs
$LogsDir = Join-Path $FilesDir "Logs"
New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
"# Logs directory" | Out-File (Join-Path $LogsDir ".gitkeep") -Encoding UTF8

Write-Host "[+] $fileCount fichiers copies" -ForegroundColor Green

# ─────────────────────────────────────────────────────────────────────────────
# ETAPE 4: Build de l'installateur
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[4/5] Compilation de l'installateur..." -ForegroundColor Yellow

Push-Location (Join-Path $ProjectRoot "Installer")
& $MSBuild $InstallerProject /t:Rebuild /p:Configuration=Release /m /v:minimal /nologo
if ($LASTEXITCODE -ne 0) {
    Write-Host "[-] Erreur lors de la compilation de l'installateur" -ForegroundColor Red
    Pop-Location
    exit 1
}
Pop-Location
Write-Host "[+] Installateur compile avec succes" -ForegroundColor Green

# ─────────────────────────────────────────────────────────────────────────────
# ETAPE 5: Resume et creation ZIP optionnel
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[5/5] Finalisation..." -ForegroundColor Yellow

$SetupExe = Join-Path $InstallerBinRelease "XNRGYEngineeringAutomationToolsSetup.exe"
$SetupSize = [math]::Round((Get-Item $SetupExe).Length / 1MB, 2)
$FilesSize = [math]::Round((Get-ChildItem $FilesDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB, 2)

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  [+] BUILD TERMINE AVEC SUCCES" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
Write-Host "  Installateur: $SetupExe" -ForegroundColor White
Write-Host "  Taille Setup: $SetupSize MB" -ForegroundColor White
Write-Host "  Taille Files: $FilesSize MB" -ForegroundColor White
Write-Host "  Fichiers:     $fileCount" -ForegroundColor White
Write-Host ""

if ($CreateZip) {
    Write-Host "  Creation du package ZIP..." -ForegroundColor Yellow
    $ZipName = "XNRGYEngineeringAutomationToolsSetup_v1.0.0.zip"
    $ZipPath = Join-Path $InstallerBinRelease $ZipName
    
    if (Test-Path $ZipPath) {
        Remove-Item $ZipPath -Force
    }
    
    # Creer un dossier temporaire pour le ZIP
    $TempDir = Join-Path $env:TEMP "XnrgySetupPackage"
    if (Test-Path $TempDir) {
        Remove-Item $TempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    
    # Copier Setup + Files
    Copy-Item $SetupExe -Destination $TempDir -Force
    Copy-Item $FilesDir -Destination (Join-Path $TempDir "Files") -Recurse -Force
    
    # Creer le ZIP
    Compress-Archive -Path "$TempDir\*" -DestinationPath $ZipPath -Force
    
    $ZipSize = [math]::Round((Get-Item $ZipPath).Length / 1MB, 2)
    Write-Host "  [+] Package ZIP cree: $ZipPath ($ZipSize MB)" -ForegroundColor Green
    
    # Nettoyer
    Remove-Item $TempDir -Recurse -Force
}

Write-Host ""
Write-Host "  Pour distribuer:" -ForegroundColor Cyan
Write-Host "  1. Copier le dossier 'Installer\bin\Release' complet" -ForegroundColor White
Write-Host "  2. Ou utiliser -CreateZip pour creer un package ZIP" -ForegroundColor White
Write-Host ""
