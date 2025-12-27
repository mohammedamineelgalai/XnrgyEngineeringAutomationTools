#######################################################################
#  XNRGY Engineering Automation Tools - BUILD & RUN SCRIPT
#  Version: 2.0.0
#  Author: Mohammed Amine Elgalai
#  Date: 2025-12-26
#
#  Usage:
#    .\build-and-run.ps1              # Build Release + Run
#    .\build-and-run.ps1 -Debug       # Build Debug + Run
#    .\build-and-run.ps1 -Clean       # Clean + Build Release + Run
#    .\build-and-run.ps1 -BuildOnly   # Build sans lancer
#    .\build-and-run.ps1 -KillOnly    # Tuer les instances existantes
#######################################################################

param(
    [switch]$Debug,
    [switch]$Clean,
    [switch]$BuildOnly,
    [switch]$KillOnly
)

$ErrorActionPreference = "Stop"
$ProjectName = "XnrgyEngineeringAutomationTools"
$Configuration = if ($Debug) { "Debug" } else { "Release" }

# Fonction pour afficher le header
function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║     XNRGY ENGINEERING AUTOMATION TOOLS - BUILD & RUN v2.0         ║" -ForegroundColor Cyan
    Write-Host "║     Configuration: $Configuration                                         ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

# Fonction pour trouver MSBuild
function Get-MSBuildPath {
    $paths = @(
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    )
    
    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    throw "MSBuild introuvable! Installer Visual Studio 2022."
}

# Fonction pour tuer les instances
function Stop-ExistingInstances {
    Write-Host "  [1/4] Arret des instances existantes..." -ForegroundColor Yellow
    $processes = Get-Process -Name $ProjectName -ErrorAction SilentlyContinue
    if ($processes) {
        $processes | Stop-Process -Force
        Start-Sleep -Seconds 1
        Write-Host "        ✓ $($processes.Count) instance(s) arretee(s)" -ForegroundColor Green
    } else {
        Write-Host "        ✓ Aucune instance en cours" -ForegroundColor Gray
    }
}

Show-Header

# Mode KillOnly
if ($KillOnly) {
    Stop-ExistingInstances
    Write-Host ""
    Write-Host "  Termine!" -ForegroundColor Green
    exit 0
}

# ETAPE 1: Tuer les instances existantes
Stop-ExistingInstances

# ETAPE 2: Clean (optionnel)
if ($Clean) {
    Write-Host ""
    Write-Host "  [2/4] Nettoyage du projet..." -ForegroundColor Yellow
    Remove-Item -Path "bin\$Configuration" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "obj\$Configuration" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "        ✓ Nettoyage termine" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "  [2/4] Nettoyage..." -ForegroundColor DarkGray
    Write-Host "        - Ignore (utiliser -Clean pour nettoyer)" -ForegroundColor DarkGray
}

# ETAPE 3: Compilation
Write-Host ""
Write-Host "  [3/4] Compilation en mode $Configuration..." -ForegroundColor Yellow

$msbuildPath = Get-MSBuildPath
Write-Host "        MSBuild: $msbuildPath" -ForegroundColor DarkGray

$buildArgs = @(
    "$ProjectName.csproj",
    "/p:Configuration=$Configuration",
    "/t:Rebuild",
    "/v:minimal",
    "/nologo",
    "/m"
)

$buildOutput = & $msbuildPath $buildArgs 2>&1
$buildSuccess = $LASTEXITCODE -eq 0

if (-not $buildSuccess) {
    Write-Host ""
    Write-Host "        ✗ ERREUR DE COMPILATION" -ForegroundColor Red
    Write-Host ""
    Write-Host $buildOutput -ForegroundColor Red
    exit 1
}

# Compter les warnings
$warnings = ($buildOutput | Select-String -Pattern "warning" -AllMatches).Matches.Count
if ($warnings -gt 0) {
    Write-Host "        ✓ Compilation reussie ($warnings warnings)" -ForegroundColor Yellow
} else {
    Write-Host "        ✓ Compilation reussie (0 warnings)" -ForegroundColor Green
}

# Verifier l'executable
$exePath = "bin\$Configuration\$ProjectName.exe"
if (-not (Test-Path $exePath)) {
    Write-Host "        ✗ ERREUR: Executable introuvable: $exePath" -ForegroundColor Red
    exit 1
}

$exeSize = [math]::Round((Get-Item $exePath).Length / 1MB, 2)
Write-Host "        Executable: $exePath ($exeSize MB)" -ForegroundColor DarkGray

# ETAPE 4: Lancement
if ($BuildOnly) {
    Write-Host ""
    Write-Host "  [4/4] Lancement..." -ForegroundColor DarkGray
    Write-Host "        - Ignore (mode BuildOnly)" -ForegroundColor DarkGray
} else {
    Write-Host ""
    Write-Host "  [4/4] Lancement de l'application..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath
    Start-Sleep -Seconds 2
    Write-Host "        ✓ Application lancee!" -ForegroundColor Green
}

# Fin
Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                          TERMINE                                   ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
