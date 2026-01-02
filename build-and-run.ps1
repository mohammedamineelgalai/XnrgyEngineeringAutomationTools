#######################################################################
#  XNRGY Engineering Automation Tools - BUILD & RUN SCRIPT
#  Version: 1.0.0
#  Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
#  Date: 2026-01-02
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

# Se positionner dans le répertoire du script (important pour chemins relatifs)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $ScriptDir

# Fonction pour afficher le header
function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║     XNRGY ENGINEERING AUTOMATION TOOLS - BUILD & RUN v1.0.0       ║" -ForegroundColor Cyan
    Write-Host "║     Configuration: $Configuration                                         ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

# Fonction pour trouver MSBuild
function Get-MSBuildPath {
    $paths = @(
        # VS 18 (prioritaire)
        "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Insiders\MSBuild\Current\Bin\MSBuild.exe",
        # VS 2022
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
    
    throw "MSBuild introuvable! Installer Visual Studio 2022 ou VS 18."
}

# Fonction pour tuer les instances - FORCE MODE
function Stop-ExistingInstances {
    Write-Host "  [1/4] Arret des instances existantes..." -ForegroundColor Yellow
    
    $killed = $false
    $maxRetries = 5
    $retryDelay = 1
    
    for ($retry = 1; $retry -le $maxRetries; $retry++) {
        # Tuer par nom de processus
        $processes = Get-Process -Name $ProjectName -ErrorAction SilentlyContinue
        if ($processes) {
            foreach ($proc in $processes) {
                try {
                    $proc.Kill()
                    $proc.WaitForExit(3000)
                } catch { }
            }
            Write-Host "        [+] $($processes.Count) instance(s) arretee(s)" -ForegroundColor Green
            $killed = $true
        }
        
        # Aussi tuer VaultAutomationTool si présent (ancien nom)
        $oldProcesses = Get-Process -Name "VaultAutomationTool" -ErrorAction SilentlyContinue
        if ($oldProcesses) {
            foreach ($proc in $oldProcesses) {
                try {
                    $proc.Kill()
                    $proc.WaitForExit(3000)
                } catch { }
            }
            Write-Host "        [+] VaultAutomationTool arrete" -ForegroundColor Green
            $killed = $true
        }
        
        # Forcer avec taskkill /F /T (plus robuste - force + arbre de processus)
        $null = cmd /c "taskkill /F /T /IM $ProjectName.exe 2>nul"
        $null = cmd /c "taskkill /F /T /IM VaultAutomationTool.exe 2>nul"
        
        # Vérifier si vraiment arrêté
        Start-Sleep -Milliseconds 500
        $stillRunning = Get-Process -Name $ProjectName -ErrorAction SilentlyContinue
        $stillRunningOld = Get-Process -Name "VaultAutomationTool" -ErrorAction SilentlyContinue
        
        if (-not $stillRunning -and -not $stillRunningOld) {
            break
        }
        
        if ($retry -lt $maxRetries) {
            Write-Host "        [!] Processus encore actif, nouvelle tentative ($retry/$maxRetries)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $retryDelay
        }
    }
    
    # Vérification finale
    $finalCheck = Get-Process -Name $ProjectName -ErrorAction SilentlyContinue
    if ($finalCheck) {
        Write-Host "        [-] ERREUR: Impossible de fermer l'application apres $maxRetries tentatives" -ForegroundColor Red
        Write-Host "        [-] Fermez l'application manuellement et relancez le script" -ForegroundColor Red
        exit 1
    }
    
    if (-not $killed) {
        Write-Host "        [+] Aucune instance en cours" -ForegroundColor Gray
    }
    
    # Attendre un peu pour libérer les fichiers
    Start-Sleep -Seconds 2
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
    Pop-Location
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
    Pop-Location
    exit 1
}

$exeSize = [math]::Round((Get-Item $exePath).Length / 1MB, 2)
Write-Host "        Executable: $exePath ($exeSize MB)" -ForegroundColor DarkGray

# ETAPE 4: Lancer l'application (sauf BuildOnly)
if (-not $BuildOnly) {
    Write-Host ""
    Write-Host "  [4/4] Lancement de l'application..." -ForegroundColor Yellow
    try {                                                                                       
        Start-Process -FilePath $exePath
        Write-Host "        ✓ Application lancee" -ForegroundColor Green
    } catch {
        Write-Host "        ✗ ERREUR: Impossible de lancer l'application" -ForegroundColor Red
        Pop-Location
        exit 1
    }
} else {
    Write-Host ""
    Write-Host "  [4/4] Lancement de l'application..." -ForegroundColor DarkGray
    Write-Host "        - Ignore (mode BuildOnly)" -ForegroundColor DarkGray
} 

# Restaurer le répertoire original
Pop-Location

# Fin
Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                          TERMINE                                   ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
