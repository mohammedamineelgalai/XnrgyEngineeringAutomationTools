#######################################################################
#  XNRGY Engineering Automation Tools - BUILD & RUN SCRIPT
#  Version: 2.3.0
#  Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
#  Date: 2026-01-17
#
#  Usage:
#    .\build-and-run.ps1                    # Build Release + Run
#    .\build-and-run.ps1 -Debug             # Build Debug + Run
#    .\build-and-run.ps1 -Clean             # Clean + Build Release + Run
#    .\build-and-run.ps1 -BuildOnly         # Build sans lancer
#    .\build-and-run.ps1 -KillOnly          # Tuer les instances existantes
#    .\build-and-run.ps1 -WithInstaller     # Build App + Installer
#    .\build-and-run.ps1 -InstallerOnly     # Build uniquement l'installateur
#    .\build-and-run.ps1 -CreatePackage     # Build App + Installer + ZIP
#    .\build-and-run.ps1 -Deploy            # Build + Deploy Firebase (HTML + Rules + Config)
#    .\build-and-run.ps1 -DeployAll         # Deploy Firebase complet (hosting + rules + database)
#    .\build-and-run.ps1 -WithInstaller -Deploy  # Build + Installer + Deploy Firebase complet
#    .\build-and-run.ps1 -Publish           # Build + Installer + Publish to GitHub Releases
#    .\build-and-run.ps1 -Publish -NewVersion "1.1.0"  # Create new version on GitHub
#    .\build-and-run.ps1 -SyncFirebase      # Sync Firebase data AVANT deploiement
#    .\build-and-run.ps1 -Auto              # [ROBO MODE] Tout automatique: Clean+Build+Sign+Installer+Sync+Deploy+Publish
#    .\build-and-run.ps1 -Auto              # [ROBO MODE] Tout automatique: Build+Sign+Installer+Sync+Deploy+Publish
#######################################################################

param(
    [switch]$Debug,
    [switch]$Clean,
    [switch]$BuildOnly,
    [switch]$KillOnly,
    [switch]$WithInstaller,
    [switch]$InstallerOnly,
    [switch]$CreatePackage,
    [switch]$DeployAdmin,      # Legacy - Deploy hosting only
    [switch]$Deploy,           # Deploy hosting + rules + config
    [switch]$DeployAll,        # Alias pour Deploy
    [switch]$Publish,          # Publish to GitHub Releases
    [switch]$SyncFirebase,     # Sync Firebase data avant deploy
    [switch]$Auto,             # MODE ROBO: Tout automatique dans l'ordre
    [string]$NewVersion = ""   # New version number (empty = update existing)
)

# ═══════════════════════════════════════════════════════════════════════════════
# MODE AUTO (ROBO) - Sequence complete automatique
# ═══════════════════════════════════════════════════════════════════════════════
if ($Auto) {
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║     [ROBO MODE] SEQUENCE AUTOMATIQUE COMPLETE ACTIVEE             ║" -ForegroundColor Magenta
    Write-Host "╠═══════════════════════════════════════════════════════════════════╣" -ForegroundColor Magenta
    Write-Host "║  Sequence obligatoire:                                            ║" -ForegroundColor Magenta
    Write-Host "║    1. Kill instances                                              ║" -ForegroundColor White
    Write-Host "║    2. Clean projet                                                ║" -ForegroundColor White
    Write-Host "║    3. Build Release + Signature                                   ║" -ForegroundColor White
    Write-Host "║    4. Build Installer + Signature                                 ║" -ForegroundColor White
    Write-Host "║    5. Sync Firebase (preserver devices/users/auditLog)            ║" -ForegroundColor White
    Write-Host "║    6. Deploy Firebase (Hosting + Rules + Config)                  ║" -ForegroundColor White
    Write-Host "║    7. Publish GitHub Releases                                     ║" -ForegroundColor White
    Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""
    
    # Activer tous les flags necessaires
    $Clean = $true
    $WithInstaller = $true
    $SyncFirebase = $true
    $Deploy = $true
    $Publish = $true
}

# Si DeployAll est specifie, activer Deploy
if ($DeployAll) { $Deploy = $true }
# Si Publish est specifie, activer WithInstaller
if ($Publish) { $WithInstaller = $true }

$ErrorActionPreference = "Stop"
$ProjectName = "XnrgyEngineeringAutomationTools"
$InstallerName = "XnrgyInstaller"
$Configuration = if ($Debug) { "Debug" } else { "Release" }

# GitHub Release Settings
$GitHubOwner = "mohammedamineelgalai"
$GitHubRepo = "XnrgyEngineeringAutomationTools"
$GitHubReleasesUrl = "https://github.com/$GitHubOwner/$GitHubRepo/releases"

# Chemins Firebase
$FirebaseDir = "$PSScriptRoot\Firebase Realtime Database configuration"
$FirebaseCLI = "$FirebaseDir\firebase-tools-instant-win.exe"
$AdminPanelDir = "$FirebaseDir\admin-panel"
$FirebaseRulesFile = "$FirebaseDir\firebase-rules.json"
$FirebaseInitFile = "$FirebaseDir\firebase-init.json"
$FirebaseDatabaseURL = "https://xeat-remote-control-default-rtdb.firebaseio.com"
$SyncFirebaseScript = "$PSScriptRoot\Sync-Firebase.ps1"

# Code Signing Settings
$SigningDir = "$PSScriptRoot\.signing"
$SigningCertFile = "$SigningDir\XnrgyCodeSigning.pfx"
$SigningCertPassword = "Xnrgy2026!"
$SigningTimestampServer = "http://timestamp.digicert.com"

# Calculer le nombre d'etapes dynamiquement
$TotalSteps = 4  # Base: Kill, Clean, Build App, Run
if ($WithInstaller -or $InstallerOnly -or $CreatePackage) { $TotalSteps = 6 }
if ($CreatePackage) { $TotalSteps = 7 }
if ($Deploy) { $TotalSteps += 3 }  # +3 etapes: Hosting, Rules, Config
if ($Publish) { $TotalSteps += 1 }  # +1 etape: GitHub Release
if ($DeployAdmin) { $TotalSteps = 1 }  # Mode standalone legacy

# Se positionner dans le répertoire du script (important pour chemins relatifs)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $ScriptDir

# Fonction pour afficher le header
function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║     XNRGY ENGINEERING AUTOMATION TOOLS - BUILD & RUN v2.3.0       ║" -ForegroundColor Cyan
    if ($Auto) {
        Write-Host "║     Mode: [ROBO] AUTOMATIQUE COMPLET                              ║" -ForegroundColor Magenta
    } else {
        Write-Host "║     Configuration: $($Configuration.PadRight(43))║" -ForegroundColor Cyan
    }
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

# Fonction pour copier TOUTES les dependances
function Copy-AllDependencies {
    param([string]$SourceDir, [string]$DestDir)
    
    Write-Host "        [>] Copie de TOUTES les dependances..." -ForegroundColor Cyan
    
    # Supprimer et recreer le dossier destination
    if (Test-Path $DestDir) {
        Remove-Item $DestDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    
    # Extensions a copier (TOUT ce qui est necessaire)
    $IncludeExtensions = @("*.exe", "*.dll", "*.config", "*.json", "*.ico", "*.png", "*.jpg", "*.gif", "*.bmp", "*.xaml", "*.resources", "*.pri")
    
    # Fichiers a exclure
    $ExcludePatterns = @("*.pdb", "*.xml", "*.vshost.*", "*.CodeAnalysisLog.*", "*TestAdapter*", "*nunit*", "*xunit*", "*moq*")
    
    $totalFiles = 0
    $totalSize = 0
    
    # Copier tous les fichiers
    foreach ($ext in $IncludeExtensions) {
        $files = Get-ChildItem -Path $SourceDir -Filter $ext -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $exclude = $false
            foreach ($pattern in $ExcludePatterns) {
                if ($file.Name -like $pattern) { $exclude = $true; break }
            }
            if (-not $exclude) {
                Copy-Item $file.FullName -Destination $DestDir -Force
                $totalFiles++
                $totalSize += $file.Length
            }
        }
    }
    
    # Copier les sous-dossiers importants (runtimes, etc.)
    $SubFolders = @("runtimes", "ref", "lib", "Resources", "Themes", "Assets", "Images", "x86", "x64")
    foreach ($folder in $SubFolders) {
        $subPath = Join-Path $SourceDir $folder
        if (Test-Path $subPath) {
            Copy-Item $subPath -Destination (Join-Path $DestDir $folder) -Recurse -Force
            $subFiles = (Get-ChildItem (Join-Path $DestDir $folder) -Recurse -File -ErrorAction SilentlyContinue).Count
            $totalFiles += $subFiles
            Write-Host "        [+] Dossier '$folder' ($subFiles fichiers)" -ForegroundColor Gray
        }
    }
    
    # Copier l'icone depuis Resources si manquante
    $IconSource = Join-Path $ScriptDir "Resources\XnrgyEngineeringAutomationTools.ico"
    $IconDest = Join-Path $DestDir "XnrgyEngineeringAutomationTools.ico"
    if ((Test-Path $IconSource) -and -not (Test-Path $IconDest)) {
        Copy-Item $IconSource -Destination $IconDest -Force
        $totalFiles++
    }
    
    # Creer le dossier Logs
    New-Item -ItemType Directory -Path (Join-Path $DestDir "Logs") -Force | Out-Null
    
    $sizeMB = [math]::Round($totalSize / 1MB, 2)
    Write-Host "        [+] $totalFiles fichiers copies ($sizeMB MB)" -ForegroundColor Green
    
    # Verifier les DLL critiques (noms exacts des packages NuGet)
    $CriticalDLLs = @(
        "Newtonsoft.Json.dll",
        "UglyToad.PdfPig.dll",
        "EPPlus.dll",
        "NLog.dll",
        "CommunityToolkit.Mvvm.dll",
        "System.Text.Json.dll",
        "System.Memory.dll",
        "System.Buffers.dll"
    )
    
    $missing = @()
    foreach ($dll in $CriticalDLLs) {
        if (-not (Test-Path (Join-Path $DestDir $dll))) { $missing += $dll }
    }
    
    if ($missing.Count -gt 0) {
        Write-Host "        [!] DLLs manquantes: $($missing -join ', ')" -ForegroundColor Yellow
    } else {
        Write-Host "        [+] Toutes les DLLs critiques presentes" -ForegroundColor Green
    }
    
    return $totalFiles
}

# Fonction pour signer un executable avec le certificat XNRGY
function Sign-Executable {
    param([string]$ExePath)
    
    if (-not (Test-Path $SigningCertFile)) {
        Write-Host "        [!] Certificat de signature introuvable: $SigningCertFile" -ForegroundColor Yellow
        Write-Host "        [i] L'executable ne sera pas signe" -ForegroundColor Gray
        return $false
    }
    
    # Trouver signtool.exe
    $signtoolPaths = @(
        "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64\signtool.exe",
        "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64\signtool.exe",
        "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe",
        "C:\Program Files (x86)\Windows Kits\10\App Certification Kit\signtool.exe"
    )
    
    $signtool = $null
    foreach ($path in $signtoolPaths) {
        if (Test-Path $path) {
            $signtool = $path
            break
        }
    }
    
    if (-not $signtool) {
        Write-Host "        [!] signtool.exe introuvable - Installation Windows SDK requise" -ForegroundColor Yellow
        return $false
    }
    
    try {
        Write-Host "        [>] Signature de: $(Split-Path $ExePath -Leaf)" -ForegroundColor Cyan
        
        $signArgs = @(
            "sign",
            "/f", $SigningCertFile,
            "/p", $SigningCertPassword,
            "/tr", $SigningTimestampServer,
            "/td", "sha256",
            "/fd", "sha256",
            "/d", "XNRGY Engineering Automation Tools",
            "/du", "https://github.com/mohammedamineelgalai/XnrgyEngineeringAutomationTools",
            $ExePath
        )
        
        $signOutput = & $signtool $signArgs 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "        [+] Signature reussie!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "        [-] Echec signature: $signOutput" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "        [-] Erreur signature: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

Show-Header

# Mode DeployAdmin - Deployer admin-panel sur Firebase
if ($DeployAdmin) {
    Write-Host "  [1/1] Deploiement Admin Panel sur Firebase Hosting..." -ForegroundColor Yellow
    
    if (-not (Test-Path $FirebaseCLI)) {
        Write-Host "        [-] Firebase CLI introuvable: $FirebaseCLI" -ForegroundColor Red
        Pop-Location
        exit 1
    }
    
    $adminPanelDir = "$ScriptDir\Firebase Realtime Database configuration\admin-panel"
    Push-Location $adminPanelDir
    
    & $FirebaseCLI deploy --only hosting
    $deployResult = $LASTEXITCODE
    
    Pop-Location
    Pop-Location
    
    if ($deployResult -eq 0) {
        Write-Host "        [+] Deploiement reussi!" -ForegroundColor Green
    } else {
        Write-Host "        [-] Echec du deploiement" -ForegroundColor Red
    }
    exit $deployResult
}

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

# Signer l'executable principal
Write-Host "        [>] Signature de l'executable..." -ForegroundColor Cyan
Sign-Executable -ExePath (Join-Path $PSScriptRoot $exePath)

# ═══════════════════════════════════════════════════════════════════════════════
# ETAPE 4: Compilation de l'installateur AUTO-EXTRACTIBLE (si demande)
# ═══════════════════════════════════════════════════════════════════════════════
if ($WithInstaller -or $InstallerOnly -or $CreatePackage) {
    Write-Host ""
    Write-Host "  [4/$TotalSteps] Preparation de l'installateur AUTO-EXTRACTIBLE..." -ForegroundColor Yellow
    
    $InstallerBinRelease = "Installer\bin\$Configuration"
    $FilesDir = "$InstallerBinRelease\Files"
    $SourceBinDir = "bin\$Configuration"
    $EmbeddedZipPath = "Installer\Resources\AppFiles.zip"
    
    # Utiliser la nouvelle fonction qui copie TOUT
    $fileCount = Copy-AllDependencies -SourceDir $SourceBinDir -DestDir $FilesDir
    
    # NOUVEAU: Creer le ZIP embarque pour l'installateur auto-extractible
    Write-Host "        [>] Creation du ZIP embarque pour installateur..." -ForegroundColor Cyan
    
    # Supprimer l'ancien ZIP s'il existe
    if (Test-Path $EmbeddedZipPath) { Remove-Item $EmbeddedZipPath -Force }
    
    # Creer le ZIP des fichiers de l'application
    Compress-Archive -Path "$FilesDir\*" -DestinationPath $EmbeddedZipPath -CompressionLevel Optimal -Force
    
    $zipSize = [math]::Round((Get-Item $EmbeddedZipPath).Length / 1MB, 2)
    Write-Host "        [+] ZIP embarque cree: $zipSize MB" -ForegroundColor Green
    
    # Compiler l'installateur (avec le ZIP embarque)
    Write-Host ""
    Write-Host "  [5/$TotalSteps] Compilation de l'installateur..." -ForegroundColor Yellow
    
    $installerArgs = @(
        "Installer\$InstallerName.csproj",
        "/p:Configuration=$Configuration",
        "/t:Rebuild",
        "/v:minimal",
        "/nologo",
        "/m"
    )
    
    $installerOutput = & $msbuildPath $installerArgs 2>&1
    $installerSuccess = $LASTEXITCODE -eq 0
    
    if (-not $installerSuccess) {
        Write-Host "        [-] ERREUR compilation installateur" -ForegroundColor Red
        Write-Host $installerOutput -ForegroundColor Red
        Pop-Location
        exit 1
    }
    
    $setupExe = "$InstallerBinRelease\XNRGYEngineeringAutomationToolsSetup.exe"
    if (Test-Path $setupExe) {
        $setupSize = [math]::Round((Get-Item $setupExe).Length / 1MB, 2)
        Write-Host "        [+] Installateur AUTO-EXTRACTIBLE compile ($setupSize MB)" -ForegroundColor Green
        
        # Verifier que la taille est coherente (doit etre > taille du ZIP)
        if ($setupSize -lt $zipSize) {
            Write-Host "        [!] ATTENTION: Taille installateur ($setupSize MB) < ZIP ($zipSize MB)" -ForegroundColor Yellow
            Write-Host "        [!] Le ZIP n'est peut-etre pas embarque correctement" -ForegroundColor Yellow
        } else {
            Write-Host "        [+] Verification OK: Installateur contient les fichiers embarques" -ForegroundColor Green
        }
        
        # Signer l'installateur
        Write-Host "        [>] Signature de l'installateur..." -ForegroundColor Cyan
        Sign-Executable -ExePath (Join-Path $PSScriptRoot $setupExe)
    } else {
        Write-Host "        [-] Installateur non trouve" -ForegroundColor Red
    }
    
    # Creer package ZIP additionnel si demande (pour distribution alternative)
    if ($CreatePackage) {
        Write-Host ""
        Write-Host "  [6/$TotalSteps] Creation du package ZIP (distribution alternative)..." -ForegroundColor Yellow
        
        $ZipName = "XNRGYEngineeringAutomationToolsSetup_v1.0.0.zip"
        $ZipPath = "$InstallerBinRelease\$ZipName"
        
        if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
        
        # Creer dossier temp pour le ZIP
        $TempDir = Join-Path $env:TEMP "XnrgySetupPackage"
        if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
        
        Copy-Item $setupExe -Destination $TempDir -Force
        Copy-Item $FilesDir -Destination "$TempDir\Files" -Recurse -Force
        
        Compress-Archive -Path "$TempDir\*" -DestinationPath $ZipPath -Force
        
        $ZipSize = [math]::Round((Get-Item $ZipPath).Length / 1MB, 2)
        Write-Host "        [+] Package ZIP cree ($ZipSize MB)" -ForegroundColor Green
        Write-Host "        Chemin: $ZipPath" -ForegroundColor DarkGray
        
        Remove-Item $TempDir -Recurse -Force
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# ETAPE SYNC: Synchroniser Firebase AVANT deploiement (preserver devices/users)
# ═══════════════════════════════════════════════════════════════════════════════
if ($SyncFirebase -or $Deploy) {
    Write-Host ""
    Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
    Write-Host "        SYNCHRONISATION FIREBASE (Preservation donnees)" -ForegroundColor Magenta
    Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
    
    if (Test-Path $SyncFirebaseScript) {
        Write-Host "        [>] Telechargement des donnees dynamiques..." -ForegroundColor Yellow
        try {
            # Executer le script de sync
            & $SyncFirebaseScript -PreDeploy
            Write-Host "        [+] Synchronisation terminee" -ForegroundColor Green
        }
        catch {
            Write-Host "        [!] Avertissement sync: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "        [i] Deploiement continue avec les donnees locales" -ForegroundColor Gray
        }
    } else {
        Write-Host "        [!] Script Sync-Firebase.ps1 introuvable" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ═══════════════════════════════════════════════════════════════════════════════
# ETAPE DEPLOY: Deployer Firebase complet (Hosting + Rules + Config)
# ═══════════════════════════════════════════════════════════════════════════════
if ($Deploy) {
    $baseStep = if ($CreatePackage) { 7 } elseif ($WithInstaller -or $InstallerOnly) { 6 } else { 4 }
    
    # Verifier que les fichiers existent
    if (-not (Test-Path $FirebaseCLI)) {
        Write-Host "        [-] Firebase CLI introuvable: $FirebaseCLI" -ForegroundColor Red
    } elseif (-not (Test-Path $AdminPanelDir)) {
        Write-Host "        [-] Dossier admin-panel introuvable: $AdminPanelDir" -ForegroundColor Red
    } else {
        Write-Host ""
        Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host "        DEPLOIEMENT FIREBASE COMPLET" -ForegroundColor Cyan
        Write-Host "        Projet: xeat-remote-control" -ForegroundColor Gray
        Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
        
        # Lancer le terminal Firebase UNE SEULE FOIS pour toutes les commandes
        Write-Host ""
        Write-Host "  [$($baseStep + 1)/$TotalSteps] Ouverture terminal Firebase..." -ForegroundColor Yellow
        
        $proc = Start-Process -FilePath $FirebaseCLI -WorkingDirectory $FirebaseDir -PassThru
        
        Write-Host "        [~] Attente initialisation Firebase CLI (~6 sec)..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 6
        
        # Configurer SendKeys
        Add-Type -AssemblyName System.Windows.Forms
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.AppActivate($proc.Id) | Out-Null
        Start-Sleep -Milliseconds 500
        
        # ─────────────────────────────────────────────────────────────────────
        # ETAPE A: Deployer Hosting (admin-panel HTML)
        # ─────────────────────────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  [$($baseStep + 1)/$TotalSteps] Deploiement HOSTING (admin-panel)..." -ForegroundColor Yellow
        Write-Host "        [>] URL: https://xeat-remote-control.web.app/" -ForegroundColor Cyan
        
        [System.Windows.Forms.SendKeys]::SendWait("firebase deploy --only hosting")
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        
        Write-Host "        [+] Commande hosting envoyee!" -ForegroundColor Green
        Write-Host "        [~] Attente deploiement hosting (~15 sec)..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 15
        
        # Reactiver la fenetre
        $wshell.AppActivate($proc.Id) | Out-Null
        Start-Sleep -Milliseconds 300
        
        # ─────────────────────────────────────────────────────────────────────
        # ETAPE B: Deployer Rules (firebase deploy --only database)
        # ─────────────────────────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  [$($baseStep + 2)/$TotalSteps] Deploiement RULES (database.rules.json)..." -ForegroundColor Yellow
        
        if (Test-Path $FirebaseRulesFile) {
            Write-Host "        [>] Fichier: $FirebaseRulesFile" -ForegroundColor Gray
            Write-Host "        [>] Commande: firebase deploy --only database" -ForegroundColor Gray
            
            # La bonne commande est firebase deploy --only database
            # qui utilise le fichier database.rules.json configure dans firebase.json
            [System.Windows.Forms.SendKeys]::SendWait("firebase deploy --only database")
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            
            Write-Host "        [+] Commande rules envoyee!" -ForegroundColor Green
            Write-Host "        [~] Attente deploiement rules (~8 sec)..." -ForegroundColor DarkGray
            Start-Sleep -Seconds 8
            
            # Reactiver la fenetre
            $wshell.AppActivate($proc.Id) | Out-Null
            Start-Sleep -Milliseconds 300
        } else {
            Write-Host "        [!] Fichier firebase-rules.json introuvable - IGNORE" -ForegroundColor Yellow
        }
        
        # ─────────────────────────────────────────────────────────────────────
        # ETAPE C: Deployer Config initiale (firebase-init.json) via CLI
        # NOTE: Cette etape est OPTIONNELLE car Sync-Firebase preserve les donnees
        # ─────────────────────────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  [$($baseStep + 3)/$TotalSteps] Deploiement CONFIG (firebase-init.json)..." -ForegroundColor Yellow
        
        if (Test-Path $FirebaseInitFile) {
            Write-Host "        [>] Fichier: $FirebaseInitFile" -ForegroundColor Gray
            Write-Host "        [>] Database: $FirebaseDatabaseURL" -ForegroundColor Gray
            
            # Convertir en JSON standard (sans les doubles espaces de PowerShell)
            $tempJsonFile = "$FirebaseDir\config-deploy.json"
            $jsonObject = Get-Content $FirebaseInitFile -Raw -Encoding UTF8 | ConvertFrom-Json
            
            # Utiliser Newtonsoft.Json si disponible, sinon compacter avec regex
            $jsonCompact = $jsonObject | ConvertTo-Json -Depth 100 -Compress
            # Nettoyer les artefacts PowerShell (doubles espaces apres :)
            $jsonCompact = $jsonCompact -replace ':\s+', ':'
            $jsonCompact = $jsonCompact -replace ',\s+', ','
            
            # Ecrire en UTF-8 sans BOM
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::WriteAllText($tempJsonFile, $jsonCompact, $utf8NoBom)
            
            $jsonSize = [math]::Round((Get-Item $tempJsonFile).Length / 1KB, 2)
            Write-Host "        [i] JSON nettoye: $jsonSize KB (UTF-8 sans BOM)" -ForegroundColor DarkGray
            
            # Utiliser database:set avec le fichier
            [System.Windows.Forms.SendKeys]::SendWait("firebase database:set / config-deploy.json")
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            
            # Attendre que la question de confirmation apparaisse
            Start-Sleep -Seconds 3
            
            # Repondre "y" a la confirmation "You are about to overwrite all data..."
            $wshell.AppActivate($proc.Id) | Out-Null
            Start-Sleep -Milliseconds 300
            [System.Windows.Forms.SendKeys]::SendWait("y")
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            
            Write-Host "        [+] Commande config envoyee + confirmation 'y'!" -ForegroundColor Green
            Write-Host "        [~] Attente deploiement config (~20 sec pour gros fichier)..." -ForegroundColor DarkGray
            Start-Sleep -Seconds 20
            
            # Nettoyer le fichier temporaire
            Remove-Item $tempJsonFile -Force -ErrorAction SilentlyContinue
        } else {
            Write-Host "        [!] Fichier firebase-init.json introuvable - IGNORE" -ForegroundColor Yellow
        }
        
        Write-Host ""
        Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "        DEPLOIEMENT FIREBASE TERMINE!" -ForegroundColor Green
        Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host ""
        Write-Host "        [i] Admin Panel: https://xeat-remote-control.web.app/" -ForegroundColor Cyan
        Write-Host "        [i] Database:    $FirebaseDatabaseURL" -ForegroundColor Cyan
        Write-Host ""
        
        # Fermer automatiquement le terminal Firebase CLI
        Write-Host "        [>] Fermeture du terminal Firebase CLI..." -ForegroundColor Gray
        try {
            if ($proc -and -not $proc.HasExited) {
                # Envoyer "exit" pour fermer proprement
                $wshell.AppActivate($proc.Id) | Out-Null
                Start-Sleep -Milliseconds 300
                [System.Windows.Forms.SendKeys]::SendWait("exit")
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                Start-Sleep -Seconds 2
                
                # Si toujours actif, forcer la fermeture
                if (-not $proc.HasExited) {
                    $proc.Kill()
                }
                Write-Host "        [+] Terminal Firebase CLI ferme" -ForegroundColor Green
            }
        } catch {
            Write-Host "        [!] Impossible de fermer le terminal Firebase: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# ETAPE PUBLISH: Publier sur GitHub Releases via API REST
# ═══════════════════════════════════════════════════════════════════════════════
if ($Publish) {
    $publishStep = $TotalSteps - 1  # Avant derniere etape (avant Run)
    if ($BuildOnly -or $InstallerOnly -or $CreatePackage) { $publishStep = $TotalSteps }
    
    Write-Host ""
    Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
    Write-Host "        PUBLICATION GITHUB RELEASES (API)" -ForegroundColor Magenta
    Write-Host "        Repo: $GitHubOwner/$GitHubRepo" -ForegroundColor Gray
    Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Magenta
    
    # Verifier que l'installateur existe
    $InstallerBinRelease = "Installer\bin\$Configuration"
    $setupExe = "$InstallerBinRelease\XNRGYEngineeringAutomationToolsSetup.exe"
    $setupExeFullPath = Join-Path $PSScriptRoot $setupExe
    $filesDir = Join-Path $PSScriptRoot "$InstallerBinRelease\Files"
    
    if (-not (Test-Path $setupExeFullPath)) {
        Write-Host "        [-] Installateur introuvable: $setupExe" -ForegroundColor Red
        Write-Host "        [!] Utilisez -WithInstaller pour creer l'installateur d'abord" -ForegroundColor Yellow
    } elseif (-not (Test-Path $filesDir)) {
        Write-Host "        [-] Dossier Files introuvable: $filesDir" -ForegroundColor Red
    } else {
        # CREER UN ZIP contenant Setup.exe + Files/
        Write-Host ""
        Write-Host "  [$publishStep/$TotalSteps] Creation du package ZIP pour GitHub..." -ForegroundColor Yellow
        
        $zipFileName = "XNRGYEngineeringAutomationToolsSetup.zip"
        $zipPath = Join-Path $PSScriptRoot "$InstallerBinRelease\$zipFileName"
        
        # Supprimer l'ancien ZIP s'il existe
        if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
        
        # Creer un dossier temporaire pour le ZIP
        $tempZipDir = Join-Path $env:TEMP "XnrgySetupPackage_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        if (Test-Path $tempZipDir) { Remove-Item $tempZipDir -Recurse -Force }
        New-Item -ItemType Directory -Path $tempZipDir -Force | Out-Null
        
        # Copier Setup.exe et Files/
        Copy-Item $setupExeFullPath -Destination $tempZipDir -Force
        Copy-Item $filesDir -Destination (Join-Path $tempZipDir "Files") -Recurse -Force
        
        # Creer le ZIP
        Compress-Archive -Path "$tempZipDir\*" -DestinationPath $zipPath -Force
        
        # Nettoyer
        Remove-Item $tempZipDir -Recurse -Force
        
        $zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
        Write-Host "        [+] Package ZIP cree: $zipFileName ($zipSize MB)" -ForegroundColor Green
        
        # Maintenant uploader le ZIP au lieu du .exe seul
        Write-Host ""
        Write-Host "  Publication via GitHub API..." -ForegroundColor Yellow
        
        # Token GitHub - Ordre de priorite:
        # 1. Variable d'environnement GITHUB_TOKEN (comme backup-all-projects-complete.ps1)
        # 2. Fichier local .github-token
        # 3. Demander a l'utilisateur
        $GitHubToken = $null
        $tokenFile = "$PSScriptRoot\.github-token"
        
        # 1. Essayer la variable d'environnement
        if ($env:GITHUB_TOKEN) {
            $GitHubToken = $env:GITHUB_TOKEN
            Write-Host "        [+] Token charge depuis `$env:GITHUB_TOKEN" -ForegroundColor Green
        }
        # 2. Essayer le fichier local
        elseif (Test-Path $tokenFile) {
            $GitHubToken = (Get-Content $tokenFile -Raw).Trim()
            Write-Host "        [+] Token charge depuis .github-token" -ForegroundColor Green
        }
        
        # 3. Si toujours pas de token, demander
        if ([string]::IsNullOrEmpty($GitHubToken)) {
            Write-Host ""
            Write-Host "        [!] Token GitHub requis pour l'upload automatique" -ForegroundColor Yellow
            Write-Host "        [i] Option 1: Definir `$env:GITHUB_TOKEN = 'ghp_...'" -ForegroundColor Cyan
            Write-Host "        [i] Option 2: Entrer le token ci-dessous (sera sauvegarde)" -ForegroundColor Cyan
            Write-Host ""
            $GitHubToken = Read-Host "        Entrez votre GitHub Personal Access Token"
            
            if (-not [string]::IsNullOrEmpty($GitHubToken)) {
                # Sauvegarder le token pour les prochaines fois
                $GitHubToken | Out-File $tokenFile -Encoding UTF8 -NoNewline
                Write-Host "        [+] Token sauvegarde dans .github-token" -ForegroundColor Green
                
                # Ajouter au .gitignore si pas deja present
                $gitignore = "$PSScriptRoot\.gitignore"
                if (Test-Path $gitignore) {
                    $content = Get-Content $gitignore -Raw
                    if ($content -notmatch "\.github-token") {
                        Add-Content $gitignore "`n.github-token"
                    }
                } else {
                    ".github-token" | Out-File $gitignore -Encoding UTF8
                }
            }
        }
        
        if ([string]::IsNullOrEmpty($GitHubToken)) {
            Write-Host "        [-] Token non fourni - Publication annulee" -ForegroundColor Red
        } else {
            # Determiner la version
            $version = if ($NewVersion -ne "") { $NewVersion } else { "1.0.0" }
            $tagName = if ($version.StartsWith("v")) { $version } else { "v$version" }
            
            # Headers pour l'API GitHub
            $headers = @{
                "Authorization" = "Bearer $GitHubToken"
                "Accept" = "application/vnd.github+json"
                "X-GitHub-Api-Version" = "2022-11-28"
            }
            
            $apiBase = "https://api.github.com/repos/$GitHubOwner/$GitHubRepo"
            
            try {
                # Verifier si la release existe deja
                Write-Host "        [>] Verification release existante..." -ForegroundColor Gray
                $releaseUrl = "$apiBase/releases/tags/$tagName"
                $existingRelease = $null
                
                try {
                    $existingRelease = Invoke-RestMethod -Uri $releaseUrl -Headers $headers -Method Get
                    Write-Host "        [i] Release $tagName trouvee (ID: $($existingRelease.id))" -ForegroundColor Cyan
                } catch {
                    if ($_.Exception.Response.StatusCode -ne 404) {
                        throw
                    }
                    Write-Host "        [i] Release $tagName n'existe pas encore" -ForegroundColor Gray
                }
                
                $releaseId = $null
                
                if ($existingRelease -and $NewVersion -eq "") {
                    # Utiliser la release existante
                    $releaseId = $existingRelease.id
                    Write-Host "        [>] Mise a jour de la release existante..." -ForegroundColor Cyan
                    
                    # Supprimer l'ancien asset s'il existe
                    $assetName = "XNRGYEngineeringAutomationToolsSetup.exe"
                    $existingAssets = $existingRelease.assets | Where-Object { $_.name -eq $assetName }
                    
                    foreach ($asset in $existingAssets) {
                        Write-Host "        [>] Suppression ancien fichier: $($asset.name)..." -ForegroundColor Gray
                        $deleteUrl = "$apiBase/releases/assets/$($asset.id)"
                        Invoke-RestMethod -Uri $deleteUrl -Headers $headers -Method Delete | Out-Null
                    }
                } else {
                    # Creer une nouvelle release
                    Write-Host "        [>] Creation nouvelle release $tagName..." -ForegroundColor Cyan
                    
                    $releaseBody = @{
                        tag_name = $tagName
                        target_commitish = "main"
                        name = "XEAT $version"
                        body = "## XNRGY Engineering Automation Tools $version`n`n### Installation`n1. Telecharger ``XNRGYEngineeringAutomationToolsSetup.exe```n2. Executer l'installateur`n3. Suivre les instructions`n`n### Build Info`n- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n- Configuration: $Configuration"
                        draft = $false
                        prerelease = $false
                    } | ConvertTo-Json
                    
                    $createUrl = "$apiBase/releases"
                    $newRelease = Invoke-RestMethod -Uri $createUrl -Headers $headers -Method Post -Body $releaseBody -ContentType "application/json"
                    $releaseId = $newRelease.id
                    Write-Host "        [+] Release creee (ID: $releaseId)" -ForegroundColor Green
                }
                
                # Uploader le fichier
                Write-Host "        [>] Upload du fichier..." -ForegroundColor Cyan
                $fileName = Split-Path $setupExeFullPath -Leaf
                $fileSize = (Get-Item $setupExeFullPath).Length
                $fileSizeMB = [math]::Round($fileSize / 1MB, 2)
                Write-Host "        [i] Fichier: $fileName ($fileSizeMB MB)" -ForegroundColor Gray
                
                $uploadUrl = "https://uploads.github.com/repos/$GitHubOwner/$GitHubRepo/releases/$releaseId/assets?name=$fileName"
                
                $uploadHeaders = @{
                    "Authorization" = "Bearer $GitHubToken"
                    "Accept" = "application/vnd.github+json"
                    "Content-Type" = "application/octet-stream"
                }
                
                $fileBytes = [System.IO.File]::ReadAllBytes($setupExeFullPath)
                $uploadResult = Invoke-RestMethod -Uri $uploadUrl -Headers $uploadHeaders -Method Post -Body $fileBytes
                
                Write-Host ""
                Write-Host "        [+] PUBLICATION REUSSIE!" -ForegroundColor Green
                Write-Host "        [i] URL: https://github.com/$GitHubOwner/$GitHubRepo/releases/tag/$tagName" -ForegroundColor Cyan
                Write-Host "        [i] Download: $($uploadResult.browser_download_url)" -ForegroundColor Cyan
                
            } catch {
                Write-Host "        [-] ERREUR: $($_.Exception.Message)" -ForegroundColor Red
                if ($_.Exception.Response) {
                    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $errorBody = $reader.ReadToEnd()
                    Write-Host "        [-] Details: $errorBody" -ForegroundColor Red
                }
            }
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# ETAPE FINALE: Lancer l'application (sauf BuildOnly ou InstallerOnly)
# ═══════════════════════════════════════════════════════════════════════════════
$lastStep = if ($Publish) { $TotalSteps } elseif ($Deploy) { $TotalSteps } elseif ($CreatePackage) { 7 } elseif ($WithInstaller -or $InstallerOnly) { 6 } else { 4 }

if (-not $BuildOnly -and -not $InstallerOnly -and -not $CreatePackage) {
    Write-Host ""
    Write-Host "  [$lastStep/$TotalSteps] Lancement de l'application..." -ForegroundColor Yellow
    try {                                                                                       
        Start-Process -FilePath $exePath
        Write-Host "        [+] Application lancee" -ForegroundColor Green
    } catch {
        Write-Host "        [-] ERREUR: Impossible de lancer l'application" -ForegroundColor Red
        Pop-Location
        exit 1
    }
} else {
    Write-Host ""
    Write-Host "  [$lastStep/$TotalSteps] Lancement de l'application..." -ForegroundColor DarkGray
    Write-Host "        - Ignore (mode BuildOnly/InstallerOnly)" -ForegroundColor DarkGray
}

# Restaurer le répertoire original
Pop-Location

# Fin
Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                          TERMINE                                   ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
