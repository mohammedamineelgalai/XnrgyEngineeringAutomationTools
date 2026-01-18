########################################################################
#  XNRGY Firebase Sync Script
#  Telecharge les donnees dynamiques de Firebase AVANT deploiement
#  Preserve: devices, users, auditLog, statistics
#  
#  Usage:
#    .\Sync-Firebase.ps1                # Sync complet
#    .\Sync-Firebase.ps1 -Download      # Telecharger uniquement
#    .\Sync-Firebase.ps1 -Merge         # Merge avec firebase-init.json
#    .\Sync-Firebase.ps1 -Backup        # Creer un backup horodate
########################################################################

param(
    [switch]$Download,
    [switch]$Merge,
    [switch]$Backup,
    [switch]$PreDeploy  # Mode pre-deploiement: Download + Merge automatique
)

$ErrorActionPreference = "Stop"

# Configuration
$FirebaseUrl = "https://xeat-remote-control-default-rtdb.firebaseio.com"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$FirebaseDir = "$ScriptDir\Firebase Realtime Database configuration"
$InitFile = "$FirebaseDir\firebase-init.json"
$BackupDir = "$FirebaseDir\backups"
$DownloadFile = "$FirebaseDir\firebase-live.json"

# Sections a preserver (donnees dynamiques)
$DynamicSections = @(
    "devices",
    "users", 
    "userActiveSessions",
    "auditLog",
    "statistics",
    "telemetryEvents",
    "errorReports",
    "broadcasts"  # Garder les broadcasts actifs
)

# Sections statiques (templates - ne pas ecraser)
$StaticSections = @(
    "appConfig",
    "welcomeMessages",
    "welcomeMessage",
    "audit",
    "commands",
    "departments",
    "deviceCommands",
    "featureFlags",
    "forceUpdate",
    "integrations",
    "killSwitch",
    "maintenance",
    "notifications",
    "roles",
    "scheduling",
    "security",
    "sites",
    "systemInfo",
    "telemetry",
    "updates"
)

function Write-Header {
    Write-Host ""
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host "  XNRGY Firebase Sync Tool" -ForegroundColor Cyan
    Write-Host "  Database: $FirebaseUrl" -ForegroundColor DarkGray
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Download-FirebaseData {
    Write-Host "[>] Telechargement des donnees Firebase..." -ForegroundColor Yellow
    
    try {
        $response = Invoke-RestMethod -Uri "$FirebaseUrl/.json" -Method Get -ContentType "application/json"
        
        if ($response) {
            # Convertir en JSON formatte
            $json = $response | ConvertTo-Json -Depth 20 -Compress:$false
            
            # Sauvegarder
            $json | Out-File -FilePath $DownloadFile -Encoding UTF8 -Force
            
            $fileSize = [math]::Round((Get-Item $DownloadFile).Length / 1KB, 2)
            Write-Host "[+] Donnees telechargees: $DownloadFile ($fileSize KB)" -ForegroundColor Green
            
            # Afficher les statistiques
            $deviceCount = ($response.devices.PSObject.Properties | Where-Object { $_.Name -ne "templateDevice" }).Count
            $userCount = ($response.users.PSObject.Properties | Where-Object { $_.Name -ne "placeholder" }).Count
            $logCount = ($response.auditLog.PSObject.Properties | Where-Object { $_.Name -ne "placeholder" }).Count
            $broadcastCount = ($response.broadcasts.PSObject.Properties | Where-Object { $_.Name -ne "placeholder" -and $_.Value.status.active -eq $true }).Count
            
            Write-Host ""
            Write-Host "    [i] Devices enregistres: $deviceCount" -ForegroundColor Cyan
            Write-Host "    [i] Utilisateurs: $userCount" -ForegroundColor Cyan
            Write-Host "    [i] Logs d'audit: $logCount" -ForegroundColor Cyan
            Write-Host "    [i] Broadcasts actifs: $broadcastCount" -ForegroundColor Cyan
            
            return $response
        }
    }
    catch {
        Write-Host "[-] Erreur telechargement: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Create-Backup {
    Write-Host "[>] Creation du backup..." -ForegroundColor Yellow
    
    # Creer dossier backups si necessaire
    if (-not (Test-Path $BackupDir)) {
        New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFile = "$BackupDir\firebase-backup_$timestamp.json"
    
    if (Test-Path $DownloadFile) {
        Copy-Item $DownloadFile -Destination $backupFile -Force
        Write-Host "[+] Backup cree: $backupFile" -ForegroundColor Green
    }
    elseif (Test-Path $InitFile) {
        Copy-Item $InitFile -Destination $backupFile -Force
        Write-Host "[+] Backup cree depuis init: $backupFile" -ForegroundColor Green
    }
    
    # Nettoyer les vieux backups (garder les 10 derniers)
    $backups = Get-ChildItem "$BackupDir\firebase-backup_*.json" | Sort-Object LastWriteTime -Descending
    if ($backups.Count -gt 10) {
        $backups | Select-Object -Skip 10 | Remove-Item -Force
        Write-Host "[i] Anciens backups nettoyes (garde les 10 derniers)" -ForegroundColor Gray
    }
}

function Merge-FirebaseData {
    param([object]$LiveData)
    
    Write-Host "[>] Fusion des donnees..." -ForegroundColor Yellow
    
    if (-not $LiveData) {
        Write-Host "[-] Pas de donnees live a fusionner" -ForegroundColor Red
        return
    }
    
    # Charger le fichier init template
    if (-not (Test-Path $InitFile)) {
        Write-Host "[-] Fichier init introuvable: $InitFile" -ForegroundColor Red
        return
    }
    
    $initContent = Get-Content $InitFile -Raw | ConvertFrom-Json
    
    # Creer un nouvel objet fusionne
    $merged = @{}
    
    # 1. Copier les sections statiques depuis init (templates)
    foreach ($section in $StaticSections) {
        if ($initContent.PSObject.Properties[$section]) {
            $merged[$section] = $initContent.$section
            Write-Host "    [i] Section statique: $section (depuis template)" -ForegroundColor Gray
        }
    }
    
    # 2. Copier les sections dynamiques depuis live (donnees reelles)
    foreach ($section in $DynamicSections) {
        if ($LiveData.PSObject.Properties[$section]) {
            $merged[$section] = $LiveData.$section
            $count = ($LiveData.$section.PSObject.Properties).Count
            Write-Host "    [+] Section dynamique: $section ($count entrees)" -ForegroundColor Cyan
        }
        elseif ($initContent.PSObject.Properties[$section]) {
            # Fallback au template si pas de donnees live
            $merged[$section] = $initContent.$section
            Write-Host "    [!] Section dynamique: $section (fallback template)" -ForegroundColor Yellow
        }
    }
    
    # 3. Sauvegarder le fichier fusionne
    $mergedJson = $merged | ConvertTo-Json -Depth 20 -Compress:$false
    $mergedJson | Out-File -FilePath $InitFile -Encoding UTF8 -Force
    
    $fileSize = [math]::Round((Get-Item $InitFile).Length / 1KB, 2)
    Write-Host ""
    Write-Host "[+] Fusion terminee: $InitFile ($fileSize KB)" -ForegroundColor Green
}

function Show-Summary {
    Write-Host ""
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host "  SYNC TERMINE" -ForegroundColor Green
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Fichiers:" -ForegroundColor White
    Write-Host "    - Live:   $DownloadFile" -ForegroundColor Gray
    Write-Host "    - Init:   $InitFile" -ForegroundColor Gray
    Write-Host "    - Backup: $BackupDir" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Prochaine etape:" -ForegroundColor White
    Write-Host "    .\build-and-run.ps1 -Deploy -BuildOnly" -ForegroundColor Yellow
    Write-Host ""
}

# ============================================================
# MAIN
# ============================================================

Write-Header

# Mode PreDeploy = Download + Backup + Merge
if ($PreDeploy) {
    $Download = $true
    $Backup = $true
    $Merge = $true
}

# Si aucun flag, faire tout par defaut
if (-not $Download -and -not $Backup -and -not $Merge) {
    $Download = $true
    $Backup = $true
    $Merge = $true
}

$liveData = $null

if ($Download) {
    $liveData = Download-FirebaseData
    Write-Host ""
}

if ($Backup) {
    Create-Backup
    Write-Host ""
}

if ($Merge) {
    if (-not $liveData -and (Test-Path $DownloadFile)) {
        $liveData = Get-Content $DownloadFile -Raw | ConvertFrom-Json
    }
    Merge-FirebaseData -LiveData $liveData
}

Show-Summary
