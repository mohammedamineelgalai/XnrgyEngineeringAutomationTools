<#
.SYNOPSIS
    Upload COMPLET de C:\Vault vers Vault PROD_XNRGY (Production)
    
.DESCRIPTION
    Ce script uploade TOUT le contenu de C:\Vault vers $/ dans Vault Production.
    - Preserve la structure complete des dossiers (C:\Vault\X\Y\Z -> $/X/Y/Z)
    - Exclut les fichiers temporaires (.bak, .old, ~$, etc.)
    - Cree automatiquement les dossiers Vault si necessaires
    - Gere les fichiers IPJ separement
    - Premier upload/checkin des fichiers propres
    
.PARAMETER LocalRoot
    Racine locale a uploader (defaut: C:\Vault)
    
.PARAMETER UploadAll
    Switch pour uploader sans confirmation de chaque fichier
    
.EXAMPLE
    .\Upload-ToVaultProd.ps1
    # Uploade tout C:\Vault vers $/
    
.EXAMPLE
    .\Upload-ToVaultProd.ps1 -LocalRoot "C:\Vault\Engineering\Library"
    # Uploade seulement Engineering\Library vers $/Engineering/Library
    
.NOTES
    Auteur: Mohammed Amine Elgalai - XNRGY Climate Systems
    Date: 2025-12-30
    PRODUCTION VAULT: PROD_XNRGY sur VAULTPOC
    
    STRUCTURE ATTENDUE:
    C:\Vault\Engineering\Projects\[PROJECT]\[REF]\[MODULE]\...
    C:\Vault\Engineering\Library\...
    C:\Vault\Engineering\Content Center Files\...
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$LocalRoot = "C:\Vault",
    
    [Parameter(Mandatory=$false)]
    [switch]$UploadAll
)

# ============================================================================
# Configuration Vault PRODUCTION
# ============================================================================
$VaultServer = "VAULTPOC"
$VaultName = "PROD_XNRGY"
$VaultUser = "mohammedamine.elgalai"
$VaultPassword = "Vtr8aPz2*"

# Racine Locale = C:\Vault -> Racine Vault = $/
$LocalVaultRoot = "C:\Vault"

# Extensions a EXCLURE
$ExcludedExtensions = @(".bak", ".old", ".tmp", ".log", ".lck", ".dwl", ".dwl2", ".idw.bak", ".v")

# Prefixes a EXCLURE
$ExcludedPrefixes = @("~`$", "._", "Backup_", ".~")

# Dossiers a EXCLURE (ne pas uploader ces dossiers et leur contenu)
$ExcludedFolders = @(
    "OldVersions",
    "Backup", 
    ".vault", 
    ".git", 
    ".vs", 
    "obj", 
    "bin",
    "Workspace",
    "vltcache"
)

# Fichiers specifiques a exclure
$ExcludedFileNames = @(
    "desktop.ini",
    "Thumbs.db",
    ".DS_Store"
)

# Extensions Inventor (traitement special pour les IPJ)
$InventorExtensions = @(".ipt", ".iam", ".idw", ".dwg", ".ipn", ".ide", ".ipj")

# ============================================================================
# Variables globales
# ============================================================================
$ErrorActionPreference = "Continue"
$script:FilesTotal = 0
$script:FilesUploaded = 0
$script:FilesFailed = 0
$script:FilesSkipped = 0
$script:FoldersCreated = 0
$script:FoldersExisting = 0
$script:FailedFiles = @()

# Assemblies Vault SDK
$VaultSDKPath = "C:\Program Files\Autodesk\Vault Professional 2026\SDK"

# Cache des dossiers Vault (evite de recreer)
$script:VaultFolderCache = @{}

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
        "DEBUG"   { "DarkGray" }
        default   { "White" }
    }
    
    $prefix = switch ($Level) {
        "SUCCESS" { "[+]" }
        "ERROR"   { "[-]" }
        "WARNING" { "[!]" }
        "INFO"    { "[i]" }
        "DEBUG"   { "[>]" }
        default   { "[>]" }
    }
    
    Write-Host "[$timestamp] $prefix $Message" -ForegroundColor $color
    
    # Log to file
    $logFile = Join-Path $PSScriptRoot "Upload-ToVaultProd_$(Get-Date -Format 'yyyyMMdd').log"
    "[$timestamp] [$Level] $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

function Load-VaultSDK {
    Write-Log "Chargement du Vault SDK..." "INFO"
    
    try {
        # Chemins possibles pour les assemblies Vault
        $sdkPaths = @(
            "C:\Program Files\Autodesk\Autodesk Vault 2026 SDK\bin\x64",
            "C:\Program Files\Autodesk\Vault Client 2026\Explorer",
            "C:\Program Files\Autodesk\Vault Professional 2026\Explorer",
            $VaultSDKPath
        )
        
        # Trouver le premier chemin valide
        $validPath = $null
        foreach ($path in $sdkPaths) {
            $testFile = Join-Path $path "Autodesk.DataManagement.Client.Framework.Vault.dll"
            if (Test-Path $testFile) {
                $validPath = $path
                Write-Log "SDK trouve dans: $path" "DEBUG"
                break
            }
        }
        
        if ($null -eq $validPath) {
            Write-Log "Vault SDK introuvable dans les chemins connus" "ERROR"
            return $false
        }
        
        # Charger les assemblies Vault dans l'ordre correct
        $assemblies = @(
            "Autodesk.Connectivity.WebServices.dll",
            "Autodesk.Connectivity.WebServicesTools.dll", 
            "Autodesk.DataManagement.Client.Framework.dll",
            "Autodesk.DataManagement.Client.Framework.Vault.dll",
            "Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
        )
        
        foreach ($asm in $assemblies) {
            $asmPath = Join-Path $validPath $asm
            if (Test-Path $asmPath) {
                [System.Reflection.Assembly]::LoadFrom($asmPath) | Out-Null
                Write-Log "  Charge: $asm" "DEBUG"
            }
            else {
                Write-Log "  Non trouve: $asm" "WARNING"
            }
        }
        
        Write-Log "Vault SDK charge" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Erreur chargement SDK: $_" "ERROR"
        return $false
    }
}

function Connect-ToVault {
    Write-Log "Connexion a Vault $VaultName sur $VaultServer..." "INFO"
    
    try {
        # Utiliser VDF pour la connexion
        $serverIdentifier = New-Object Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections.ServerIdentifier($VaultServer)
        
        $authFlags = [Autodesk.Connectivity.WebServicesTools.AuthenticationFlags]::Standard
        
        $connection = [Autodesk.DataManagement.Client.Framework.Vault.Library.ConnectionManager]::LogIn(
            $serverIdentifier,
            $VaultName,
            $VaultUser,
            $VaultPassword,
            $authFlags,
            $null
        )
        
        if ($null -ne $connection) {
            Write-Log "Connecte a $VaultName (User: $VaultUser)" "SUCCESS"
            return $connection
        }
        else {
            Write-Log "Connexion retourne null" "ERROR"
            return $null
        }
    }
    catch {
        Write-Log "Erreur connexion Vault: $_" "ERROR"
        return $null
    }
}

function Should-ExcludeFile {
    param([string]$FilePath)
    
    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    # Verifier nom de fichier exclu
    foreach ($name in $ExcludedFileNames) {
        if ($fileName -eq $name) {
            return $true
        }
    }
    
    # Verifier extension exclue
    foreach ($ext in $ExcludedExtensions) {
        if ($fileName.EndsWith($ext, [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    
    # Verifier prefixe exclu
    foreach ($prefix in $ExcludedPrefixes) {
        if ($fileName.StartsWith($prefix)) {
            return $true
        }
    }
    
    # Verifier si dans dossier exclu
    foreach ($folder in $ExcludedFolders) {
        if ($FilePath -like "*\$folder\*" -or $FilePath -like "*\$folder") {
            return $true
        }
    }
    
    return $false
}

function Get-VaultPathFromLocal {
    param([string]$LocalPath)
    
    # Convertir C:\Vault\X\Y\Z en $/X/Y/Z
    if ($LocalPath.StartsWith($LocalVaultRoot, [StringComparison]::OrdinalIgnoreCase)) {
        $relativePath = $LocalPath.Substring($LocalVaultRoot.Length).TrimStart('\')
        if ([string]::IsNullOrEmpty($relativePath)) {
            return "`$"
        }
        return "`$/" + $relativePath.Replace("\", "/")
    }
    
    # Cas ou LocalRoot n'est pas C:\Vault
    if ($LocalPath.StartsWith($LocalRoot, [StringComparison]::OrdinalIgnoreCase)) {
        # Calculer le chemin relatif depuis LocalRoot
        $relativePath = $LocalPath.Substring($LocalRoot.Length).TrimStart('\')
        
        # Si LocalRoot est C:\Vault\Engineering\Library, le chemin Vault est $/Engineering/Library/...
        $vaultBase = Get-VaultPathFromLocal $LocalRoot
        
        if ([string]::IsNullOrEmpty($relativePath)) {
            return $vaultBase
        }
        return $vaultBase + "/" + $relativePath.Replace("\", "/")
    }
    
    return $null
}

function Ensure-VaultFolder {
    param(
        [System.Object]$Connection,
        [string]$VaultPath
    )
    
    # Verifier le cache
    if ($script:VaultFolderCache.ContainsKey($VaultPath)) {
        return $script:VaultFolderCache[$VaultPath]
    }
    
    try {
        # Verifier si le dossier existe
        $folder = $Connection.WebServiceManager.DocumentService.GetFolderByPath($VaultPath)
        $script:VaultFolderCache[$VaultPath] = $folder
        $script:FoldersExisting++
        return $folder
    }
    catch {
        # Dossier n'existe pas - le creer
        try {
            # Trouver le parent
            $lastSlash = $VaultPath.LastIndexOf('/')
            if ($lastSlash -le 0) {
                # Racine $/ - ne peut pas creer
                try {
                    $rootFolder = $Connection.WebServiceManager.DocumentService.GetFolderByPath("`$")
                    $script:VaultFolderCache["`$"] = $rootFolder
                    return $rootFolder
                }
                catch {
                    return $null
                }
            }
            
            $parentPath = $VaultPath.Substring(0, $lastSlash)
            $folderName = $VaultPath.Substring($lastSlash + 1)
            
            # Creer le parent recursivement
            $parentFolder = Ensure-VaultFolder -Connection $Connection -VaultPath $parentPath
            
            if ($null -eq $parentFolder) {
                Write-Log "Parent introuvable pour: $VaultPath" "ERROR"
                return $null
            }
            
            # Creer le dossier
            Write-Log "    [>] Creation dossier: $VaultPath" "DEBUG"
            $newFolder = $Connection.WebServiceManager.DocumentService.AddFolder($folderName, $parentFolder.Id, $false)
            $script:FoldersCreated++
            $script:VaultFolderCache[$VaultPath] = $newFolder
            return $newFolder
        }
        catch {
            # Peut-etre que le dossier a ete cree entre temps
            try {
                $folder = $Connection.WebServiceManager.DocumentService.GetFolderByPath($VaultPath)
                $script:VaultFolderCache[$VaultPath] = $folder
                return $folder
            }
            catch {
                Write-Log "Erreur creation dossier $VaultPath : $_" "ERROR"
                return $null
            }
        }
    }
}

function Upload-FileToVault {
    param(
        [System.Object]$Connection,
        [string]$LocalFilePath,
        [string]$VaultFolderPath
    )
    
    $fileName = [System.IO.Path]::GetFileName($LocalFilePath)
    
    try {
        # Obtenir ou creer le dossier Vault
        $folder = Ensure-VaultFolder -Connection $Connection -VaultPath $VaultFolderPath
        
        if ($null -eq $folder) {
            Write-Log "    [-] Dossier Vault introuvable: $VaultFolderPath" "ERROR"
            $script:FailedFiles += @{Path=$LocalFilePath; Error="Dossier introuvable"}
            return $false
        }
        
        # Verifier si le fichier existe deja
        try {
            $existingFiles = $Connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId($folder.Id, $false)
            $existingFile = $existingFiles | Where-Object { $_.Name -eq $fileName }
            
            if ($null -ne $existingFile) {
                $script:FilesSkipped++
                return $true  # Skip silencieusement
            }
        }
        catch { }
        
        # Lire le fichier
        $fileInfo = Get-Item $LocalFilePath
        
        # Upload via FileManager
        $stream = [System.IO.File]::OpenRead($LocalFilePath)
        
        try {
            $vdfFolder = New-Object Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.Folder($Connection, $folder)
            
            $result = $Connection.FileManager.AddFile(
                $vdfFolder,
                $fileName,
                "Upload initial - XNRGY Template",
                $fileInfo.LastWriteTimeUtc,
                $null,  # associations
                $null,  # bom
                [Autodesk.Connectivity.WebServices.FileClassification]::None,
                $false, # hidden
                $stream
            )
            
            if ($null -ne $result) {
                $script:FilesUploaded++
                return $true
            }
        }
        finally {
            $stream.Close()
            $stream.Dispose()
        }
        
        Write-Log "    [-] Upload echoue (resultat null)" "ERROR"
        $script:FailedFiles += @{Path=$LocalFilePath; Error="Resultat null"}
        return $false
    }
    catch {
        $errorMsg = $_.Exception.Message
        # Verifier si c'est juste un fichier existant
        if ($errorMsg -like "*already exists*" -or $errorMsg -like "*1008*") {
            $script:FilesSkipped++
            return $true
        }
        
        Write-Log "    [-] ERREUR: $errorMsg" "ERROR"
        $script:FailedFiles += @{Path=$LocalFilePath; Error=$errorMsg}
        return $false
    }
}

# ============================================================================
# Main - Upload C:\Vault -> $/
# ============================================================================

Clear-Host
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  UPLOAD COMPLET VERS VAULT PRODUCTION - XNRGY" -ForegroundColor Cyan
Write-Host "  Vault: $VaultName sur $VaultServer" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Source: C:\Vault -> Destination: `$/" -ForegroundColor White
Write-Host ""

# Verifier si un sous-dossier specifique est demande
if ($LocalRoot -ne "C:\Vault" -and $LocalRoot -ne $LocalVaultRoot) {
    Write-Host "  Mode: Upload partiel" -ForegroundColor Yellow
    Write-Host "  Dossier: $LocalRoot" -ForegroundColor White
}
else {
    Write-Host "  Mode: Upload COMPLET de C:\Vault" -ForegroundColor Green
    $LocalRoot = $LocalVaultRoot
}

Write-Host ""

# Verifier que le dossier local existe
if (-not (Test-Path $LocalRoot)) {
    Write-Log "Dossier local introuvable: $LocalRoot" "ERROR"
    Read-Host "Appuyez sur Entree pour fermer"
    exit 1
}

# Scanner les fichiers
Write-Log "Scan des fichiers dans $LocalRoot..." "INFO"
Write-Log "Exclusions: OldVersions, Backup, .vault, .git, .bak, .tmp, .lck..." "DEBUG"
Write-Host ""

$allFiles = Get-ChildItem -Path $LocalRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
    -not (Should-ExcludeFile -FilePath $_.FullName)
}

$script:FilesTotal = $allFiles.Count

# Compter les dossiers uniques
$uniqueFolders = $allFiles | ForEach-Object { [System.IO.Path]::GetDirectoryName($_.FullName) } | Select-Object -Unique
$totalFolders = $uniqueFolders.Count

Write-Log "Fichiers a traiter: $script:FilesTotal" "INFO"
Write-Log "Dossiers uniques: $totalFolders" "INFO"
Write-Host ""

# Afficher un apercu par type
$byExtension = $allFiles | Group-Object { [System.IO.Path]::GetExtension($_.Name).ToLower() } | Sort-Object Count -Descending | Select-Object -First 10
Write-Host "  Apercu par extension:" -ForegroundColor DarkGray
foreach ($grp in $byExtension) {
    Write-Host "    $($grp.Name): $($grp.Count) fichiers" -ForegroundColor DarkGray
}
Write-Host ""

if ($script:FilesTotal -eq 0) {
    Write-Log "Aucun fichier a uploader" "WARNING"
    Read-Host "Appuyez sur Entree pour fermer"
    exit
}

# Confirmation IMPORTANTE pour PROD
Write-Host "================================================================" -ForegroundColor Red
Write-Host "  [!] ATTENTION: VAULT PRODUCTION ($VaultName)" -ForegroundColor Red
Write-Host "================================================================" -ForegroundColor Red
Write-Host ""
Write-Host "  Vous allez uploader $script:FilesTotal fichiers vers PRODUCTION!" -ForegroundColor Yellow
Write-Host "  Structure: C:\Vault\* -> `$/\*" -ForegroundColor Yellow
Write-Host ""
$confirm = Read-Host "Tapez 'PROD' pour confirmer"

if ($confirm -ne "PROD") {
    Write-Log "Operation annulee (confirmation incorrecte)" "WARNING"
    Read-Host "Appuyez sur Entree pour fermer"
    exit
}

Write-Host ""

# Charger le SDK
if (-not (Load-VaultSDK)) {
    Write-Log "Impossible de charger le Vault SDK" "ERROR"
    Read-Host "Appuyez sur Entree pour fermer"
    exit 1
}

# Connexion au Vault
$connection = Connect-ToVault
if ($null -eq $connection) {
    Write-Log "Impossible de se connecter au Vault" "ERROR"
    Read-Host "Appuyez sur Entree pour fermer"
    exit 1
}

Write-Host ""
Write-Log "Debut de l'upload vers $VaultName..." "INFO"
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

$counter = 0
$startTime = Get-Date
$lastProgressTime = $startTime

foreach ($file in $allFiles) {
    $counter++
    
    # Calculer le chemin Vault
    $vaultFilePath = Get-VaultPathFromLocal $file.FullName
    $vaultFolderPath = Get-VaultPathFromLocal $file.DirectoryName
    
    # Afficher progression tous les 50 fichiers ou toutes les 10 secondes
    $now = Get-Date
    if ($counter % 50 -eq 0 -or ($now - $lastProgressTime).TotalSeconds -ge 10) {
        $elapsed = ($now - $startTime).TotalMinutes
        $rate = $counter / [Math]::Max($elapsed, 0.01)
        $remaining = ($script:FilesTotal - $counter) / [Math]::Max($rate, 1)
        
        Write-Host ""
        Write-Log "Progression: $counter/$script:FilesTotal ($([Math]::Round($counter * 100 / $script:FilesTotal))%)" "INFO"
        Write-Log "  Uploades: $script:FilesUploaded | Ignores: $script:FilesSkipped | Echecs: $script:FilesFailed" "INFO"
        Write-Log "  Vitesse: $([Math]::Round($rate, 1)) fichiers/min | Reste: ~$([Math]::Round($remaining, 1)) min" "INFO"
        Write-Host ""
        
        $lastProgressTime = $now
    }
    
    # Upload
    $success = Upload-FileToVault -Connection $connection -LocalFilePath $file.FullName -VaultFolderPath $vaultFolderPath
    
    if (-not $success) {
        $script:FilesFailed++
    }
    
    # Afficher seulement les erreurs
    if (-not $success -and $script:FilesFailed -le 20) {
        Write-Log "  [-] $($file.Name) -> $vaultFolderPath" "ERROR"
    }
}

# Deconnexion
try {
    [Autodesk.DataManagement.Client.Framework.Vault.Library.ConnectionManager]::LogOut($connection)
    Write-Log "Deconnecte du Vault" "INFO"
}
catch { }

# Temps total
$totalTime = (Get-Date) - $startTime

# Resume
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  RESUME UPLOAD VERS $VaultName" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Log "Temps total: $([Math]::Round($totalTime.TotalMinutes, 1)) minutes" "INFO"
Write-Log "Total fichiers: $script:FilesTotal" "INFO"
Write-Log "Uploades avec succes: $script:FilesUploaded" "SUCCESS"
Write-Log "Ignores (deja existants): $script:FilesSkipped" "WARNING"
Write-Log "Echecs: $script:FilesFailed" "ERROR"
Write-Log "Dossiers crees: $script:FoldersCreated" "INFO"
Write-Log "Dossiers existants: $script:FoldersExisting" "DEBUG"
Write-Host ""

# Lister les echecs si presents
if ($script:FailedFiles.Count -gt 0) {
    Write-Host "Fichiers en echec:" -ForegroundColor Red
    $failLogPath = Join-Path $PSScriptRoot "Upload-FailedFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    foreach ($failed in $script:FailedFiles | Select-Object -First 50) {
        Write-Host "  [-] $($failed.Path)" -ForegroundColor Red
        "[-] $($failed.Path) - $($failed.Error)" | Out-File -FilePath $failLogPath -Append -Encoding UTF8
    }
    
    if ($script:FailedFiles.Count -gt 50) {
        Write-Host "  ... et $($script:FailedFiles.Count - 50) autres" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Log "Liste complete des echecs sauvegardee: $failLogPath" "INFO"
}

Write-Host "================================================================" -ForegroundColor Cyan
if ($script:FilesFailed -eq 0) {
    Write-Host "  [+] UPLOAD TERMINE AVEC SUCCES!" -ForegroundColor Green
}
else {
    Write-Host "  [!] UPLOAD TERMINE AVEC $script:FilesFailed ERREURS" -ForegroundColor Yellow
}
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Read-Host "Appuyez sur Entree pour fermer"
