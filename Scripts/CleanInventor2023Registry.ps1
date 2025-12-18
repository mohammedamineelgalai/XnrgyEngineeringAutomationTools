# ═══════════════════════════════════════════════════════════════════════════════
# Script: CleanInventor2023Registry.ps1
# Description: Nettoie toutes les références Inventor 2023 du registre Windows
# Auteur: Smart Tools Amine - XNRGY
# Date: 2025-12-15
# ═══════════════════════════════════════════════════════════════════════════════
# ATTENTION: Ce script doit être exécuté en tant qu'Administrateur
# ═══════════════════════════════════════════════════════════════════════════════

#Requires -RunAsAdministrator

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  NETTOYAGE DES RÉFÉRENCES INVENTOR 2023 DU REGISTRE WINDOWS" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Fonction pour supprimer une clé de registre avec logging
function Remove-RegistryKeyIfExists {
    param (
        [string]$Path,
        [string]$Description
    )
    
    if (Test-Path $Path) {
        try {
            # Vérifier si la valeur contient "2023"
            $defaultValue = (Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue).'(default)'
            if ($defaultValue -and $defaultValue -like "*2023*") {
                Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                Write-Host "  [SUPPRIMÉ] $Description" -ForegroundColor Green
                Write-Host "             $Path" -ForegroundColor Gray
                return $true
            }
        }
        catch {
            Write-Host "  [ERREUR] Impossible de supprimer: $Path" -ForegroundColor Red
            Write-Host "           $($_.Exception.Message)" -ForegroundColor DarkRed
        }
    }
    return $false
}

# Fonction pour rechercher et lister les clés contenant "Inventor 2023"
function Find-Inventor2023Keys {
    param (
        [string]$RootPath,
        [string]$SearchPattern = "*2023*"
    )
    
    $foundKeys = @()
    
    try {
        $keys = Get-ChildItem -Path $RootPath -Recurse -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -like "*Inventor*" }
        
        foreach ($key in $keys) {
            try {
                $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                $defaultVal = $props.'(default)'
                
                if ($defaultVal -and $defaultVal -like "*Inventor 2023*") {
                    $foundKeys += @{
                        Path = $key.PSPath
                        Name = $key.Name
                        Value = $defaultVal
                    }
                }
            }
            catch { }
        }
    }
    catch { }
    
    return $foundKeys
}

# ═══════════════════════════════════════════════════════════════════════════════
# ÉTAPE 1: Lister les CLSIDs Inventor qui pointent vers 2023
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "ÉTAPE 1: Recherche des CLSIDs pointant vers Inventor 2023..." -ForegroundColor Yellow
Write-Host "─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

$clsidsToClean = @()

# CLSIDs connus d'Inventor
$knownCLSIDs = @(
    @{ CLSID = "{C343ED84-A129-11d3-B799-0060B0F159EF}"; Name = "Inventor.ApprenticeServer" },
    @{ CLSID = "{B6B5DC40-96E3-11d2-B774-0060B0F159EF}"; Name = "Inventor.Application" },
    @{ CLSID = "{27DA820F-F161-4F7F-B5A3-7E0E2F9C0E9C}"; Name = "Inventor.InventorServer" }
)

foreach ($item in $knownCLSIDs) {
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$($item.CLSID)"
    
    # Vérifier InprocServer32
    $inprocPath = "$clsidPath\InprocServer32"
    if (Test-Path $inprocPath) {
        $value = (Get-ItemProperty -Path $inprocPath -ErrorAction SilentlyContinue).'(default)'
        if ($value -and $value -like "*Inventor 2023*") {
            Write-Host "  [TROUVÉ] $($item.Name) (InprocServer32)" -ForegroundColor Magenta
            Write-Host "           CLSID: $($item.CLSID)" -ForegroundColor Gray
            Write-Host "           Chemin: $value" -ForegroundColor Gray
            $clsidsToClean += @{ Path = $inprocPath; Name = "$($item.Name) InprocServer32"; CLSID = $item.CLSID }
        }
    }
    
    # Vérifier LocalServer32
    $localPath = "$clsidPath\LocalServer32"
    if (Test-Path $localPath) {
        $value = (Get-ItemProperty -Path $localPath -ErrorAction SilentlyContinue).'(default)'
        if ($value -and $value -like "*Inventor 2023*") {
            Write-Host "  [TROUVÉ] $($item.Name) (LocalServer32)" -ForegroundColor Magenta
            Write-Host "           CLSID: $($item.CLSID)" -ForegroundColor Gray
            Write-Host "           Chemin: $value" -ForegroundColor Gray
            $clsidsToClean += @{ Path = $localPath; Name = "$($item.Name) LocalServer32"; CLSID = $item.CLSID }
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# ÉTAPE 2: Recherche étendue dans le registre
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "ÉTAPE 2: Recherche étendue des références Inventor 2023..." -ForegroundColor Yellow
Write-Host "─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

$searchPaths = @(
    "HKLM:\SOFTWARE\Classes\CLSID",
    "HKLM:\SOFTWARE\Classes\Inventor.ApprenticeServer",
    "HKLM:\SOFTWARE\Classes\Inventor.ApprenticeServer.1",
    "HKLM:\SOFTWARE\Classes\Inventor.Application",
    "HKLM:\SOFTWARE\Classes\Inventor.Application.1",
    "HKLM:\SOFTWARE\Classes\Inventor.InventorServer",
    "HKLM:\SOFTWARE\Classes\Inventor.InventorServer.1",
    "HKLM:\SOFTWARE\Autodesk\Inventor",
    "HKCU:\SOFTWARE\Autodesk\Inventor",
    "HKLM:\SOFTWARE\WOW6432Node\Classes\CLSID"
)

$allFound = @()

foreach ($searchPath in $searchPaths) {
    if (Test-Path $searchPath) {
        Write-Host "  Recherche dans: $searchPath" -ForegroundColor DarkGray
        
        try {
            Get-ChildItem -Path $searchPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
                    foreach ($prop in $props.PSObject.Properties) {
                        if ($prop.Value -and $prop.Value -is [string] -and $prop.Value -like "*Inventor 2023*") {
                            $allFound += @{
                                Path = $_.PSPath
                                Property = $prop.Name
                                Value = $prop.Value
                            }
                        }
                    }
                }
                catch { }
            }
        }
        catch { }
    }
}

Write-Host ""
Write-Host "  Trouvé $($allFound.Count) références à Inventor 2023" -ForegroundColor Cyan

# ═══════════════════════════════════════════════════════════════════════════════
# ÉTAPE 3: Afficher le résumé avant suppression
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "ÉTAPE 3: Résumé des éléments à nettoyer" -ForegroundColor Yellow
Write-Host "─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

if ($allFound.Count -eq 0) {
    Write-Host "  ✅ Aucune référence Inventor 2023 trouvée dans le registre!" -ForegroundColor Green
    Write-Host ""
    exit 0
}

Write-Host ""
Write-Host "  Les références suivantes seront nettoyées:" -ForegroundColor White
Write-Host ""

$index = 1
foreach ($item in $allFound) {
    Write-Host "  [$index] $($item.Path)" -ForegroundColor White
    Write-Host "      Propriété: $($item.Property)" -ForegroundColor Gray
    Write-Host "      Valeur: $($item.Value)" -ForegroundColor DarkGray
    Write-Host ""
    $index++
}

# ═══════════════════════════════════════════════════════════════════════════════
# ÉTAPE 4: Confirmation et suppression
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
$confirmation = Read-Host "Voulez-vous supprimer ces $($allFound.Count) références? (O/N)"

if ($confirmation -eq "O" -or $confirmation -eq "o" -or $confirmation -eq "Y" -or $confirmation -eq "y") {
    Write-Host ""
    Write-Host "ÉTAPE 4: Suppression des références..." -ForegroundColor Yellow
    Write-Host "─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    
    $deleted = 0
    $errors = 0
    
    foreach ($item in $allFound) {
        try {
            $regPath = $item.Path -replace "Microsoft.PowerShell.Core\\Registry::", ""
            
            # Si c'est une propriété spécifique (pas default), la supprimer
            if ($item.Property -ne "(default)" -and $item.Property -ne "PSPath") {
                Remove-ItemProperty -Path $item.Path -Name $item.Property -Force -ErrorAction Stop
                Write-Host "  [OK] Propriété supprimée: $($item.Property)" -ForegroundColor Green
            }
            else {
                # Supprimer la clé entière si c'est la valeur par défaut
                # Mais attention, ne pas supprimer les clés parentes importantes
                $keyName = Split-Path $item.Path -Leaf
                if ($keyName -like "*2023*" -or $keyName -eq "InprocServer32" -or $keyName -eq "LocalServer32") {
                    # Ne pas supprimer InprocServer32/LocalServer32, juste mettre à jour la valeur
                    Write-Host "  [INFO] Clé système - mise à jour manuelle requise: $regPath" -ForegroundColor Yellow
                }
            }
            $deleted++
        }
        catch {
            Write-Host "  [ERREUR] $($item.Path): $($_.Exception.Message)" -ForegroundColor Red
            $errors++
        }
    }
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  RÉSULTAT: $deleted traité(s), $errors erreur(s)" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    
    # ═══════════════════════════════════════════════════════════════════════════════
    # ÉTAPE 5: Réparer les CLSIDs en pointant vers Inventor 2026
    # ═══════════════════════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "ÉTAPE 5: Réparation des CLSIDs pour Inventor 2026..." -ForegroundColor Yellow
    Write-Host "─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    
    $inventor2026Path = "C:\Program Files\Autodesk\Inventor 2026\Bin"
    
    if (Test-Path $inventor2026Path) {
        # Réparer ApprenticeServer
        $apprenticeDll = "$inventor2026Path\RxApprenticeServer.dll"
        if (Test-Path $apprenticeDll) {
            $apprenticeClsid = "HKLM:\SOFTWARE\Classes\CLSID\{C343ED84-A129-11d3-B799-0060B0F159EF}\InprocServer32"
            if (Test-Path $apprenticeClsid) {
                try {
                    Set-ItemProperty -Path $apprenticeClsid -Name "(default)" -Value $apprenticeDll -Force
                    Write-Host "  [RÉPARÉ] ApprenticeServer → $apprenticeDll" -ForegroundColor Green
                }
                catch {
                    Write-Host "  [ERREUR] ApprenticeServer: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "  [INFO] RxApprenticeServer.dll non trouvé dans Inventor 2026" -ForegroundColor Yellow
        }
        
        # Réparer Inventor.Application (si nécessaire)
        $inventorExe = "$inventor2026Path\Inventor.exe"
        if (Test-Path $inventorExe) {
            $appClsid = "HKLM:\SOFTWARE\Classes\CLSID\{B6B5DC40-96E3-11d2-B774-0060B0F159EF}\LocalServer32"
            if (Test-Path $appClsid) {
                $currentValue = (Get-ItemProperty -Path $appClsid -ErrorAction SilentlyContinue).'(default)'
                if ($currentValue -like "*2023*") {
                    try {
                        Set-ItemProperty -Path $appClsid -Name "(default)" -Value "$inventorExe /Automation" -Force
                        Write-Host "  [RÉPARÉ] Inventor.Application → $inventorExe" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "  [ERREUR] Inventor.Application: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        }
    }
    else {
        Write-Host "  [ERREUR] Inventor 2026 non trouvé dans: $inventor2026Path" -ForegroundColor Red
    }
}
else {
    Write-Host ""
    Write-Host "  Opération annulée." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  NETTOYAGE TERMINÉ" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  💡 CONSEIL: Redémarrez votre ordinateur pour que les changements" -ForegroundColor Yellow
Write-Host "              prennent effet complètement." -ForegroundColor Yellow
Write-Host ""
