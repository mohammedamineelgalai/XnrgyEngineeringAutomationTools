# Script agressif pour forcer la mise a jour du nom dans la barre des taches
# Ce script nettoie tous les caches possibles et force Windows a recharger les informations

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Mise a jour forcee du nom dans la barre des taches" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Etape 1: Arreter l'explorateur Windows
Write-Host "[1/6] Arret de l'explorateur Windows..." -ForegroundColor Yellow
Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Write-Host "    OK - Explorateur arrete" -ForegroundColor Green

# Etape 2: Nettoyer le cache d'icones
Write-Host "[2/6] Nettoyage du cache d'icones..." -ForegroundColor Yellow
$iconCachePaths = @(
    "$env:LOCALAPPDATA\IconCache.db",
    "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*.db"
)
foreach ($path in $iconCachePaths) {
    if (Test-Path $path) {
        Remove-Item $path -Force -ErrorAction SilentlyContinue
        Write-Host "    OK - Cache supprime: $path" -ForegroundColor Green
    }
}
Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" -Filter "iconcache*.db" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

# Etape 3: Nettoyer le cache de la barre des taches
Write-Host "[3/6] Nettoyage du cache de la barre des taches..." -ForegroundColor Yellow
$taskbarCache = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"
if (Test-Path $taskbarCache) {
    $files = Get-ChildItem $taskbarCache -Filter "*.automaticDestinations-ms" -ErrorAction SilentlyContinue
    $count = ($files | Measure-Object).Count
    $files | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Host "    OK - $count fichier(s) supprime(s)" -ForegroundColor Green
}

# Etape 4: Nettoyer le cache des raccourcis
Write-Host "[4/6] Nettoyage du cache des raccourcis..." -ForegroundColor Yellow
$shortcutCache = "$env:LOCALAPPDATA\Microsoft\Windows\Shell"
if (Test-Path "$shortcutCache\Bags") {
    Remove-Item "$shortcutCache\Bags" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "    OK - Cache des raccourcis supprime" -ForegroundColor Green
}

# Etape 5: Nettoyer le cache de la base de registre (sans modifier, juste info)
Write-Host "[5/6] Verification du registre..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\TrayNotify"
if (Test-Path $regPath) {
    Write-Host "    INFO - Cle de registre trouvee (non modifiee pour securite)" -ForegroundColor Yellow
    Write-Host "    NOTE: Windows regenerera cette cle au redemarrage de l'explorateur" -ForegroundColor Yellow
}

# Etape 6: Redemarrer l'explorateur Windows
Write-Host "[6/6] Redemarrage de l'explorateur Windows..." -ForegroundColor Yellow
Start-Process "explorer.exe"
Start-Sleep -Seconds 3
Write-Host "    OK - Explorateur redemarre" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Actions finales REQUISES:" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. IMPORTANT: Desepinglez l'application de la barre des taches" -ForegroundColor Red
Write-Host "   - Clic droit sur l'icone -> Desepingler de la barre des taches" -ForegroundColor White
Write-Host ""
Write-Host "2. Fermez TOUTES les instances de l'application si elle est ouverte" -ForegroundColor Yellow
Write-Host ""
Write-Host "3. Lancez l'application depuis:" -ForegroundColor Cyan
Write-Host "   bin\Release\XnrgyEngineeringAutomationTools.exe" -ForegroundColor White
Write-Host ""
Write-Host "4. Reepinglez l'application a la barre des taches" -ForegroundColor Cyan
Write-Host "   - Clic droit sur l'icone -> Epingler a la barre des taches" -ForegroundColor White
Write-Host ""
Write-Host "5. Si le probleme persiste, redemarrez Windows completement" -ForegroundColor Yellow
Write-Host ""
Write-Host "Le nouveau nom devrait maintenant apparaitre!" -ForegroundColor Green
Write-Host ""




