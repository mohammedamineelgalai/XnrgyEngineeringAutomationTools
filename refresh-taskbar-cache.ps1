# Script pour rafraîchir le cache de la barre des tâches Windows
# Ce script nettoie le cache d'icônes et redémarre l'explorateur Windows

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Rafraîchissement du cache de la barre des tâches" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Étape 1: Arrêter l'explorateur Windows
Write-Host "[1/4] Arret de l'explorateur Windows..." -ForegroundColor Yellow
Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Write-Host "    OK - Explorateur arrete" -ForegroundColor Green

# Étape 2: Nettoyer le cache d'icônes
Write-Host "[2/4] Nettoyage du cache d'icônes..." -ForegroundColor Yellow
$iconCachePath = "$env:LOCALAPPDATA\IconCache.db"
if (Test-Path $iconCachePath) {
    Remove-Item $iconCachePath -Force -ErrorAction SilentlyContinue
    Write-Host "    OK - Cache d'icônes supprime" -ForegroundColor Green
} else {
    Write-Host "    INFO - Cache d'icônes non trouve" -ForegroundColor Yellow
}

# Nettoyer aussi le cache de la barre des tâches
$taskbarCache = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"
if (Test-Path $taskbarCache) {
    $files = Get-ChildItem $taskbarCache -Filter "*.automaticDestinations-ms" -ErrorAction SilentlyContinue
    $count = ($files | Measure-Object).Count
    $files | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Host "    OK - $count fichier(s) de cache supprime(s)" -ForegroundColor Green
}

# Étape 3: Redémarrer l'explorateur Windows
Write-Host "[3/4] Redemarrage de l'explorateur Windows..." -ForegroundColor Yellow
Start-Process "explorer.exe"
Start-Sleep -Seconds 3
Write-Host "    OK - Explorateur redemarre" -ForegroundColor Green

# Étape 4: Instructions
Write-Host "[4/4] Instructions finales:" -ForegroundColor Yellow
Write-Host ""
Write-Host "OK - Le cache a ete nettoye" -ForegroundColor Green
Write-Host ""
Write-Host "Actions a effectuer:" -ForegroundColor Cyan
Write-Host "  1. Desepinglez l'application de la barre des taches (clic droit -> Desepingler)" -ForegroundColor White
Write-Host "  2. Lancez l'application depuis bin\Release\XnrgyEngineeringAutomationTools.exe" -ForegroundColor White
Write-Host "  3. Reepinglez l'application a la barre des taches" -ForegroundColor White
Write-Host ""
Write-Host "Le nouveau nom 'XNRGY Engineering Automation Tools' devrait maintenant apparaitre!" -ForegroundColor Green
Write-Host ""



