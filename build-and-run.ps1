# XNRGY Engineering Automation Tools - BUILD & RUN SCRIPT
# Version: 1.0.0

$ProjectName = "XnrgyEngineeringAutomationTools"

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "  XNRGY ENGINEERING AUTOMATION TOOLS - BUILD & RUN" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

# ETAPE 1: Forcer arret
Write-Host "[1/3] Arret des instances existantes..." -ForegroundColor Yellow
taskkill /F /IM "$ProjectName.exe" 2>$null | Out-Null
Start-Sleep -Seconds 1
Write-Host "      OK" -ForegroundColor Green
Write-Host ""

# ETAPE 2: Compiler
Write-Host "[2/3] Compilation en mode Release..." -ForegroundColor Yellow

$msbuildPath = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe"
if (-not (Test-Path $msbuildPath)) {
    $msbuildPath = "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe"
}
if (-not (Test-Path $msbuildPath)) {
    $msbuildPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"
}

Write-Host "      MSBuild: $msbuildPath" -ForegroundColor Gray

& $msbuildPath "$ProjectName.csproj" /p:Configuration=Release /t:Rebuild /v:minimal /nologo /m

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "      [ERREUR] Echec compilation" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "      [OK] Compilation reussie!" -ForegroundColor Green
Write-Host ""

$exePath = "bin\Release\$ProjectName.exe"
if (-not (Test-Path $exePath)) {
    Write-Host "      [ERREUR] Executable introuvable: $exePath" -ForegroundColor Red
    exit 1
}

# ETAPE 3: Lancer
Write-Host "[3/3] Lancement de l'application..." -ForegroundColor Yellow
Write-Host "      Chemin: $exePath" -ForegroundColor Gray

Start-Process -FilePath $exePath
Start-Sleep -Seconds 2
Write-Host "      [OK] Application lancee!" -ForegroundColor Green

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "  TERMINE" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""
