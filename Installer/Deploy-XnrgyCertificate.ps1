param([switch]$AllUsers, [switch]$Uninstall)
$ErrorActionPreference = "Stop"
$CertThumbprint = "46F467C6081BC6F42507FF02CDF3D418373081E7"

Write-Host "`n========== XNRGY Certificate Deployment Tool ==========`n" -ForegroundColor Cyan

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if ($AllUsers -and -not (Test-Admin)) {
    Write-Host "[-] ERREUR: -AllUsers necessite les droits admin!" -ForegroundColor Red
    exit 1
}

$loc = if ($AllUsers) { "LocalMachine" } else { "CurrentUser" }

if ($Uninstall) {
    Write-Host "[>] Suppression..." -ForegroundColor Yellow
    foreach ($s in @("TrustedPublisher","Root")) {
        $c = Get-ChildItem "Cert:\$loc\$s" -EA SilentlyContinue | Where-Object {$_.Thumbprint -eq $CertThumbprint}
        if ($c) {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s,$loc)
            $store.Open("ReadWrite"); $store.Remove($c); $store.Close()
            Write-Host "[+] Supprime de $loc\$s" -ForegroundColor Green
        }
    }
    exit 0
}

Write-Host "[1/3] Recherche certificat..." -ForegroundColor Yellow
$cert = Get-ChildItem "Cert:\CurrentUser\My" -EA SilentlyContinue | Where-Object {$_.Thumbprint -eq $CertThumbprint}

if (-not $cert) {
    $pfx = Join-Path $PSScriptRoot "XnrgyCodeSigning.pfx"
    if (Test-Path $pfx) {
        $pwd = ConvertTo-SecureString "Xnrgy2026!" -AsPlainText -Force
        $cert = Import-PfxCertificate -FilePath $pfx -CertStoreLocation "Cert:\CurrentUser\My" -Password $pwd
    }
}

if (-not $cert) { Write-Host "[-] Certificat introuvable!" -ForegroundColor Red; exit 1 }
Write-Host "      [+] Trouve: $($cert.Subject)" -ForegroundColor Green

foreach ($s in @("TrustedPublisher","Root")) {
    Write-Host "[>] Installation dans $s..." -ForegroundColor Yellow
    $ex = Get-ChildItem "Cert:\$loc\$s" -EA SilentlyContinue | Where-Object {$_.Thumbprint -eq $cert.Thumbprint}
    if ($ex) { Write-Host "      [i] Deja installe" -ForegroundColor Cyan }
    else {
        try {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s,$loc)
            $store.Open("ReadWrite"); $store.Add($cert); $store.Close()
            Write-Host "      [+] OK!" -ForegroundColor Green
        } catch { Write-Host "      [!] Erreur: $_" -ForegroundColor Yellow }
    }
}

Write-Host "`n========== CERTIFICAT INSTALLE! ==========`n" -ForegroundColor Green
Write-Host "Publisher sera: XNRGY Climate Systems ULC" -ForegroundColor Cyan
Write-Host "Plus d avertissement SmartScreen!`n" -ForegroundColor Cyan
