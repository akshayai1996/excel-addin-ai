# ==============================================================
#  Excel AI Assistant - Desktop Setup Script
#  Run this ONCE as Administrator to:
#  1. Generate the self-signed SSL certificate
#  2. Trust it in the Windows Certificate Store (required for
#     Excel Desktop's WebView2 embedded browser)
#  3. Show sideloading instructions
# ==============================================================

param(
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ServerDir = Join-Path $ScriptDir "server"
$CertFile  = Join-Path $ServerDir "cert.pem"
$KeyFile   = Join-Path $ServerDir "key.pem"
$CertName  = "ExcelAIAssistantLocalhost"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Excel AI Assistant - Desktop Setup" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ── Uninstall mode ───────────────────────────────────────────
if ($Uninstall) {
    Write-Host "[UNINSTALL] Removing trusted certificate..." -ForegroundColor Yellow
    $existing = Get-ChildItem -Path "Cert:\LocalMachine\Root" | Where-Object { $_.FriendlyName -eq $CertName }
    foreach ($c in $existing) {
        Remove-Item $c.PSPath -Force
        Write-Host "  Removed: $($c.Thumbprint)" -ForegroundColor Green
    }
    Write-Host "[DONE] Certificate removed." -ForegroundColor Green
    exit 0
}

# ── Check admin rights ────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERROR] This script must be run as Administrator." -ForegroundColor Red
    Write-Host "  Right-click setup_desktop.ps1 → 'Run with PowerShell as Administrator'" -ForegroundColor Yellow
    Write-Host ""
    pause
    exit 1
}

# ── Step 1: Generate certificate (via Python) ────────────────
Write-Host "[1/3] Generating SSL certificate for localhost..." -ForegroundColor White

if (-not (Test-Path $CertFile) -or -not (Test-Path $KeyFile)) {
    Write-Host "  Running Python to generate cert..." -ForegroundColor Gray
    Push-Location $ScriptDir
    python -c @"
import sys, os
sys.path.insert(0, os.path.join('server'))
# Inline cert generation (mirrors bridge_server.py logic)
import datetime, ipaddress
from pathlib import Path
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa

cert_file = Path('server/cert.pem')
key_file  = Path('server/key.pem')

key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
subject = issuer = x509.Name([
    x509.NameAttribute(NameOID.COUNTRY_NAME, 'IN'),
    x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, 'Maharashtra'),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, 'CommandCode Excel AI'),
    x509.NameAttribute(NameOID.COMMON_NAME, 'localhost'),
])
cert = (
    x509.CertificateBuilder()
    .subject_name(subject)
    .issuer_name(issuer)
    .public_key(key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(datetime.datetime.now(datetime.timezone.utc))
    .not_valid_after(datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=3650))
    .add_extension(x509.SubjectAlternativeName([
        x509.DNSName('localhost'),
        x509.IPAddress(ipaddress.IPv4Address('127.0.0.1')),
    ]), critical=False)
    .sign(key, hashes.SHA256())
)
key_file.write_bytes(key.private_bytes(
    encoding=serialization.Encoding.PEM,
    format=serialization.PrivateFormat.TraditionalOpenSSL,
    encryption_algorithm=serialization.NoEncryption(),
))
cert_file.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
print('Certificate generated.')
"@
    Pop-Location
} else {
    Write-Host "  Certificate already exists, skipping generation." -ForegroundColor Gray
}

if (-not (Test-Path $CertFile)) {
    Write-Host "[ERROR] Certificate generation failed. Make sure 'cryptography' is installed: pip install cryptography" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "  Certificate ready: $CertFile" -ForegroundColor Green

# ── Step 2: Import cert into Windows Trusted Root store ──────
Write-Host ""
Write-Host "[2/3] Trusting certificate in Windows Certificate Store..." -ForegroundColor White

# Remove any old cert with the same friendly name first
$existing = Get-ChildItem -Path "Cert:\LocalMachine\Root" | Where-Object { $_.FriendlyName -eq $CertName }
foreach ($c in $existing) {
    Remove-Item $c.PSPath -Force
    Write-Host "  Removed old cert: $($c.Thumbprint)" -ForegroundColor Gray
}

# Import the PEM cert
$certBytes  = [System.IO.File]::ReadAllBytes($CertFile)
# Strip PEM headers to get raw DER bytes
$b64 = [System.Text.Encoding]::ASCII.GetString($certBytes) `
    -replace '-----BEGIN CERTIFICATE-----','' `
    -replace '-----END CERTIFICATE-----','' `
    -replace '\s',''
$derBytes = [Convert]::FromBase64String($b64)

$x509 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$derBytes)
$x509.FriendlyName = $CertName

$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
$store.Open("ReadWrite")
$store.Add($x509)
$store.Close()

Write-Host "  Certificate trusted. Thumbprint: $($x509.Thumbprint)" -ForegroundColor Green

# ── Step 3: Sideloading instructions ────────────────────────
Write-Host ""
Write-Host "[3/3] Sideloading instructions for Excel Desktop" -ForegroundColor White
Write-Host ""
Write-Host "  To load the add-in into Excel Desktop:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  METHOD A – Shared Folder Catalog (Recommended):" -ForegroundColor Cyan
Write-Host "    1. Open Excel → File → Options → Trust Center → Trust Center Settings"
Write-Host "    2. Go to Trusted Add-in Catalogs"
Write-Host "    3. Add this folder as a catalog: $ScriptDir"
Write-Host "    4. Check 'Show in Menu', click OK, restart Excel"
Write-Host "    5. Go to Insert → My Add-ins → Shared Folder → select 'AI Assistant'"
Write-Host ""
Write-Host "  METHOD B – Developer Tab (Quickest for testing):" -ForegroundColor Cyan
Write-Host "    1. In Excel: File → Options → Customize Ribbon → Enable 'Developer' tab"
Write-Host "    2. Developer tab → Add-ins → Browse → select: $ScriptDir\manifest.xml"
Write-Host ""
Write-Host "  IMPORTANT: Start the bridge server first!" -ForegroundColor Red
Write-Host "    Double-click: start_server.bat" -ForegroundColor Yellow
Write-Host ""

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Setup Complete! Follow the sideloading steps above." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

pause
