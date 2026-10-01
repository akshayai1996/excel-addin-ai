$certFile = Join-Path $PSScriptRoot 'server\cert.pem'
$certName = 'ExcelAIAssistantLocalhost'

# Read PEM and convert to X509Certificate2
$pemText = Get-Content $certFile -Raw
$b64 = $pemText -replace '-----BEGIN CERTIFICATE-----','' -replace '-----END CERTIFICATE-----','' -replace '\s',''
$derBytes = [Convert]::FromBase64String($b64)
$x509 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$derBytes)
$x509.FriendlyName = $certName

# Remove old cert with same friendly name from CurrentUser\Root
Get-ChildItem 'Cert:\CurrentUser\Root' | Where-Object { $_.FriendlyName -eq $certName } | Remove-Item -Force -ErrorAction SilentlyContinue

# Install into CurrentUser\Root (no admin needed)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root','CurrentUser')
$store.Open('ReadWrite')
$store.Add($x509)
$store.Close()

Write-Host 'SUCCESS: Certificate trusted in CurrentUser store.'
Write-Host ('Thumbprint: ' + $x509.Thumbprint)
