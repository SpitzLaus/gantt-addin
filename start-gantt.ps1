# Gantt Add-in - Server-Starter
# Rechtsklick -> "Mit PowerShell ausfuehren"  ODER  im Terminal: .\start-gantt.ps1

$ErrorActionPreference = "Stop"
$nodeDir = "$env:LOCALAPPDATA\nodejs"
$proj    = $PSScriptRoot
$crt     = "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt"
$key     = "$env:USERPROFILE\.office-addin-dev-certs\localhost.key"

# Node in den PATH (portable Installation)
if (Test-Path $nodeDir) { $env:Path = "$nodeDir;$env:Path" }

# Pruefen, ob node verfuegbar ist
try { node --version | Out-Null } catch {
  Write-Host "Node.js nicht gefunden. Bitte zuerst Node installieren." -ForegroundColor Red
  exit 1
}

# Dev-Zertifikate sicherstellen
if (-not (Test-Path $crt)) {
  Write-Host "Installiere Entwickler-Zertifikate..." -ForegroundColor Cyan
  npx --yes office-addin-dev-certs install
}

Set-Location $proj
Write-Host ""
Write-Host "Gantt-Add-in Server startet auf https://localhost:3000" -ForegroundColor Green
Write-Host "Manifest zum Sideloading: $proj\manifest.xml" -ForegroundColor Green
Write-Host "Dieses Fenster offen lassen. Beenden mit STRG+C." -ForegroundColor Yellow
Write-Host ""

npx --yes http-server . -p 3000 -S -C $crt -K $key --cors -c-1
