# make-prod.ps1 - erzeugt docs/ (self-contained) + manifest.prod.xml fuer eine Host-URL.
# Beispiel:
#   .\make-prod.ps1 -HostUrl "https://<user>.github.io/<repo>"
param(
  [Parameter(Mandatory = $true)]
  [string]$HostUrl
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$nodeDir = "$env:LOCALAPPDATA\nodejs"
if (Test-Path $nodeDir) { $env:Path = "$nodeDir;$env:Path" }

# Trailing slash entfernen
$HostUrl = $HostUrl.TrimEnd("/")

# 1) Self-contained docs/taskpane.html bauen
node "$root\build.js"

# 2) manifest.prod.xml aus Template erzeugen
$tpl = Get-Content "$root\manifest.prod.template.xml" -Raw
$tpl = $tpl -replace "__HOST_URL__", $HostUrl
Set-Content "$root\manifest.prod.xml" $tpl -Encoding UTF8

# 3) manifest auch in docs/ ablegen (praktisch zum Teilen)
Copy-Item "$root\manifest.prod.xml" "$root\docs\manifest.xml" -Force

Write-Host ""
Write-Host "Fertig!" -ForegroundColor Green
Write-Host "  Host-URL:      $HostUrl" -ForegroundColor Green
Write-Host "  Statische App: docs\taskpane.html" -ForegroundColor Green
Write-Host "  Manifest:      manifest.prod.xml" -ForegroundColor Green
Write-Host ""
Write-Host "Naechste Schritte:" -ForegroundColor Yellow
Write-Host "  1. Ordner 'docs' auf GitHub Pages (oder anderen HTTPS-Host) veroeffentlichen."
Write-Host "  2. Kollegen bekommen 'manifest.prod.xml' + fuehren 'install-addin.cmd' aus."
