# install-addin.ps1 - registriert das Gantt-Add-in fuer den aktuellen Benutzer.
# Kein Server, kein Node noetig - die App wird vom Host (GitHub Pages) geladen.
# Voraussetzung: manifest.prod.xml liegt im selben Ordner wie dieses Skript.

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$manifest = Join-Path $root "manifest.prod.xml"

if (-not (Test-Path $manifest)) {
  Write-Host "manifest.prod.xml wurde nicht gefunden (im selben Ordner erwartet)." -ForegroundColor Red
  Read-Host "Enter zum Beenden"
  exit 1
}

# 1) Registry-Sideload: Manifest fuer den Benutzer registrieren
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name $manifest -Value $manifest -PropertyType String -Force | Out-Null

# 2) Office-Add-in-Cache leeren, damit es sofort erscheint
$wef = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wef) { Remove-Item "$wef\*" -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ""
Write-Host "Gantt Chart Builder wurde registriert." -ForegroundColor Green
Write-Host "Bitte PowerPoint komplett schliessen und neu oeffnen." -ForegroundColor Yellow
Write-Host "Danach erscheint im Start-Tab die Gruppe 'Gantt' mit dem Button 'Gantt Chart'." -ForegroundColor Yellow
Write-Host ""
Read-Host "Enter zum Beenden"
