@echo off
REM Gantt Chart Builder - Deinstallation (entfernt die Registrierung)
powershell -ExecutionPolicy Bypass -Command "$m=Join-Path '%~dp0' 'manifest.prod.xml'; $k='HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'; if(Test-Path $k){ Remove-ItemProperty -Path $k -Name $m -ErrorAction SilentlyContinue }; $wef=\"$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\"; if(Test-Path $wef){ Remove-Item \"$wef*\" -Recurse -Force -ErrorAction SilentlyContinue }; Write-Host 'Add-in entfernt. Bitte PowerPoint neu starten.' -ForegroundColor Green; Read-Host 'Enter zum Beenden'"
