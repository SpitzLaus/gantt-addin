@echo off
REM Gantt Chart Builder - Ein-Klick-Installer (Doppelklick)
REM Registriert das Add-in fuer den aktuellen Benutzer. Kein Server noetig.
powershell -ExecutionPolicy Bypass -File "%~dp0install-addin.ps1"
