@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Stundennachweis.ps1" -Aktion Anzeigen
pause
