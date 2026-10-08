@echo off
rem Startet den Scraper (Standard: 500 Haeuser). Weitere Optionen werden durchgereicht,
rem z.B.:  starten.bat --ziel 200 --nur-kauf
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" call "%~dp0installieren.bat" --still
if not exist ".venv\Scripts\python.exe" goto ende
".venv\Scripts\python.exe" kleinanzeigen_haeuser.py %*
:ende
echo.
pause
