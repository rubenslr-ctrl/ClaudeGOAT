@echo off
rem Installiert alles, was der Scraper braucht (Python, virtuelle Umgebung, Pakete).
rem Einfach doppelklicken. Mehrfaches Ausfuehren schadet nicht.
chcp 65001 >nul
cd /d "%~dp0"
echo === Kleinanzeigen-Scraper: Installation ===

set "PYEXE="
set "PYARG="
where py >nul 2>nul && set "PYEXE=py" && set "PYARG=-3"
if defined PYEXE goto python_ok
python -c "import sys" >nul 2>nul && set "PYEXE=python"
if defined PYEXE goto python_ok

echo Python wurde nicht gefunden - Installation ueber winget ...
where winget >nul 2>nul || goto kein_winget
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%ProgramFiles%\Python312\python.exe" set "PYEXE=%ProgramFiles%\Python312\python.exe"
if defined PYEXE goto python_ok
echo Python wurde installiert. Bitte VS Code bzw. dieses Fenster schliessen
echo und installieren.bat noch einmal starten.
goto ende_fehler

:kein_winget
echo winget ist auf diesem PC nicht verfuegbar.
echo Bitte Python von https://www.python.org/downloads/ installieren,
echo dabei "Add python.exe to PATH" anhaken und danach diese Datei erneut starten.
goto ende_fehler

:python_ok
if exist ".venv\Scripts\python.exe" goto venv_ok
echo Erstelle virtuelle Umgebung .venv ...
"%PYEXE%" %PYARG% -m venv .venv || goto ende_fehler

:venv_ok
".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
".venv\Scripts\python.exe" -m pip install -r requirements.txt || goto ende_fehler
echo.
echo Installation fertig. Starten mit test_starten.bat oder starten.bat
if /i not "%~1"=="--still" pause
exit /b 0

:ende_fehler
echo.
echo Installation NICHT abgeschlossen.
if /i not "%~1"=="--still" pause
exit /b 1
