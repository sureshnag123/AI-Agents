@echo off
REM Starts the Fracktal Works HR Portal (webapp) and opens it in the browser.
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Setting up environment for the first time, please wait...
    powershell -NoProfile -ExecutionPolicy Bypass -File "setup.ps1"
)

echo Starting Fracktal HR Portal...
start "Fracktal HR Portal" /min ".venv\Scripts\python.exe" "webapp\app.py"

REM Give the server a moment to start, then open the browser
timeout /t 3 /nobreak >nul
start "" "http://localhost:5000"

endlocal
