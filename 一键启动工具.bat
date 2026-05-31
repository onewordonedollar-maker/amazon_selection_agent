@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON_EXE=C:\Users\color\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if not exist "%CHROME_EXE%" (
  echo Could not find chrome.exe.
  echo Please install Google Chrome or edit this bat file with your Chrome path.
  pause
  exit /b 1
)

set "PROFILE_DIR=%ROOT%chrome_profile"
if not exist "%PROFILE_DIR%" mkdir "%PROFILE_DIR%"

echo Starting Amazon Selection Agent...
echo.
echo 1. Checking Streamlit on http://localhost:8501/
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $r = Invoke-WebRequest -UseBasicParsing 'http://localhost:8501/' -TimeoutSec 1; if ($r.StatusCode -eq 200) { exit 0 } else { exit 1 } } catch { exit 1 }"
if errorlevel 1 (
  echo Streamlit is not running. Starting it now...
  start "Amazon Selection Agent - Streamlit" cmd /k ""%PYTHON_EXE%" -m streamlit run "%ROOT%streamlit_app.py" --server.port 8501 --server.headless true"
  timeout /t 4 /nobreak >nul
) else (
  echo Streamlit is already running.
)

echo 2. Opening collection Chrome with remote debugging port 9222
start "" "%CHROME_EXE%" --remote-debugging-port=9222 --user-data-dir="%PROFILE_DIR%" --no-first-run --no-default-browser-check --new-window "http://localhost:8501/"

echo.
echo Done.
echo Keep this Chrome window open while collecting SellerSprite data.
echo If this is the first time, install/login SellerSprite and Amazon in this Chrome profile.
echo Profile: %PROFILE_DIR%
echo.
pause
