@echo off
setlocal

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
set "PYTHON_EXE=%ROOT%\.runtime\python\python.exe"
set "SETUP_SCRIPT=%ROOT%\setup_runtime.ps1"
set "START_SCRIPT=%ROOT%\start_streamlit.ps1"

if not exist "%PYTHON_EXE%" (
  echo Preparing the project Python runtime. This is only needed once...
  powershell -NoProfile -ExecutionPolicy Bypass -File "%SETUP_SCRIPT%"
  if errorlevel 1 (
    echo.
    echo Runtime setup failed. Check the message above, then run this file again.
    pause
    exit /b 1
  )
)

"%PYTHON_EXE%" -c "import streamlit, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Repairing project dependencies...
  powershell -NoProfile -ExecutionPolicy Bypass -File "%SETUP_SCRIPT%" -Repair
  if errorlevel 1 (
    echo.
    echo Dependency repair failed. Check the message above, then run this file again.
    pause
    exit /b 1
  )
)

set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if not exist "%CHROME_EXE%" (
  echo Could not find chrome.exe.
  echo Please install Google Chrome or edit this bat file with your Chrome path.
  pause
  exit /b 1
)

set "PROFILE_DIR=%ROOT%\chrome_profile"
if not exist "%PROFILE_DIR%" mkdir "%PROFILE_DIR%"

echo Starting Amazon Selection Agent...
powershell -NoProfile -ExecutionPolicy Bypass -File "%START_SCRIPT%" -PythonExe "%PYTHON_EXE%" -Root "%ROOT%"
if errorlevel 1 (
  echo Streamlit could not be started.
  pause
  exit /b 1
)

start "" "%CHROME_EXE%" --remote-debugging-port=9222 --user-data-dir="%PROFILE_DIR%" --no-first-run --no-default-browser-check --new-window "http://localhost:8501/"
exit /b 0
