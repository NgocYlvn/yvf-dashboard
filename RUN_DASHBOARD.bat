@echo off
setlocal
cd /d "%~dp0"
title YVF Management Dashboard

echo =====================================================
echo       YVF MANAGEMENT DASHBOARD - STARTING
echo =====================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python was not found.
  echo Please install Python and select "Add Python to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating local Python environment...
  python -m venv .venv
  if errorlevel 1 goto :error
)

echo [2/3] Checking required libraries...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 goto :error

echo [3/3] Opening dashboard in your browser...
".venv\Scripts\python.exe" -m streamlit run app.py --server.headless false --browser.gatherUsageStats false
exit /b 0

:error
echo.
echo [ERROR] Dashboard could not start.
echo Please take a screenshot of this window and send it for checking.
pause
exit /b 1
