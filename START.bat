@echo off
cd /d "%~dp0"
title IP Tool - Start

echo ============================================================
echo   IP Address Allocation Tool
echo   Keep this window open - errors will show here.
echo ============================================================
echo.

set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY (
    python --version >nul 2>&1
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo   [ERROR] Python not found on this computer.
    echo.
    echo   Install Python 3.8+ from: https://www.python.org/downloads/
    echo   IMPORTANT: check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

echo Using: %PY%
"%PY%" --version
echo.

echo [1/3] Checking environment...
"%PY%" check_env.py >nul 2>&1
if not errorlevel 1 goto start

echo       Environment incomplete, checking dependencies...
"%PY%" -c "import PySide6, openpyxl, docx" >nul 2>&1
if not errorlevel 1 goto start

echo [2/3] Installing dependencies, please wait 1-3 minutes...
"%PY%" install_deps.py
if errorlevel 1 (
    echo.
    echo   [ERROR] Failed to install dependencies.
    echo   See the report above for the recommended fix.
    echo.
    pause
    exit /b 1
)

:start
echo.
echo [3/3] Starting application...
echo.
"%PY%" app_gui.py

if errorlevel 1 (
    echo.
    echo ============================================================
    echo   The application exited with an error.
    echo   See the message above, or check error.log in this folder.
    echo ============================================================
    echo.
    pause
)
