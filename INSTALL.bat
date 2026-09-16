@echo off
cd /d "%~dp0"
title IP Tool - Install Dependencies

echo ============================================================
echo   IP Address Allocation Tool - Install Dependencies
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
echo If pip tries to COMPILE lxml and fails with
echo     "Microsoft Visual C++ 14.0 or greater is required"
echo this script will detect it and apply the right fix.
echo.
pause

"%PY%" install_deps.py

echo.
pause
