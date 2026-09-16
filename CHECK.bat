@echo off
cd /d "%~dp0"
title IP Tool - Diagnostics

set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY (
    python --version >nul 2>&1
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo.
    echo   [ERROR] Python not found on this computer.
    echo.
    echo   Install Python 3.8+ from: https://www.python.org/downloads/
    echo   IMPORTANT: check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

set "CHECK_PAUSE=1"
"%PY%" check_env.py
echo.
pause
