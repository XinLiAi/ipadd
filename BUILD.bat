@echo off
cd /d "%~dp0"
title IP Tool - Build EXE

echo ============================================================
echo   IP Address Allocation Tool - Build EXE
echo ============================================================
echo.

echo [1/5] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] Python not found.
    echo   Install Python 3.8+ from https://www.python.org/downloads/
    echo   Remember to check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
python --version
echo.

echo [2/5] Creating virtual environment...
if not exist ".venv" (
    python -m venv .venv
    if errorlevel 1 (
        echo   [ERROR] Failed to create virtual environment
        pause
        exit /b 1
    )
) else (
    echo   Already exists, skipped
)
echo.

echo [3/6] Installing dependencies, 1-3 minutes...
".venv\Scripts\python.exe" install_deps.py
if errorlevel 1 (
    echo.
    echo   [ERROR] Failed to install dependencies.
    echo   See the report above for the recommended fix.
    echo   Most common cause: lxml has no prebuilt wheel for this
    echo   Python version. Python 3.11 or 3.12 is recommended.
    echo.
    pause
    exit /b 1
)
echo.

echo [4/6] Cleaning old output...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
echo.

echo [5/6] Checking PyInstaller...
".venv\Scripts\python.exe" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo       Not installed, installing now...
    ".venv\Scripts\python.exe" -m pip install --only-binary=:all: pyinstaller -i https://mirrors.cloud.tencent.com/pypi/simple
    if errorlevel 1 (
        echo       Trying older version 6.3.0...
        ".venv\Scripts\python.exe" -m pip install --only-binary=:all: pyinstaller==6.3.0 -i https://mirrors.cloud.tencent.com/pypi/simple
    )
    if errorlevel 1 (
        echo       Trying older version 5.13.2...
        ".venv\Scripts\python.exe" -m pip install --only-binary=:all: pyinstaller==5.13.2 -i https://mirrors.cloud.tencent.com/pypi/simple
    )
)
".venv\Scripts\python.exe" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo.
    echo   [ERROR] PyInstaller is not available, cannot build EXE.
    echo   Check your network and try again.
    echo   You can still run the app with START.vbs.
    echo.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" -c "import PyInstaller;print('       PyInstaller',PyInstaller.__version__)"
echo.

echo [6/6] Building EXE, 1-3 minutes, please wait...
".venv\Scripts\python.exe" -m PyInstaller IPAddressTool.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo   [ERROR] Build failed. See errors above.
    pause
    exit /b 1
)

echo.
echo   Building diagnostic version...
".venv\Scripts\python.exe" -m PyInstaller DiagBuild.spec --noconfirm
echo.

set "QTB="
for /f "delims=" %%i in ('".venv\Scripts\python.exe" -c "import qt_compat as q;print(q.BINDING)"') do set "QTB=%%i"
if not defined QTB set "QTB=Unknown"

echo ============================================================
echo   Build finished.
echo.
echo   Qt binding bundled: %QTB%
echo.
echo   Output files:
echo       dist\IPAddressTool_%QTB%.exe
echo       dist\IPAddressTool_%QTB%_Diag.exe     (shows errors)
echo.
echo   IMPORTANT - where this EXE will run:
if /i "%QTB%"=="PySide2" (
    echo       Qt 5 ^(PySide2^)  -^>  Windows 7 SP1 / Server 2008 R2
    echo                            and later. NOT on Server 2008 ^(non-R2^).
) else (
    echo       Qt 6 ^(PySide6^)  -^>  Windows 10 1809 / Server 2016
    echo                            and later ONLY.
    echo       It will NOT run on Windows 7 / Server 2008 R2.
)
echo.
echo   Do NOT copy this EXE to an older Windows and expect it
echo   to work - build it on the target machine instead.
echo ============================================================
echo.
pause
