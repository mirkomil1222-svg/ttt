@echo off
REM Build script for creating Rash Manager v2 Windows executable
REM Run this script on Windows with Python installed

echo ========================================
echo   Rash Manager v2 - EXE Builder
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://www.python.org/
    pause
    exit /b 1
)

echo [1/4] Checking Python version...
python --version
echo.

echo [2/4] Installing/upgrading dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)
echo.

echo [3/4] Installing PyInstaller if needed...
python -m pip install pyinstaller
echo.

echo [4/4] Building executable...
echo This may take several minutes...
echo.
python -m PyInstaller build_exe.spec --clean --noconfirm
if errorlevel 1 (
    echo ERROR: Failed to build executable
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Build Complete!
echo ========================================
echo.
echo The executable is located at:
echo   dist\RashManager.exe
echo.
echo IMPORTANT NOTES:
echo 1. Make sure to copy these files to the same folder as RashManager.exe:
echo    - titul_bubble_koordinatalar_2480x3508.xlsx
echo    - Titul.pdf (if you have it)
echo.
echo 2. For PDF processing, you need to install Poppler for Windows:
echo    Download from: https://github.com/oschwartz10612/poppler-windows/releases
echo    Extract and add poppler/bin to your PATH, OR
echo    Copy poppler/bin folder contents to the same folder as the exe
echo.
echo 3. The program will create these folders automatically:
echo    - output/
echo    - test_keys/
echo.
pause

