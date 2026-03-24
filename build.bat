@echo off
echo ============================================================
echo  repack - build script
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo   Install from: https://www.python.org/
    pause
    exit /b 1
)

:: Install dependencies
echo [1/3] Installing packages...
python -m pip install pyinstaller send2trash tomli --quiet
if errorlevel 1 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)
echo       OK

:: Build with PyInstaller
echo [2/3] Building exe...
pyinstaller ^
    --onefile ^
    --console ^
    --name repack ^
    --hidden-import=tomli ^
    --hidden-import=send2trash ^
    repack.py

if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)
echo       OK

:: Copy config to dist
echo [3/3] Copying config...
copy /y config.toml dist\config.toml >nul
echo       OK

echo.
echo ============================================================
echo  Build complete!
echo  Place dist\repack.exe and dist\config.toml in the same folder.
echo  Then run install_sendto.bat
echo ============================================================
echo.
pause
