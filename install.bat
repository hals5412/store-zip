@echo off
echo ============================================================
echo  repack - install / update
echo ============================================================
echo.

cd /d "%~dp0"

:: [1/4] Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo   Install from: https://www.python.org/
    pause
    exit /b 1
)
echo [1/4] Python OK

:: [2/4] Install / update packages
echo [2/4] Installing packages...
python -m pip install --upgrade pyinstaller send2trash tomli --quiet
if errorlevel 1 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)
echo [2/4] Packages OK

:: [3/4] Build exe
echo [3/4] Building exe...
python -m PyInstaller --onefile --console --name repack --hidden-import=tomli --hidden-import=send2trash --noconfirm repack.py
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

if not exist "dist\config.toml" (
    copy /y config.toml dist\config.toml >nul
    echo       config.toml copied.
) else (
    echo       config.toml already exists - skipped to preserve settings.
)
echo [3/4] Build OK

:: [4/4] Register to SendTo
echo [4/4] Registering to SendTo...
set "EXE_PATH=%~dp0dist\repack.exe"
set "LNK_PATH=%APPDATA%\Microsoft\Windows\SendTo\repack.lnk"
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%LNK_PATH%'); $sc.TargetPath = '%EXE_PATH%'; $sc.WorkingDirectory = '%~dp0dist'; $sc.Description = 'Repack to store ZIP'; $sc.Save()"
if errorlevel 1 (
    echo [ERROR] Failed to register SendTo shortcut.
    pause
    exit /b 1
)
echo [4/4] SendTo OK

echo.
echo ============================================================
echo  Done! Right-click a file - Send to - repack
echo ============================================================
echo.
pause
