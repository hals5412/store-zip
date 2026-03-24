@echo off
echo ============================================================
echo  repack - install / update
echo ============================================================
echo.

:: Move to the folder where this bat file is located
cd /d "%~dp0"

:: ── [1/4] Check Python ──────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo   Install from: https://www.python.org/
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo Python: %%v

:: ── [2/4] Install / update packages ────────────────────────
echo.
echo [2/4] Installing packages...
python -m pip install --upgrade pyinstaller send2trash tomli --quiet
if errorlevel 1 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)
echo       OK

:: ── [3/4] Build exe ─────────────────────────────────────────
echo.
echo [3/4] Building exe...
python -m PyInstaller --onefile --console --name repack --hidden-import=tomli --hidden-import=send2trash --noconfirm repack.py
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

:: Copy config to dist (only if not already present, to preserve user edits)
if not exist "dist\config.toml" (
    copy /y config.toml dist\config.toml >nul
    echo       config.toml copied to dist\
) else (
    echo       config.toml already exists in dist\ - skipped to preserve settings
)
echo       OK

:: ── [4/4] Register to SendTo ────────────────────────────────
echo.
echo [4/4] Registering to SendTo...
set "EXE_PATH=%~dp0dist\repack.exe"
set "SENDTO_DIR=%APPDATA%\Microsoft\Windows\SendTo"
set "LNK_PATH=%SENDTO_DIR%\repack.lnk"

powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%LNK_PATH%'); $sc.TargetPath = '%EXE_PATH%'; $sc.WorkingDirectory = '%~dp0dist'; $sc.Description = 'Repack to store ZIP'; $sc.Save()"
if errorlevel 1 (
    echo [ERROR] Failed to register SendTo shortcut.
    pause
    exit /b 1
)
echo       Registered: %LNK_PATH%
echo       OK

:: ── Done ────────────────────────────────────────────────────
echo.
echo ============================================================
echo  Done!
echo  Right-click a file in Explorer - Send to - repack
echo ============================================================
echo.
pause
