@echo off
echo ============================================================
echo  repack - SendTo registration
echo ============================================================
echo.

cd /d "%~dp0"

set "EXE_PATH=%~dp0repack.exe"
set "LNK_PATH=%APPDATA%\Microsoft\Windows\SendTo\repack.lnk"

if not exist "%EXE_PATH%" (
    echo [ERROR] repack.exe not found in this folder.
    pause
    exit /b 1
)

echo Registering to SendTo...
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%LNK_PATH%'); $sc.TargetPath = '%EXE_PATH%'; $sc.WorkingDirectory = '%~dp0'; $sc.Description = 'Repack to store ZIP'; $sc.Save()"
if errorlevel 1 (
    echo [ERROR] Failed to register SendTo shortcut.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Done! Right-click a file - Send to - repack
echo ============================================================
echo.
pause
