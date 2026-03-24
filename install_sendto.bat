@echo off
echo ============================================================
echo  repack - register to SendTo
echo ============================================================
echo.

:: Find repack.exe
set "EXE_PATH=%~dp0dist\repack.exe"
if not exist "%EXE_PATH%" (
    set "EXE_PATH=%~dp0repack.exe"
    if not exist "%EXE_PATH%" (
        echo [ERROR] repack.exe not found.
        echo   Run build.bat first.
        pause
        exit /b 1
    )
)

echo repack.exe: %EXE_PATH%

:: SendTo folder
set "SENDTO_DIR=%APPDATA%\Microsoft\Windows\SendTo"
echo SendTo: %SENDTO_DIR%

:: Create shortcut via PowerShell
echo.
echo Creating shortcut...
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%SENDTO_DIR%\repack.lnk'); $sc.TargetPath = '%EXE_PATH%'; $sc.WorkingDirectory = Split-Path '%EXE_PATH%'; $sc.Description = 'Repack to store ZIP'; $sc.Save()"

if errorlevel 1 (
    echo [ERROR] Failed to create shortcut.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Done!
echo  Right-click a file in Explorer - Send to - repack
echo ============================================================
echo.
pause
