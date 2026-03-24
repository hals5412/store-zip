@echo off
chcp 65001 >nul
echo ============================================================
echo  repack を SendTo に登録します
echo ============================================================
echo.

:: repack.exe の場所を確認
set "EXE_PATH=%~dp0dist\repack.exe"
if not exist "%EXE_PATH%" (
    :: dist フォルダがない場合、同じフォルダを確認
    set "EXE_PATH=%~dp0repack.exe"
    if not exist "%EXE_PATH%" (
        echo [ERROR] repack.exe が見つかりません。
        echo   先に build.bat を実行してください。
        pause
        exit /b 1
    )
)

echo repack.exe のパス: %EXE_PATH%

:: SendTo フォルダのパスを取得
set "SENDTO_DIR=%APPDATA%\Microsoft\Windows\SendTo"
echo SendTo フォルダ: %SENDTO_DIR%

:: ショートカットを作成 (PowerShell 使用)
echo.
echo ショートカットを作成中...
powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell; ^
     $sc = $ws.CreateShortcut('%SENDTO_DIR%\repack (無圧縮ZIP変換).lnk'); ^
     $sc.TargetPath = '%EXE_PATH%'; ^
     $sc.WorkingDirectory = Split-Path '%EXE_PATH%'; ^
     $sc.Description = '圧縮ファイルを無圧縮ZIPに変換'; ^
     $sc.Save()"

if errorlevel 1 (
    echo [ERROR] ショートカットの作成に失敗しました。
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  登録完了！
echo  エクスプローラーで圧縮ファイルを右クリック
echo  → 「送る」→「repack (無圧縮ZIP変換)」で使用できます。
echo ============================================================
echo.
pause
