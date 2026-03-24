@echo off
chcp 65001 >nul
echo ============================================================
echo  repack.exe ビルドスクリプト
echo ============================================================
echo.

:: Python の確認
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python が見つかりません。
    echo   https://www.python.org/ からインストールしてください。
    pause
    exit /b 1
)

:: 必要なパッケージのインストール
echo [1/3] 依存パッケージをインストール中...
pip install pyinstaller send2trash tomli --quiet
if errorlevel 1 (
    echo [ERROR] パッケージのインストールに失敗しました。
    pause
    exit /b 1
)
echo      OK

:: PyInstaller でビルド
echo [2/3] exe をビルド中...
pyinstaller ^
    --onefile ^
    --console ^
    --name repack ^
    --hidden-import=tomli ^
    --hidden-import=send2trash ^
    repack.py

if errorlevel 1 (
    echo [ERROR] ビルドに失敗しました。
    pause
    exit /b 1
)
echo      OK

:: dist フォルダに config.toml をコピー
echo [3/3] 設定ファイルをコピー中...
copy /y config.toml dist\config.toml >nul
echo      OK

echo.
echo ============================================================
echo  ビルド完了！
echo  dist\repack.exe と dist\config.toml を同じフォルダに置いて使用してください。
echo  次に install_sendto.bat を実行してください。
echo ============================================================
echo.
pause
