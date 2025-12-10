@echo off
chcp 65001 > nul

echo ============================================================
echo   KPMG Workbench API テスト
echo ============================================================
echo.

REM 仮想環境の確認
if not exist .venv (
    echo [エラー] 仮想環境が見つかりません。
    echo まず setup.bat を実行してください。
    pause
    exit /b 1
)

REM .env ファイルの確認
if not exist .env (
    echo [エラー] .env ファイルが見つかりません。
    echo まず setup.bat を実行してください。
    pause
    exit /b 1
)

REM 仮想環境の有効化
call .venv\Scripts\activate.bat

REM API テストの実行
python test_api.py

echo.
pause
