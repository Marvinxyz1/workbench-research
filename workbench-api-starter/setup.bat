@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ============================================================
echo   KPMG Workbench API Starter - セットアップ
echo ============================================================
echo.

REM Python インストール確認
echo [1/5] Python のインストールを確認しています...
python --version > nul 2>&1
if errorlevel 1 (
    echo.
    echo [エラー] Python がインストールされていません。
    echo.
    echo 以下のリンクから Python をダウンロードしてください:
    echo   https://www.python.org/downloads/
    echo.
    echo インストール時に「Add Python to PATH」にチェックを入れてください。
    echo.
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo   Python %PYTHON_VERSION% が見つかりました。

REM 仮想環境の作成
echo.
echo [2/5] 仮想環境を作成しています...
if exist .venv (
    echo   .venv フォルダが既に存在します。スキップします。
) else (
    python -m venv .venv
    if errorlevel 1 (
        echo [エラー] 仮想環境の作成に失敗しました。
        pause
        exit /b 1
    )
    echo   仮想環境を作成しました。
)

REM 仮想環境の有効化
echo.
echo [3/5] 仮想環境を有効化しています...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [エラー] 仮想環境の有効化に失敗しました。
    pause
    exit /b 1
)
echo   仮想環境を有効化しました。

REM 依存関係のインストール
echo.
echo [4/5] 依存関係をインストールしています...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [エラー] 依存関係のインストールに失敗しました。
    pause
    exit /b 1
)
echo   依存関係をインストールしました。

REM .env ファイルの作成
echo.
echo [5/5] 環境設定ファイルを準備しています...
if exist .env (
    echo   .env ファイルが既に存在します。スキップします。
) else (
    copy .env.example .env > nul
    echo   .env ファイルを作成しました。
)

echo.
echo ============================================================
echo   セットアップ完了！
echo ============================================================
echo.
echo 次のステップ:
echo   1. .env ファイルを開いて API Key を設定してください
echo   2. run.bat を実行して API テストを行ってください
echo.
pause
