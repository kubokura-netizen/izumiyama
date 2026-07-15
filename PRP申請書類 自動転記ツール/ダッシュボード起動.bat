@echo off
rem ============================================================
rem  PRP 自動転記ツール ダッシュボード 起動（Windows）
rem   このファイルをダブルクリックすると、
rem   必要な準備（初回のみ）→ サーバ起動 → ブラウザ表示 まで自動。
rem ============================================================
chcp 932 >nul
setlocal
cd /d "%~dp0"
set DASH=98_dashboard
set VENV=.venv_dashboard

rem python の存在確認（py 優先）
where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)
where %PY% >nul 2>nul
if not %errorlevel%==0 (
  echo Python が見つかりません。https://www.python.org からインストールしてください。
  pause
  exit /b 1
)

rem 初回のみ: 仮想環境を作って必要ライブラリを入れる
if not exist "%VENV%" (
  echo 初回セットアップ中（1-2分）... 必要なライブラリを準備します。
  %PY% -m venv "%VENV%"
  "%VENV%\Scripts\python.exe" -m pip install --upgrade pip >nul
  "%VENV%\Scripts\python.exe" -m pip install -r "%DASH%\requirements.txt"
)

echo.
echo ダッシュボードを起動します。ブラウザが自動で開きます。
echo 終了するには、この黒い画面で Ctrl + C を押してください。
echo.
"%VENV%\Scripts\python.exe" "%DASH%\app.py"

pause
endlocal
