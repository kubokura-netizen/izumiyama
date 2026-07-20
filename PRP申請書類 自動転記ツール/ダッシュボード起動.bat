@echo off
rem ============================================================
rem  PRP申請書類 自動転記ツール ダッシュボード 起動（Windows）
rem   このファイルをダブルクリックすると、
rem   サーバ起動 → ブラウザ表示 まで自動で行います。
rem
rem  _runtime フォルダに Python 一式が同梱されているため、
rem  お使いのPCへの Python インストールは不要です。
rem ============================================================
chcp 932 >nul
setlocal
cd /d "%~dp0"
set DASH=98_dashboard
set RUNTIME=_runtime\python\python.exe

rem ---- 同梱ランタイムがあればそれで起動（配布先の通常ルート） ----
if exist "%RUNTIME%" (
  echo.
  echo ダッシュボードを起動します。ブラウザが自動で開きます。
  echo 終了するには、この黒い画面で Ctrl + C を押してください。
  echo.
  "%RUNTIME%" "%DASH%\app.py"
  goto :done
)

rem ---- 同梱ランタイムが無い場合（開発・メンテ用のフォールバック） ----
echo.
echo [お知らせ] 同梱ランタイム(_runtime)が見つかりません。
echo            PC の Python を使って起動を試みます。
echo.
set VENV=.venv_dashboard

where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)
where %PY% >nul 2>nul
if not %errorlevel%==0 (
  echo ------------------------------------------------------------
  echo Python が見つからないため起動できません。
  echo.
  echo 配布されたフォルダをお使いの場合は、_runtime フォルダごと
  echo コピーされているかご確認ください（コピー漏れの可能性があります）。
  echo ------------------------------------------------------------
  pause
  exit /b 1
)

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

:done
pause
endlocal
