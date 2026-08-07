@echo off
rem ============================================================
rem  Receipt Reader - RUN
rem  Reads PDFs in 01_input and writes a draft into 02_output.
rem  (All messages in Japanese are printed by the Python script.)
rem  ASCII-only on purpose so cmd never mis-parses it.
rem ============================================================
chcp 65001 >nul
title Receipt Reader
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" (
  echo.
  echo [!] Setup is not done yet. Please run the setup batch first ^(shokai-junbi^).
  echo.
  pause
  exit /b 1
)

rem Start Ollama if installed (harmless if it is already running)
where ollama >nul 2>nul && start /b "" ollama serve >nul 2>nul

echo Reading PDFs in 01_input ... please wait.
echo ------------------------------------------------------------
"%PY%" "src\receipt_ocr.py"
echo ------------------------------------------------------------
echo Done. Please open the file in the 02_output folder.
pause
