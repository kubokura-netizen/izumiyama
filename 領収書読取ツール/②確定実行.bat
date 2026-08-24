@echo off
rem ============================================================
rem  Receipt Reader - APPLY (confirm & commit)
rem  Renames + sorts originals into year-month folders and
rem  appends rows into the master expense workbook.
rem  Reads the draft(s) in 02_output where the row is marked.
rem  ASCII-only header; Japanese messages come from Python.
rem ============================================================
chcp 65001 >nul
title Receipt Reader - Apply
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set "PY=_runtime\python\python.exe"
if not exist "%PY%" set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" ( where py >nul 2>nul && set "PY=py" )
if not exist "%PY%" if not "%PY%"=="py" (
  echo [!] Runtime not found. Please run the setup batch first.
  pause
  exit /b 1
)

echo Applying the confirmed plan ... please wait.
echo ------------------------------------------------------------
"%PY%" "src\apply_plan.py"
echo ------------------------------------------------------------
pause
