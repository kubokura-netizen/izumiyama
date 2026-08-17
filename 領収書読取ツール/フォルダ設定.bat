@echo off
rem ============================================================
rem  Receipt Reader - CONFIGURE folders (input / output / master)
rem  Interactive. Writes settings.json to match this PC.
rem  ASCII-only header; Japanese prompts come from Python.
rem ============================================================
chcp 65001 >nul
title Receipt Reader - Configure
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

"%PY%" "src\configure.py"
pause
