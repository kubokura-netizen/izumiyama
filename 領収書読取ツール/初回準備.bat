@echo off
rem ============================================================
rem  Receipt Reader - SETUP (first time only)
rem  Creates a Python env (.venv) and installs the parts.
rem  Bundled OCR (Tesseract) is in _ocr\tesseract (no admin).
rem  ASCII-only on purpose so cmd never mis-parses it.
rem ============================================================
chcp 65001 >nul
title Receipt Reader - Setup
cd /d "%~dp0"
set PYTHONUTF8=1

set "PYEXE=py"
where py >nul 2>nul || set "PYEXE=python"
where %PYEXE% >nul 2>nul || (
  echo [!] Python not found. Install it from https://www.python.org and retry.
  pause
  exit /b 1
)

echo [1/4] Creating Python environment (.venv) ...
if not exist ".venv\Scripts\python.exe" %PYEXE% -m venv .venv
set "PY=.venv\Scripts\python.exe"

echo [2/4] Installing parts (a few minutes) ...
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install pymupdf opencv-python-headless pytesseract Pillow openpyxl numpy

echo [3/4] Checking bundled OCR (Tesseract) ...
if exist "_ocr\tesseract\tesseract.exe" (echo    OCR Tesseract: OK) else (echo    OCR Tesseract: MISSING - see README)

echo [4/4] Checking Ollama (optional image LLM) ...
where ollama >nul 2>nul && (echo    Ollama: installed) || (echo    Ollama: not installed [optional, boosts accuracy])

echo.
echo Setup done. Now run the RUN batch to read receipts.
pause
