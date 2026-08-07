@echo off
rem ============================================================
rem  Receipt Reader - Ollama setup (optional, boosts accuracy)
rem  Installs Ollama (no admin) and downloads the image model.
rem  Without this the tool still runs on OCR only (lower accuracy).
rem  ASCII-only on purpose.
rem ============================================================
chcp 65001 >nul
title Receipt Reader - Ollama setup
cd /d "%~dp0"

set "OEXE=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"

if not exist "%OEXE%" (
  echo [1/3] Installing Ollama ^(no admin needed, a few minutes^)...
  winget install --id Ollama.Ollama -e --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity
)
if not exist "%OEXE%" (
  where ollama >nul 2>nul && set "OEXE=ollama"
)
if not exist "%OEXE%" if not "%OEXE%"=="ollama" (
  echo.
  echo [!] Could not install Ollama automatically.
  echo     Please install it manually from  https://ollama.com  then run this again.
  echo.
  pause
  exit /b 1
)

echo [2/3] Starting Ollama server...
start /b "" "%OEXE%" serve >nul 2>nul

echo [3/3] Downloading model qwen2.5vl:3b ^(about 3GB, one time^)...
"%OEXE%" pull qwen2.5vl:3b

echo.
echo Done. High-accuracy mode is ready. You can now run the RUN batch file.
pause
