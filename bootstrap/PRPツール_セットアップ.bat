@echo off
rem ============================================================
rem  PRP tool - one-file setup / update.
rem  Double-click this file. It downloads the installer from
rem  GitHub and runs it. Nothing is bundled in this file.
rem
rem  First run  : installs the whole tool (code + Python + parts).
rem  Next runs  : updates to the latest (your data is kept).
rem
rem  This .bat is ASCII on purpose; all Japanese messages live in
rem  bootstrap/install.ps1 (UTF-8 with BOM) fetched from GitHub.
rem ============================================================
setlocal
title PRP tool setup
set "URL=https://raw.githubusercontent.com/kubokura-netizen/izumiyama/main/bootstrap/install.ps1"
set "DST=%TEMP%\prp_install.ps1"

echo.
echo Downloading the installer from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -Uri $env:URL -OutFile $env:DST -UseBasicParsing -TimeoutSec 120 } catch { exit 1 }"

if not exist "%DST%" (
  echo.
  echo [ERROR] Could not download the installer.
  echo         Please check your internet connection and try again.
  echo.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%DST%"
del "%DST%" >nul 2>nul
endlocal
