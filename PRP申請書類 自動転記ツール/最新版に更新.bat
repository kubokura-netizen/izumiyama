@echo off
rem ============================================================
rem  Patcher launcher - double-click this file to update the tool.
rem  It downloads the latest version from GitHub and applies it.
rem  Your data (01_input / 02_output / 03_logs), the bundled
rem  Python runtime (_runtime) and login profile are kept as-is.
rem
rem  All logic and Japanese messages are in 98_dashboard\_update.ps1
rem  (PowerShell handles UTF-8 correctly; this .bat stays ASCII).
rem ============================================================
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\98_dashboard\_update.ps1" -ToolDir "%CD%"
endlocal
