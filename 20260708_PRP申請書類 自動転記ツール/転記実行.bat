@echo off
chcp 65001 >nul
rem ============================================================
rem  PRP application docs - transcribe launcher (double click)
rem  1) Put the hearing sheet (.xlsx) into 01_input
rem  2) Double-click this file
rem  3) Completed docs appear in 02_output (all sheets kept)
rem ============================================================
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    py "%~dp099_data\src\transcribe.py" %*
) else (
    python "%~dp099_data\src\transcribe.py" %*
)

echo.
echo ------------------------------------------------------------
echo  Output : 02_output      (Log : 03_logs)
echo ------------------------------------------------------------
pause
endlocal
