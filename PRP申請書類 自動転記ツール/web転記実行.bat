@echo off
chcp 65001 >nul
rem ============================================================
rem  Web auto-fill for e-Saisei (plan01)
rem   - Source = OUTPUT (02_output) first, then hearing.
rem   - This tool FILLS the form (draft). It does NOT submit.
rem   - No login required on this site. Submit yourself after checking.
rem   - Manual (JP): 00_manual folder "Web jido nyuryoku ... .md"
rem ============================================================
setlocal
cd /d "%~dp0"

set "PY=python"
where py >nul 2>nul && set "PY=py"

:menu
echo.
echo ============================================================
echo   Web Auto-Fill Tool  (e-Saisei / plan01)
echo ============================================================
echo   1 : Auto-fill (step)  (fill per tab, press Enter between tabs)
echo   A : Auto-fill (bulk)  (all tabs at once -^> temp-save) *never submits*
echo   2 : Dump fields       (--dump / to build the mapping)
echo   3 : Setup             (install Playwright, first time only)
echo   Q : Quit
echo ------------------------------------------------------------
set "sel="
set /p "sel=Type 1/A/2/3/Q and press Enter: "
if /i "%sel%"=="1" goto fill
if /i "%sel%"=="A" goto autofill
if /i "%sel%"=="2" goto dump
if /i "%sel%"=="3" goto setup
if /i "%sel%"=="Q" goto end
goto menu

:fill
echo.
echo [Auto-fill step] A browser opens (no login needed). Show the plan01 form,
echo             then come back here and press Enter.
echo             Source = OUTPUT (02_output) first. Never submits.
"%PY%" "%~dp099_data\src\web_fill.py"
goto done

:autofill
echo.
echo [Auto-fill BULK] A browser opens. Show the FIRST tab of plan01, press Enter.
echo             The tool fills ALL tabs automatically, then clicks TEMP-SAVE.
echo             If a required field is missing it stops and reports where.
echo             It NEVER clicks submit/apply. Report: 03_logs\web_autofill_report_*.txt
"%PY%" "%~dp099_data\src\web_fill.py" --auto
goto done

:dump
echo.
echo [Dump] A browser opens. Show the form, press Enter.
echo        Output: 03_logs\web_fields_dump.txt
"%PY%" "%~dp099_data\src\web_fill.py" --dump
goto done

:setup
echo.
echo [Setup] Installing Playwright and Chromium (may take 1-2 min)...
"%PY%" -m pip install playwright
"%PY%" -m playwright install chromium
goto done

:done
echo.
echo ------------------------------------------------------------
pause
goto menu

:end
endlocal
