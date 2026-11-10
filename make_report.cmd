@echo off
rem SW Status Report - weekly status deck from the input workbook (R-OP-2).
rem Drag the input workbook (.xlsx) onto this file. Extra options can be added on a
rem command line, e.g.:  make_report.cmd Tasks.xlsx --check
rem Set SWSR_NO_PAUSE=1 to skip the final pause (for scripts and tests).
setlocal
set "PYZ=%~dp0statusreport.pyz"

if "%~1"=="" (
    echo SW Status Report
    echo.
    echo Drag your input workbook ^(.xlsx^) onto make_report.cmd to create this week's deck.
    echo An empty input workbook is Tasks.xlsx in this folder - see README.md.
    set "RC=2"
    goto :end
)

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo Python 3 was not found. Please install Python 3.9 or newer, or ask your IT.
    set "RC=2"
    goto :end
)

%PY% "%PYZ%" %*
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" echo Done.
if "%RC%"=="1" echo Done - but some rows had errors and were skipped, see above.
if not "%RC%"=="0" if not "%RC%"=="1" echo Nothing was written - see the message above.

:end
if not defined SWSR_NO_PAUSE pause
exit /b %RC%
