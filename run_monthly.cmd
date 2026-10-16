@echo off
REM ---------------------------------------------------------------------------
REM  The one runner. The monthly scheduled task calls this file
REM  (--entrance scheduled --region ALL); the window starts the same
REM  run_monthly.py with --entrance manual. Arguments are passed through.
REM  Double-clicked with no arguments: last month, download + reports.
REM
REM  Logs: logs\runs\<date>_<time>_<entrance>.log (also shown in the window).
REM  Exit codes: 0 ok / 1 needs attention / 2 hard failure / 3 allowlist /
REM              4 config error / 5 another run active / 6 stopped / 7 timeout
REM
REM  ASCII on purpose (see launcher\README.txt). No arguments -> wait at the end
REM  with "timeout", not "pause": very old scheduled tasks call this file with
REM  no arguments and a pause would hang them until Windows kills the task.
REM ---------------------------------------------------------------------------
chcp 65001 >nul 2>&1
cd /d "%~dp0"
set "ML_INTERACTIVE="
if "%~1"=="" set "ML_INTERACTIVE=1"
if defined ML_INTERACTIVE title Monthly Reports - running

if not exist ".venv\Scripts\python.exe" goto novenv
if defined ML_INTERACTIVE type "launcher\run_start.txt"

REM -u: unbuffered, so a timeout kill does not lose the buffered log lines.
".venv\Scripts\python.exe" -u run_monthly.py %*
set "RC=%ERRORLEVEL%"
if not defined ML_INTERACTIVE exit /b %RC%

if "%RC%"=="0" type "launcher\run_0.txt"
if "%RC%"=="1" type "launcher\run_1.txt"
if "%RC%"=="5" type "launcher\run_5.txt"
if "%RC%"=="6" type "launcher\run_6.txt"
if not "%RC%"=="0" if not "%RC%"=="1" if not "%RC%"=="5" if not "%RC%"=="6" (
  type "launcher\run_fail.txt"
  echo   %RC%
)
type "launcher\run_close.txt"
timeout /t 600 >nul
exit /b %RC%

:novenv
type "launcher\novenv.txt"
if defined ML_INTERACTIVE timeout /t 600
exit /b 4
