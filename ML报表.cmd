@echo off
REM ---------------------------------------------------------------------------
REM  Double-click to open the "ML Reports" window (ml_app.py) on tab: run.
REM  ML_Reports / start / schedule launchers all open the same window.
REM
REM  This file is ASCII on purpose: cmd.exe mis-reads non-ASCII lines that come
REM  after "chcp 65001" in the same file. User-facing Chinese text lives in
REM  launcher\*.txt and is shown with "type".
REM
REM  The interpreter is pinned to .venv: other Pythons on the machine may carry
REM  a different pandas and break the reports. "--check" runs first with
REM  python.exe so a broken setup shows a reason instead of "nothing happened".
REM ---------------------------------------------------------------------------
chcp 65001 >nul 2>&1
cd /d "%~dp0"
title ML Reports - MercadoLibre
if not exist ".venv\Scripts\pythonw.exe" goto novenv
if not exist "logs" mkdir "logs" >nul 2>&1

type "launcher\opening.txt"
".venv\Scripts\python.exe" ml_app.py --check >"logs\_window_check.txt" 2>&1
if errorlevel 1 goto broken

start "" ".venv\Scripts\pythonw.exe" ml_app.py --tab run
exit /b 0

:novenv
type "launcher\novenv.txt"
pause
exit /b 1

:broken
type "launcher\broken.txt"
type "logs\_window_check.txt"
type "launcher\broken_tail.txt"
pause
exit /b 1
