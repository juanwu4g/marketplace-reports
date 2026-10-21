@echo off
REM ---------------------------------------------------------------------------
REM  Double-click to update this machine to the latest code (tools\update.py):
REM  refuses during a monthly run or with local code edits, then git pull,
REM  pip install (UTF-8 mode), bootstrap self-check and a dry run of every
REM  market. Opens no store. ASCII only: Chinese text is in launcher\*.txt.
REM ---------------------------------------------------------------------------
chcp 65001 >nul 2>&1
cd /d "%~dp0"
title Update
if not exist ".venv\Scripts\python.exe" goto novenv
set PYTHONUTF8=1
".venv\Scripts\python.exe" -m tools.update %*
if errorlevel 1 goto failed
type "launcher\update_ok.txt"
pause
exit /b 0

:novenv
type "launcher\novenv.txt"
pause
exit /b 1

:failed
type "launcher\update_fail.txt"
pause
exit /b 1
