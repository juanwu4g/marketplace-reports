@echo off
chcp 65001 >nul 2>&1
REM ============================================================
REM  定时任务设置：把「每月自动出报表」配进 Windows 计划任务。
REM  只需要配一次。之后想看某一轮跑得怎么样（实时日志、历史记录、中止），
REM  开这个窗口或 启动.cmd 都行 —— 两边看的是同一份日志，跑的是同一个过程。
REM  日常手动跑报表请双击 启动.cmd。
REM ============================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
  echo [x] Not found: .venv\Scripts\pythonw.exe
  echo [x] 找不到虚拟环境，请先在本目录依次执行这两条命令：
  echo.
  echo     python -m venv .venv
  echo     .venv\Scripts\pip install -r requirements.txt
  echo.
  pause
  exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" schedule_gui.py
