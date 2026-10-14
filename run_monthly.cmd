@echo off
chcp 65001 >nul 2>&1
REM ============================================================
REM  唯一的运行器。计划任务调它（带 --entrance scheduled），
REM  「启动」窗口点按钮起的也是同一个 run_monthly.py（--entrance manual）。
REM  手动双击也可以跑：下载上个月数据 + 生成上个月报表。
REM  参数原样传给 run_monthly.py，见它的文件头。
REM
REM  日志：logs\runs\ 下，文件名是 日期_时间_入口.log，两个窗口下方都能实时看。
REM
REM  退出码：0 成功 / 1 有需关注项 / 2 硬失败（紫鸟、没下到数据、报表没出）
REM          3 白名单 / 4 配置错误 / 5 另一轮还在跑 / 6 被中止 / 7 下载超时
REM ============================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [x] Not found: .venv\Scripts\python.exe
  echo [x] 缺少虚拟环境，请先在本目录依次执行这两条命令：
  echo     python -m venv .venv
  echo     .venv\Scripts\pip install -r requirements.txt
  exit /b 4
)

REM -u 关掉输出缓冲：计划任务超时会强杀进程，缓冲区里的日志会全丢。
".venv\Scripts\python.exe" -u run_monthly.py %*
exit /b %ERRORLEVEL%
