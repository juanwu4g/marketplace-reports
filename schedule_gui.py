# -*- coding: utf-8 -*-
"""旧的「定时设置」窗口入口 —— 现在打开统一窗口（ml_app.py）的「每月自动运行」页。

    双击 定时设置.cmd    （生产机器上的快捷方式、README 里写的都还管用）
    或   .venv\\Scripts\\pythonw schedule_gui.py

计划任务的系统调用（生成 XML、schtasks、PowerShell 查询、电脑自检）挪到了
win_tasks.py（不带 tkinter，能测）；任务该怎么登记、旧任务怎么处理的规则在
region_choice.py。这里把原来的名字转过去，引用它们的地方不用改。
"""
import sys

from win_tasks import (RUNNER, TASK, REHEARSAL, TASK_ARGS, REHEARSAL_ARGS,  # noqa: F401
                       TIME_LIMIT, MONTHS_XML, _run, ps, kv, current_user,
                       task_state, task_xml, install, machine_checks, fix_power,
                       delete_task, run_task, all_task_states, rehearse)

TAB = "schedule"


def main():
    import ml_app
    return ml_app.main(["--tab", TAB])


if __name__ == "__main__":
    sys.exit(main())
