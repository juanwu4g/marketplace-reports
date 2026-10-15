# -*- coding: utf-8 -*-
"""旧的「启动」窗口入口 —— 现在打开统一窗口（ml_app.py）的「立即运行」页。

    双击 启动.cmd     （生产机器上的快捷方式、README 里写的都还管用）
    或   .venv\\Scripts\\pythonw gui.py

以前这里是一个单独的窗口（选月份、选店铺、需要人工处理、运行日志），定时
任务在另一个窗口（schedule_gui.py）。两边合进了一个窗口：ML报表.cmd /
ml_app.py。这个文件只留两样东西：

* main()：打开统一窗口，停在「立即运行」页；
* 原来从这里 import 的几个函数（别的脚本、旧习惯还在用），转到 ml_entrance。
"""
import sys

import region_choice
from ml_entrance import (ROOT, DEFAULT_REPORT_DIR, SETTINGS, RUNNER,  # noqa: F401
                         load_settings, save_settings, interpreter, run_hours,
                         recent_months)

TAB = "run"


def store_list(region=None):
    """店铺白名单。读不到就返回空。给了 region 就是那个国家的"全部店铺"。"""
    if region:
        return region_choice.stores(region)
    try:
        sys.path.insert(0, ROOT)
        import store_config
        return list(store_config.batch_stores())
    except Exception:
        return []


def main():
    import ml_app
    return ml_app.main(["--tab", TAB])


if __name__ == "__main__":
    sys.exit(main())
