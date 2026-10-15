# -*- coding: utf-8 -*-
"""「ML 报表」窗口的冒烟测试：用假 tkinter 把窗口搭起来、点几个按钮。

不起任何真进程：起一轮（spawn_detached）、计划任务（schtasks）、打开店铺
全都换成假的。看的是 Python 层面有没有写错，不是长相。
"""
import json
import os

import pytest

import fake_tk
from core import region_choice
from core import runlog
from core import store_config

MONTH = "2026-09"
CFG = {"default": "BOCINA_SM",
       "allowed": ["BOCINA_SM", "EWTTO_SM", "Tecnohub_CL", "Andes_CL"],
       "batch": ["BOCINA_SM", "EWTTO_SM"],
       "regions": {"Tecnohub_CL": "CL", "Andes_CL": "CL"},
       "batches": {"CL": ["Tecnohub_CL", "Andes_CL"]}}


@pytest.fixture
def app(tmp_path, monkeypatch):
    saved = fake_tk.install()
    fake_tk.ANSWERS[:] = []
    fake_tk.SHOWN[:] = []
    try:
        d = json.loads(json.dumps(CFG))
        monkeypatch.setattr(store_config, "_stores", lambda: d)
        logs = str(tmp_path / "logs")
        monkeypatch.setattr(runlog, "LOG_DIR", logs)
        from core import ml_entrance
        monkeypatch.setattr(ml_entrance, "SETTINGS", str(tmp_path / "gui_settings.json"))
        spawned = []

        class P(object):
            returncode = None

            def poll(self):
                return None

        def fake_spawn(cmd, cwd=None, stderr_path=None):
            spawned.append(cmd)
            return P()
        monkeypatch.setattr(runlog, "spawn_detached", fake_spawn)
        from core import ml_app
        monkeypatch.setattr(ml_app, "LOG_DIR", logs)
        root = fake_tk.Tk()
        a = ml_app.App(root, tab="run")
        a.spawned = spawned
        a.fake_root = root
        a.logs = logs
        yield a
    finally:
        fake_tk.uninstall(saved)


def test_window_builds_with_all_countries_by_default(app):
    assert app.fake_root.title() == "MercadoLibre 月度报表"
    assert app.choice == "ALL"
    assert app.region_box.get() == "全部国家"
    assert app.scope_var.get() == "墨西哥 2 家、智利 2 家，共 4 家"
    assert app.lb.items[0] == "BOCINA_SM　（墨西哥）"
    assert app.btn_run.cget("state") == "normal"
    assert app.stop_btn.cget("state") == "disabled"
    assert "约 1.0 小时" in app.estimate_var.get()
    app.fake_root.pump()                         # 轮询、日志查看器各跑几次不出错
    assert "还没有运行记录" in app.banner_var.get()


def test_run_all_countries_confirms_and_spawns_region_all(app):
    fake_tk.ANSWERS[:] = [True]
    app.btn_run.invoke()
    kind, title, text = fake_tk.SHOWN[-1]
    assert kind == "askokcancel" and title == "下载并出报表"
    assert "墨西哥 2 家、智利 2 家，共 4 家" in text and "约 1.0 小时" in text
    assert "顺序：墨西哥 → 智利" in text
    cmd = app.spawned[-1]
    assert cmd[2].endswith("run_monthly.py")
    assert cmd[3:] == ["--entrance", "manual", "--region", "ALL", "--month",
                       app.months[0], "--out-dir", app._out_dir()]
    # 起了之后按钮全灭，中止亮
    assert app.btn_run.cget("state") == "disabled"
    assert app.stop_btn.cget("state") == "normal"


def test_cancel_does_not_start(app):
    fake_tk.ANSWERS[:] = [False]
    app.btn_report.invoke()
    assert app.spawned == []


def test_single_country_with_a_subset(app):
    app.region_box.set("智利")
    app._pick_region()
    assert app.choice == "CL" and app.lb.items == ["Tecnohub_CL", "Andes_CL"]
    app.all_stores.set(False)
    app._toggle_stores()
    app.btn_download.invoke()                     # 一家都没选：提示，不跑
    assert fake_tk.SHOWN[-1][0] == "showwarning" and app.spawned == []
    app.lb.selection_set(1)
    fake_tk.ANSWERS[:] = [True]
    app.btn_report.invoke()
    assert "智利的 1 家：Andes_CL" in fake_tk.SHOWN[-1][2]
    assert app.spawned[-1][3:] == ["--entrance", "manual", "--region", "CL",
                                   "--step", "report", "--month", app.months[0],
                                   "--stores", "Andes_CL", "--out-dir", app._out_dir()]


def test_human_actions_tab_and_rerun(app, tmp_path):
    os.makedirs(runlog.runs_dir(app.logs), exist_ok=True)
    batch = os.path.join(app.logs, "b.json")
    with open(batch, "w", encoding="utf-8") as f:
        json.dump({"stores": [{"store": "Andes_CL", "status": "failed",
                               "missing": [{"form": "ventas", "cause": "PAGE_BLOCKED"}]}]}, f)
    rid = "20261014_020000_scheduled"
    runlog.write_json(runlog.run_paths(rid, app.logs)["json"],
                      {"run_id": rid, "month": MONTH, "region": "ALL", "stores": "all",
                       "status": "done", "exit": 1,
                       "regions": [{"region": "CL", "batch_summary": batch}]})
    open(runlog.run_paths(rid, app.logs)["log"], "w").close()
    app._refresh_human(announce=False)
    assert app.human_lb.items == ["Andes_CL（智利） —— 要登录 / 验证"]
    tab = [t for t in app.tabs.tabs_ if t[0] is app.tab_human.outer][0]
    assert "需要人工处理（1）" in tab[1]["text"]
    assert app.human_rerun.cget("state") == "normal"
    fake_tk.ANSWERS[:] = [True]
    app.human_rerun.invoke()
    assert app.spawned[-1][3:] == ["--entrance", "manual", "--region", "CL",
                                   "--month", MONTH, "--stores", "Andes_CL",
                                   "--out-dir", app._out_dir(),
                                   "--reports-for", "all"]


def test_schedule_overview_renders(app):
    st = {"DAY": "15", "TIME": "02:00", "ARGS": "--entrance scheduled --region MX",
          "RESULT": "0", "XML": ""}
    states = {region_choice.ALL_TASK: None, "MX月度报表": st, "CL月度报表": None,
              "PE月度报表": None}
    app._show_schedule(states)
    assert app.sched_head.get().startswith("● 现在用的是按国家登记的旧任务")
    var, btn = app.legacy_rows["MX"]
    assert var.get().startswith("每月 15 号 02:00") and btn.cget("state") == "normal"
    assert app.legacy_rows["CL"][0].get() == "未登记"
    assert (app.day.get(), app.hour.get()) == (15, 2)


def test_preview_lists_one_command_per_country(app, monkeypatch):
    from core import ml_app
    ran = []

    class R(object):
        stdout = "ok".encode("utf-8")

    monkeypatch.setattr(ml_app.subprocess, "run", lambda cmd, **k: ran.append(cmd) or R())

    class T(object):
        def __init__(self, target=None, daemon=None):
            self.target = target

        def start(self):
            self.target()
    monkeypatch.setattr(ml_app.threading, "Thread", T)
    app.btn_preview.invoke()
    app.fake_root.pump()
    assert [c[-2:] if "--region" in c else c[-1:] for c in ran] == [
        ["--dry-run"], ["--region", "CL"]]


def test_pick_tab_and_legacy_entrances(app):
    from core import ml_entrance
    assert ml_entrance.pick_tab([]) == "run"
    assert ml_entrance.pick_tab(["--tab", "schedule"]) == "schedule"
    assert ml_entrance.pick_tab(["--tab=human"]) == "human"
    assert ml_entrance.pick_tab(["定时"]) == "schedule"
    # 启动.cmd / 定时设置.cmd 和它们的 gui.py / schedule_gui.py 已并入 ML报表.cmd
    assert not os.path.exists(os.path.join(ml_entrance.ROOT, "gui.py"))
    assert not os.path.exists(os.path.join(ml_entrance.ROOT, "启动.cmd"))


def test_live_run_shows_in_banner_and_log_columns(app):
    rid = "20261015_020000_scheduled"
    p = runlog.run_paths(rid, app.logs)
    os.makedirs(os.path.dirname(p["log"]), exist_ok=True)
    lines = [
        "2026-10-15 02:00:00 | 定时 | 启动 | - | 信息 | 区域　　：全部国家",
        "2026-10-15 02:00:01 | 定时 | 启动 | - | 信息 | [MX] ━━ 第 1/2 个国家：墨西哥（MX），2 家店铺 ━━",
        "2026-10-15 02:00:02 | 定时 | 下载 | BOCINA_SM | 警告 | [MX]   [警告] 慢",
        "2026-10-15 03:00:01 | 定时 | 启动 | - | 信息 | [CL] ━━ 第 2/2 个国家：智利（CL），2 家店铺 ━━",
        "2026-10-15 03:00:02 | 定时 | 下载 | Andes_CL | 信息 | [CL] # [2/2] Andes_CL",
    ]
    with open(p["log"], "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    runlog.write_json(p["json"], {"run_id": rid, "entrance": "scheduled",
                                  "region": "ALL", "month": MONTH,
                                  "status": "running", "pid": os.getpid()})
    runlog.write_current({"run_id": rid, "pid": os.getpid(), "entrance": "scheduled",
                          "region": "ALL", "current_region": "CL",
                          "region_index": 2, "region_count": 2, "month": MONTH},
                         app.logs)
    app.fake_root.pump(4)
    assert not app.view.count_var.get().startswith("日志查看器出错"), app.view.count_var.get()
    assert app.view.follower.regions == ["MX", "CL"]
    assert app.view._with_region                      # 一轮多国：多出国家那一列
    shown = app.view.txt.text()
    assert "02:00:02\t墨西哥\t下载\tBOCINA_SM\t警告\t  [警告] 慢" in shown
    banner = app.banner_var.get()
    assert banner.startswith("● 正在运行（定时）：智利（CL），第 2/2 个国家")
    assert "Andes_CL（第 2/2 家）" in banner
    assert app.stop_btn.cget("state") == "normal"
    assert app.btn_run.cget("state") == "disabled"
    # 只看智利
    app.view.region.set("智利")
    app.view._rerender()
    assert "BOCINA_SM" not in app.view.txt.text() and "Andes_CL" in app.view.txt.text()
    # 中止：写请求文件
    fake_tk.ANSWERS[:] = [True]
    app.stop_btn.invoke()
    assert runlog.stop_requested(rid, app.logs)
