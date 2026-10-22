# -*- coding: utf-8 -*-
"""「月度报表」窗口的冒烟测试：用假 tkinter 把窗口搭起来、点几个按钮。

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


# three platforms: MercadoLibre Mexico + Chile, Falabella Chile, Amazon Mexico
MULTI = {"default": "BOCINA_SM",
         "allowed": ["BOCINA_SM", "EWTTO_SM", "Andes_CL", "AIS_CL-Fala", "DAB_CL-Fala",
                     "SM-Amazon"],
         "catalog": {"BOCINA_SM": {"platform": "ML", "country": "MX"},
                     "EWTTO_SM": {"platform": "ML", "country": "MX"},
                     "Andes_CL": {"platform": "ML", "country": "CL"},
                     "AIS_CL-Fala": {"platform": "FA", "country": "CL"},
                     "DAB_CL-Fala": {"platform": "FA", "country": "CL"},
                     "SM-Amazon": {"platform": "AZ", "country": "MX"}},
         "batches": {"ML-MX": ["BOCINA_SM", "EWTTO_SM"], "ML-CL": ["Andes_CL"],
                     "FA-CL": ["AIS_CL-Fala", "DAB_CL-Fala"], "AZ-MX": ["SM-Amazon"]}}


@pytest.fixture
def app(tmp_path, monkeypatch):
    yield from _app(tmp_path, monkeypatch, CFG)


@pytest.fixture
def app3(tmp_path, monkeypatch):
    yield from _app(tmp_path, monkeypatch, MULTI)


def _app(tmp_path, monkeypatch, cfg):
    saved = fake_tk.install()
    fake_tk.ANSWERS[:] = []
    fake_tk.SHOWN[:] = []
    try:
        d = json.loads(json.dumps(cfg))
        monkeypatch.setattr(store_config, "_stores", lambda: d)
        logs = str(tmp_path / "logs")
        monkeypatch.setattr(runlog, "LOG_DIR", logs)
        from core import entrance
        monkeypatch.setattr(entrance, "SETTINGS", str(tmp_path / "gui_settings.json"))
        spawned = []

        class P(object):
            returncode = None

            def poll(self):
                return None

        def fake_spawn(cmd, cwd=None, stderr_path=None):
            spawned.append(cmd)
            return P()
        monkeypatch.setattr(runlog, "spawn_detached", fake_spawn)
        from core import app as app_mod
        monkeypatch.setattr(app_mod, "LOG_DIR", logs)
        root = fake_tk.Tk()
        a = app_mod.App(root, tab="run")
        a.spawned = spawned
        a.fake_root = root
        a.logs = logs
        yield a
    finally:
        fake_tk.uninstall(saved)


def test_window_builds_with_all_countries_by_default(app):
    assert app.fake_root.title() == "月度报表"
    tv = app.run_picker.tv
    # 平台 → 国家 → 店铺，全都勾着
    assert tv.get_children("") == ("p:ML",)
    assert tv.get_children("p:ML") == ("u:MX", "u:CL")
    assert tv.get_children("u:CL") == ("s:CL:Tecnohub_CL", "s:CL:Andes_CL")
    marks = app.run_picker.marks
    assert tv.text("p:ML") == " 美客多 MercadoLibre" and tv.text("u:MX") == " 墨西哥"
    assert marks["p:ML"] == marks["u:MX"] == marks["s:CL:Andes_CL"] == "all"
    assert app.scope_var.get() == "已选：美客多：墨西哥 2 家、智利 2 家 —— 共 4 家"
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
    assert "美客多 墨西哥：全部 2 家" in text and "约 1.0 小时" in text
    assert "顺序：美客多 墨西哥 → 美客多 智利" in text
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
    app.run_picker.set_all(False)
    assert app.estimate_var.get() == "一家店都没勾"
    # 一家都没勾：按钮灰掉，点了也不跑
    assert app.btn_run.cget("state") == "disabled"
    assert app.btn_download.cget("state") == "disabled"
    app.btn_download.invoke()
    assert app.spawned == []
    app.run_picker.toggle("s:CL:Andes_CL")
    tv = app.run_picker.tv
    marks = app.run_picker.marks
    assert marks["u:CL"] == marks["p:ML"] == "some" and marks["u:MX"] == "none"
    assert marks["s:CL:Andes_CL"] == "all" and marks["s:CL:Tecnohub_CL"] == "none"
    assert tv.item("u:CL")["values"][0] == "1/2 家"
    fake_tk.ANSWERS[:] = [True]
    app.btn_report.invoke()
    assert "美客多 智利：只跑 1 家（Andes_CL）" in fake_tk.SHOWN[-1][2]
    assert app.spawned[-1][3:] == ["--entrance", "manual", "--region", "CL",
                                   "--step", "report", "--month", app.months[0],
                                   "--stores", "Andes_CL", "--out-dir", app._out_dir()]


def test_whole_country_plus_part_of_another(app):
    app.run_picker.set_all(False)
    app.run_picker.toggle("u:MX")                 # 墨西哥全部
    app.run_picker.toggle("s:CL:Tecnohub_CL")     # 智利只要一家
    fake_tk.ANSWERS[:] = [True]
    app.btn_run.invoke()
    assert app.spawned[-1][3:7] == ["--entrance", "manual", "--region", "MX,CL"]
    assert app.spawned[-1][-4:-2] == ["--stores", "Tecnohub_CL"]
    # 平台那一行：点一下从部分 -> 全部
    app.run_picker.toggle("p:ML")
    assert app.scope_var.get().endswith("共 4 家")


def _human_run(app, stores):
    """一轮全部国家的记录：每个 (区域, 店) 都要登录。"""
    os.makedirs(runlog.runs_dir(app.logs), exist_ok=True)
    regions = []
    for region, names in stores:
        batch = os.path.join(app.logs, "b_%s.json" % region)
        with open(batch, "w", encoding="utf-8") as f:
            json.dump({"stores": [{"store": n, "status": "failed",
                                   "missing": [{"form": "ventas", "cause": "PAGE_BLOCKED"}]}
                                  for n in names]}, f)
        regions.append({"region": region, "batch_summary": batch})
    rid = "20261014_020000_scheduled"
    runlog.write_json(runlog.run_paths(rid, app.logs)["json"],
                      {"run_id": rid, "month": MONTH, "region": "ALL", "stores": "all",
                       "status": "done", "exit": 1, "regions": regions})
    open(runlog.run_paths(rid, app.logs)["log"], "w").close()
    app._refresh_human(announce=False)


def test_human_actions_tab_and_rerun(app, tmp_path):
    _human_run(app, [("MX", ["BOCINA_SM"]), ("CL", ["Andes_CL"])])
    tv, hp = app.human_picker.tv, app.human_picker
    assert tv.get_children("p:ML") == ("u:MX", "u:CL")
    assert tv.get_children("u:CL") == ("s:CL:Andes_CL",)
    assert tv.item("s:CL:Andes_CL")["values"][1] == "要登录 / 验证"
    tab = [t for t in app.tabs.tabs_ if t[0] is app.tab_human.outer][0]
    assert "需要人工处理（2）" in tab[1]["text"]
    # nothing ticked at first: both buttons off, the hint says to tick
    assert hp.picked_stores() == []
    assert app.human_open.cget("state") == "disabled"
    assert app.human_rerun.cget("state") == "disabled"
    assert "勾选" in app.human_how.get()
    # tick only the Chilean store: buttons act on that one
    hp.toggle("s:CL:Andes_CL")
    assert app.human_rerun.cget("state") == "normal"
    assert app.human_open.cget("text") == "① 打开勾选的店铺（1 家）"
    assert app.human_rerun.cget("text").endswith("（1 家）")
    fake_tk.ANSWERS[:] = [True]
    app.human_rerun.invoke()
    assert app.spawned[-1][3:] == ["--entrance", "manual", "--region", "CL",
                                   "--month", MONTH, "--stores", "Andes_CL",
                                   "--out-dir", app._out_dir(),
                                   "--reports-for", "all"]
    # the first re-run is under way: buttons are off until it ends
    assert app.human_rerun.cget("state") == "disabled"
    app.launch = None
    app._lock(False, force=True)
    # a refresh keeps the tick; new stores stay unticked
    app._refresh_human(announce=False)
    assert hp.picked_stores() == ["Andes_CL"]
    # 全选 ticks both, across countries -> one run over both
    hp.set_all(True)
    fake_tk.ANSWERS[:] = [True]
    app.human_rerun.invoke()
    assert app.spawned[-1][3:7] == ["--entrance", "manual", "--region", "MX,CL"]


def test_open_only_the_ticked_stores(app, monkeypatch):
    from core import app as app_mod
    _human_run(app, [("MX", ["BOCINA_SM", "EWTTO_SM"])])
    opened = []

    class R(object):
        returncode = 0

    monkeypatch.setattr(app_mod.subprocess, "run",
                        lambda cmd, **k: opened.append(cmd[cmd.index("--store") + 1]) or R())

    class T(object):
        def __init__(self, target=None, daemon=None):
            self.target = target

        def start(self):
            self.target()
    monkeypatch.setattr(app_mod.threading, "Thread", T)
    app.human_picker.toggle("s:MX:EWTTO_SM")
    app.human_open.invoke()
    app.fake_root.pump()
    assert opened == ["EWTTO_SM"]


def test_schedule_overview_renders(app):
    st = {"DAY": "15", "TIME": "02:00", "ARGS": "--entrance scheduled --region MX",
          "RESULT": "0", "XML": ""}
    states = {region_choice.ALL_TASK: None, "MX月度报表": st, "CL月度报表": None,
              "PE月度报表": None}
    app._show_schedule(states)
    assert app.sched_head.get().startswith("● 现在用的是按国家登记的旧任务")
    var, btn = app.legacy_rows["MX月度报表"]
    assert var.get().startswith("每月 15 号 02:00") and btn.cget("state") == "normal"
    assert app.legacy_rows["CL月度报表"][0].get() == "未登记"
    assert (app.days.get(), app.hour.get()) == ("15", 2)
    # only old tasks that exist on this machine are shown
    assert app.legacy_frames["MX月度报表"].winfo_ismapped()
    assert not app.legacy_frames["CL月度报表"].winfo_ismapped()
    assert not app.legacy_none.winfo_ismapped()
    assert app.sched_detail.get().startswith("勾选：")   # 新任务还没登记：说的是勾选
    app._show_schedule({region_choice.ALL_TASK: dict(st, ARGS="--entrance scheduled --region ALL")})
    assert app.sched_detail.get().startswith("会跑：美客多：墨西哥 2 家、智利 2 家")
    assert not app.legacy_frames["MX月度报表"].winfo_ismapped()
    assert app.legacy_none.winfo_ismapped()


def test_preview_lists_one_command_per_country(app, monkeypatch):
    from core import app as app_mod
    ran = []

    class R(object):
        stdout = "ok".encode("utf-8")

    monkeypatch.setattr(app_mod.subprocess, "run", lambda cmd, **k: ran.append(cmd) or R())

    class T(object):
        def __init__(self, target=None, daemon=None):
            self.target = target

        def start(self):
            self.target()
    monkeypatch.setattr(app_mod.threading, "Thread", T)
    app.btn_preview.invoke()
    app.fake_root.pump()
    assert [c[-2:] if "--region" in c else c[-1:] for c in ran] == [
        ["--dry-run"], ["--region", "CL"]]


def test_pick_tab_and_legacy_entrances(app):
    from core import entrance
    assert entrance.pick_tab([]) == "run"
    assert entrance.pick_tab(["--tab", "schedule"]) == "schedule"
    assert entrance.pick_tab(["--tab=human"]) == "human"
    assert entrance.pick_tab(["定时"]) == "schedule"
    # 启动.cmd / 定时设置.cmd 和它们的 gui.py / schedule_gui.py 已并入 月度报表.cmd
    assert not os.path.exists(os.path.join(entrance.ROOT, "gui.py"))
    assert not os.path.exists(os.path.join(entrance.ROOT, "启动.cmd"))


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


# ------------------------------------------------------------ 三个平台

def _sync_threads(monkeypatch):
    from core import app as app_mod

    class T(object):
        def __init__(self, target=None, daemon=None):
            self.target = target

        def start(self):
            self.target()
    monkeypatch.setattr(app_mod.threading, "Thread", T)


def test_three_platforms_are_grouped_platform_first(app3):
    tv = app3.run_picker.tv
    assert tv.get_children("") == ("p:ML", "p:FA", "p:AZ")
    assert [tv.text(i) for i in tv.get_children("")] == [
        " 美客多 MercadoLibre", " Falabella", " 亚马逊 Amazon"]
    assert tv.get_children("p:FA") == ("u:FA-CL",)
    assert tv.item("u:FA-CL")["values"][0] == "2 家"
    # the schedule tab has the same tree, countries only (no store rows)
    st = app3.sched_picker.tv
    assert st.get_children("p:AZ") == ("u:AZ-MX",) and st.get_children("u:AZ-MX") == ()


def test_run_now_over_two_platforms(app3):
    app3.run_picker.set_all(False)
    app3.run_picker.toggle("p:FA")
    app3.run_picker.toggle("u:AZ-MX")
    fake_tk.ANSWERS[:] = [True]
    app3.btn_run.invoke()
    text = fake_tk.SHOWN[-1][2]
    assert "Falabella 智利：全部 2 家" in text and "亚马逊 墨西哥：全部 1 家" in text
    assert "顺序：Falabella 智利 → 亚马逊 墨西哥" in text
    assert app3.spawned[-1][3:7] == ["--entrance", "manual", "--region", "FA-CL,AZ-MX"]
    assert "--stores" not in app3.spawned[-1]


def test_schedule_follows_the_ticks_and_the_registered_task(app3, monkeypatch):
    from core import app as app_mod
    _sync_threads(monkeypatch)
    got = []
    monkeypatch.setattr(app_mod.W, "install", lambda name, xml: got.append((name, xml)) or (0, ""))
    app3._show_schedule({region_choice.ALL_TASK: None})
    # everything ticked: the same task production already has (--region ALL)
    fake_tk.ANSWERS[:] = [True]
    app3.do_install()
    app3.fake_root.pump()
    assert got[-1][0] == "月度报表" and "--region ALL" in got[-1][1]
    assert "以后新加的平台 / 国家也会自动跑" in fake_tk.SHOWN[-1][2]
    # untick Amazon: the same task, an explicit list
    app3.sched_picker.toggle("u:AZ-MX")
    assert "改过了" in app3.sched_scope_var.get()
    fake_tk.ANSWERS[:] = [True]
    app3.do_install()
    app3.fake_root.pump()
    assert got[-1][0] == "月度报表" and "--region MX,CL,FA-CL" in got[-1][1]
    assert "只跑勾上的" in fake_tk.SHOWN[-1][2]
    assert "改过了" not in app3.sched_scope_var.get()
    # a registered list shows up as ticks and in the page's notes
    reg = {"DAY": "15", "TIME": "02:00", "RESULT": "0", "XML": "",
           "ARGS": "--entrance scheduled --region FA-CL"}
    app3._show_schedule({region_choice.ALL_TASK: reg})
    assert dict(app3.sched_picker.selection()) == {"FA-CL": None}
    assert app3.sched_detail.get().startswith("会跑：Falabella：智利 2 家")
    assert app3.sched_head.get().endswith("自动跑勾选的平台和国家")


def test_last_run_column(app3, tmp_path):
    rid = "20261015_020000_scheduled"
    runlog.write_json(runlog.run_paths(rid, app3.logs)["json"],
                      {"run_id": rid, "region": "ALL", "status": "done", "exit": 1,
                       "regions": [{"region": "FA-CL", "exit": 1, "status": "done"}]})
    open(runlog.run_paths(rid, app3.logs)["log"], "w").close()
    app3._on_run_state({"state": "finished"})
    assert app3.run_picker.tv.item("u:FA-CL")["values"][1] == "10-15 ⚠ 需关注"
    assert app3.sched_picker.tv.item("u:FA-CL")["values"][1] == "10-15 ⚠ 需关注"
