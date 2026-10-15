# -*- coding: utf-8 -*-
"""统一入口（ML报表.cmd / ml_app.py）背后的规则：全部国家、计划任务怎么登记、
旧任务怎么处理、确认框写什么、按钮亮不亮、双击的 .cmd 长什么样。"""
import datetime
import io
import re
import json
import os

import pytest

import ml_entrance as E
import region_choice as RC
import run_monthly
import runlog
import store_config
import win_tasks

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MX = ["BOCINA_SM", "EWTTO_SM"]
CL = ["Tecnohub_CL", "Andes_CL"]
CFG = {"default": "BOCINA_SM", "allowed": MX + CL, "batch": MX,
       "regions": {"Tecnohub_CL": "CL", "Andes_CL": "CL"},
       "batches": {"CL": CL}}


@pytest.fixture
def cfg(monkeypatch):
    d = json.loads(json.dumps(CFG))
    monkeypatch.setattr(store_config, "_stores", lambda: d)
    return d


# ------------------------------------------------------------ 全部国家

def test_choice_labels_and_default(cfg):
    assert RC.choice_labels() == ["全部国家", "墨西哥", "智利", "秘鲁"]
    assert RC.from_label("全部国家") == "ALL"
    assert RC.label("ALL") == "全部国家"
    assert RC.remembered_choice({}) == "ALL"
    assert RC.remembered_choice({"region": "cl"}) == "CL"
    assert RC.remembered_choice({"region": "ALL"}) == "ALL"
    assert RC.remembered_choice({"region": "AR"}) == "ALL"
    assert RC.remembered({"region": "ALL"}) == "MX"        # 旧函数不变


def test_lineup_skips_countries_without_stores(cfg):
    assert RC.lineup("ALL") == [("MX", MX), ("CL", CL)]
    assert RC.lineup("CL") == [("CL", CL)]
    assert RC.scope_text(RC.lineup("ALL")) == "墨西哥 2 家、智利 2 家，共 4 家"
    assert RC.scope_text(RC.lineup("MX")) == "墨西哥 2 家"
    assert RC.scope_text([]).startswith("没有店铺")


def test_preview_runs_once_per_country(cfg):
    got = RC.preview_cmds("py", "rr.py", "2026-09", None, "ALL")
    assert got == [("MX", ["py", "rr.py", "--month", "2026-09", "--dry-run"]),
                   ("CL", ["py", "rr.py", "--month", "2026-09", "--dry-run",
                           "--region", "CL"])]


# ------------------------------------------------------------ 计划任务

def test_all_task_name_and_arguments():
    assert RC.task_name("ALL") == "ML月度报表"
    assert RC.task_name("MX") == "MX月度报表"                # 生产上的名字不动
    assert RC.task_args("ALL") == "--entrance scheduled --region ALL"
    assert RC.rehearsal_name("ALL") == "ML月度报表-试跑"
    assert RC.rehearsal_args("ALL") == "--entrance scheduled --region ALL --step report"
    a = run_monthly.parse_args(RC.task_args("ALL").split())
    assert (a.entrance, a.region, a.month, a.stores) == ("scheduled", "ALL", None, None)
    assert RC.args_region(RC.task_args("ALL")) == "ALL"
    assert RC.time_limit_hours(1) == 6 and RC.time_limit_hours(3) == 18
    assert RC.time_limit_hours(0) == 6
    assert RC.time_limit_xml(12) == "PT12H"


def _st(day="15", time="02:00", args="--entrance scheduled --region MX",
        limit="PT6H", result="0"):
    return {"DAY": day, "TIME": time, "ARGS": args, "RESULT": result,
            "NEXT": "2026/11/15 2:00:00", "LAST": "2026/10/15 2:00:00",
            "XML": "<Settings><ExecutionTimeLimit>%s</ExecutionTimeLimit></Settings>"
                   % limit}


def _states(**kw):
    d = {"ML月度报表": None, "MX月度报表": None, "CL月度报表": None, "PE月度报表": None}
    d.update(kw)
    return d


def test_overview_nothing_registered():
    ov = RC.schedule_overview(_states(), 2)
    assert ov["main"] is None and ov["legacy"] == []
    assert ov["headline"] == "● 还没设置，现在不会自动跑" and ov["kind"] == "warn"


def test_overview_production_today_only_the_mexican_task():
    """生产机器现在的样子：只有 MX月度报表。显示出来、推荐替换，绝不自动删。"""
    mx = _st()
    ov = RC.schedule_overview(_states(**{"MX月度报表": mx}), 2)
    assert ov["main"] is None
    assert ov["legacy"] == [("MX", "MX月度报表", mx)]
    assert "MX月度报表（每月 15 号 02:00）" in ov["headline"]
    assert ov["kind"] == "warn"
    assert any("ML月度报表" in t and "替换" in t for _, t in ov["notes"])
    q = RC.replace_question(ov["legacy"])
    assert "MX月度报表（墨西哥，每月 15 号 02:00）" in q
    assert "「是」" in q and "「否」" in q and "「取消」" in q
    assert RC.replace_question([]) is None


def test_overview_two_old_tasks_that_clash():
    ov = RC.schedule_overview(_states(**{
        "MX月度报表": _st(), "CL月度报表": _st(time="03:00",
                                               args="--entrance scheduled --region CL")}), 2)
    assert any(k == "bad" and "撞车" in t for k, t in ov["notes"])


def test_overview_new_task_plus_leftover_old_one_is_red():
    main = _st(args="--entrance scheduled --region ALL", limit="PT12H")
    ov = RC.schedule_overview(_states(**{"ML月度报表": main, "MX月度报表": _st()}), 2)
    assert ov["headline"] == "● 已设置：每月 15 号 02:00 自动跑全部国家"
    assert ov["kind"] == "bad"
    assert any("MX月度报表" in t and "两次" in t for k, t in ov["notes"] if k == "bad")


def test_overview_new_task_alone_is_green():
    main = _st(args="--entrance scheduled --region ALL", limit="PT12H")
    ov = RC.schedule_overview(_states(**{"ML月度报表": main}), 2)
    assert ov["kind"] == "good" and ov["notes"] == []


def test_overview_new_task_with_wrong_args_or_short_limit():
    ov = RC.schedule_overview(_states(**{"ML月度报表": _st(limit="PT12H")}), 2)
    assert any(k == "bad" and "不是全部国家" in t for k, t in ov["notes"])
    ov = RC.schedule_overview(_states(**{"ML月度报表": _st(
        args="--entrance scheduled --region ALL", limit="PT6H")}), 3)
    assert any(k == "warn" and "6 小时" in t and "18 小时" in t for k, t in ov["notes"])
    assert RC.registered_limit({"XML": ""}) is None


def test_task_verdicts_are_plain_chinese():
    assert RC.task_verdict("0") == "上次跑完，一切正常"
    assert "跳过" in RC.task_verdict(5)
    assert "代码 99" in RC.task_verdict("99")
    assert RC.task_when(None) == "未登记"


def test_task_xml_for_the_all_countries_task():
    x = win_tasks.task_xml(RC.task_args("ALL"), day=15, hour=2, minute=5,
                           time_limit=RC.time_limit_xml(12), who="PC\\meli")
    assert "<Arguments>--entrance scheduled --region ALL</Arguments>" in x
    assert "<ExecutionTimeLimit>PT12H</ExecutionTimeLimit>" in x
    assert "<Day>15</Day>" in x and "T02:05:00</StartBoundary>" in x
    assert "<LogonType>InteractiveToken</LogonType>" in x       # 紫鸟要桌面
    assert "<StartWhenAvailable>true</StartWhenAvailable>" in x
    assert x.count("<Command>") == 1 and "run_monthly.cmd</Command>" in x
    # 试跑的任务没有触发器
    r = win_tasks.task_xml(RC.rehearsal_args("ALL"), who="PC\\meli")
    assert "<Triggers></Triggers>" in r and "<ExecutionTimeLimit>PT6H<" in r


# ------------------------------------------------------------ 窗口的规则

def test_recent_months_and_estimates():
    assert E.recent_months(3, today=datetime.date(2026, 1, 20)) == [
        "2025-12", "2025-11", "2025-10"]
    assert E.run_hours(2) == "30 分钟" and E.run_hours(12) == "3.0 小时"
    assert E.estimate("both", 15) == "约 3.8 小时"
    assert E.estimate("report", 15, 2) == "约 2 分钟"
    assert "几分钟" in E.estimate("collect", 15)


def test_check_stores():
    assert E.check_stores("ALL", False, []) == (None, None)     # 全部国家没有子集
    assert E.check_stores("MX", True, ["A"]) == (None, None)
    assert E.check_stores("MX", False, ["A"]) == (["A"], None)
    sub, problem = E.check_stores("MX", False, [])
    assert sub is None and "一家都没选" in problem


def test_confirm_text_states_scope_count_and_time(cfg):
    title, text = E.confirm_text("both", "ALL", "2026-09", latest="2026-09")
    assert title == "下载并出报表"
    assert "范围：墨西哥 2 家、智利 2 家，共 4 家" in text
    assert "预计：约 1.0 小时" in text
    assert "顺序：墨西哥 → 智利" in text
    assert "不要手动去点紫鸟" in text and "窗口可以关掉" in text
    assert "不是最近一个月" not in text

    _, text = E.confirm_text("report", "CL", "2026-07", subset=["Andes_CL"],
                             latest="2026-09")
    assert "智利的 1 家：Andes_CL" in text
    assert "不出全店汇总表" in text
    assert "2026-07 不是最近一个月（2026-09）" in text and "_m202607" in text
    assert "紫鸟" not in text                    # 只出报表不开浏览器


def test_runner_command_for_each_choice():
    cmd = E.runner_cmd("both", "2026-09", "ALL", py="py", out_dir="OUT")
    assert cmd[0] == "py" and cmd[1] == "-u" and cmd[2].endswith("run_monthly.py")
    assert cmd[3:] == ["--entrance", "manual", "--region", "ALL",
                       "--month", "2026-09", "--out-dir", "OUT"]
    a = run_monthly.parse_args(cmd[3:])
    assert (a.entrance, a.region, a.step) == ("manual", "ALL", "both")
    cmd = E.runner_cmd("collect", "2026-09", "MX", py="py")
    assert cmd[3:] == ["--entrance", "manual", "--region", "MX", "--step",
                       "collect", "--month", "2026-09"]


def test_buttons_are_disabled_when_they_would_do_nothing():
    idle = E.button_states(busy=False, has_stores=True, has_human=False)
    assert idle["run"] and idle["report"] and not idle["stop"]
    assert not idle["rerun"] and not idle["open_stores"]
    busy = E.button_states(busy=True, has_stores=True, has_human=True,
                           picked_human=True)
    assert not busy["run"] and not busy["collect"] and busy["stop"]
    assert not busy["rerun"] and busy["preview"]          # 预览只读，跑着也能点
    empty = E.button_states(busy=False, has_stores=False, has_human=True,
                            picked_human=False)
    assert not empty["run"] and not empty["preview"] and not empty["rerun"]


def test_human_texts(cfg):
    h = {"month": "2026-09", "actions": [
        {"store": "Andes_CL", "title": "要登录 / 验证", "how": "去登录", "rerun": True},
        {"store": "EWTTO_SM", "title": "缺权限", "how": "开通权限", "rerun": False}]}
    assert E.human_title(h) == "需要人工处理（2）"
    assert E.human_title({"actions": []}) == "需要人工处理"
    text, kind = E.human_headline(h)
    assert kind == "bad" and "2 家店铺" in text
    assert E.human_headline({"actions": []})[1] == "good"
    assert "（处理之前重跑没用）" in E.human_detail(h["actions"])
    ok, title, text = E.rerun_confirm(RC.rerun_plan(h["actions"]), "2026-09")
    assert ok is True and "Andes_CL" in text and "智利全部店铺" in text
    mixed = [dict(h["actions"][0]), {"store": "BOCINA_SM", "title": "x", "how": "",
                                     "rerun": True}]
    ok, _, text = E.rerun_confirm(RC.rerun_plan(mixed), "2026-09")
    assert ok is False and "分属不同的国家" in text


def test_status_line(cfg):
    assert E.status_line(None)[0].startswith("还没有运行记录")
    st = {"state": "finished", "kind": "warn", "entrance": "scheduled",
          "region": "ALL", "month": "2026-09", "text": "跑完了，但有需要关注的项"}
    text, kind = E.status_line(st)
    assert text == "⚠ 最近一轮（定时，全部国家 2026-09）：跑完了，但有需要关注的项"
    assert kind == "warn"
    run = {"state": "running", "entrance": "manual", "run_id": "20261015_020000_manual",
           "month": "2026-09", "summary": {"region": "CL"}, "kind": "run"}
    text, kind = E.status_line(run, now=datetime.datetime(2026, 10, 15, 2, 0, 9))
    assert text == "● 正在运行（手动）：智利（CL）  ·  2026-09  ·  已运行 0:00:09"
    stalled = dict(run, stalled=True, text="没有响应")
    assert E.status_line(stalled)[1] == "bad"


def test_settings_roundtrip_and_tabs(tmp_path):
    p = str(tmp_path / "s.json")
    assert E.load_settings(p) == {}
    E.update_settings(p, region="ALL")
    E.update_settings(p, out_dir="D:/x")
    assert E.load_settings(p) == {"region": "ALL", "out_dir": "D:/x"}
    assert E.pick_tab(["--tab", "sched"]) == "schedule"
    assert E.pick_tab(["--tab", "whatever"]) == "run"


def test_self_check_explains_instead_of_a_traceback(monkeypatch):
    import builtins
    real = builtins.__import__

    def no_tk(name, *a, **k):
        if name == "tkinter":
            raise ImportError("No module named 'tkinter'")
        return real(name, *a, **k)
    monkeypatch.setattr(builtins, "__import__", no_tk)
    out = io.StringIO()
    assert E.self_check(out) == 1
    text = out.getvalue()
    assert "界面库 tkinter" in text and "怎么办" in text
    assert "Traceback" not in text


# ------------------------------------------------------------ 控制台

def test_console_line_is_short_and_keeps_the_level():
    line = runlog.format_line("2026-10-13 14:05:01", "scheduled", "下载",
                              "BOCINA_TA02", "警告", "没有出现新行")
    assert runlog.console_line(line) == "14:05:01  下载  BOCINA_TA02  [警告]  没有出现新行"
    info = runlog.format_line("2026-10-13 14:05:01", "manual", "启动", "-", "信息", "x")
    assert runlog.console_line(info) == "14:05:01  启动  x"
    assert runlog.console_line("随便一行") == "随便一行"


# ------------------------------------------------------------ 双击的 .cmd

def _cmd(name):
    with io.open(os.path.join(ROOT, name), "rb") as f:
        raw = f.read()
    return raw, raw.decode("utf-8")


LAUNCHERS = ["ML报表.cmd", "启动.cmd", "定时设置.cmd", "run_monthly.cmd"]


@pytest.mark.parametrize("name", LAUNCHERS)
def test_launchers_are_ascii_and_their_messages_exist(name):
    # 2026-10-15：cmd.exe 在文件中途 chcp 65001 之后会把含中文的行切错（没有
    # 真正控制台时 "'M' 不是内部或外部命令"、run_monthly.py 根本没起来）。
    # .cmd 只用 ASCII，中文提示放 launcher/*.txt 用 type 显示。
    raw, text = _cmd(name)
    assert b"\r\n" in raw and b"\n" not in raw.replace(b"\r\n", b""), "批处理必须 CRLF"
    assert all(b < 0x80 for b in raw), "%s 里有非 ASCII 字符" % name
    assert "chcp 65001" in text
    for ref in re.findall(r'type "(launcher\\[^"]+)"', text):
        assert os.path.exists(os.path.join(ROOT, ref.replace("\\", os.sep))), ref


@pytest.mark.parametrize("name,tab", [("ML报表.cmd", "run"), ("启动.cmd", "run"),
                                      ("定时设置.cmd", "schedule")])
def test_launchers_open_the_one_window_on_the_right_tab(name, tab):
    raw, text = _cmd(name)
    lines = text.splitlines()
    assert lines[0] == "@echo off"
    assert any(l.startswith("title ") for l in lines)
    assert 'ml_app.py --check' in text                    # 先自检，出错才停住
    assert 'start "" ".venv\\Scripts\\pythonw.exe" ml_app.py --tab %s' % tab in text
    assert "pause" in text and ":broken" in text and ":novenv" in text


def test_runner_cmd_never_pauses_for_the_scheduled_task():
    raw, text = _cmd("run_monthly.cmd")
    assert '".venv\\Scripts\\python.exe" -u run_monthly.py %*' in text
    assert "exit /b %RC%" in text
    # 从不 pause（很早以前的计划任务不带参数调它，pause 会让任务挂到超时）。
    # 带参数（计划任务）时，跑完立刻 exit，走不到后面的 timeout。
    code = [l.strip().lower() for l in text.splitlines()
            if l.strip() and not l.strip().lower().startswith("rem")]
    assert not any("pause" in c.split() for c in code)
    run = code.index('".venv\\scripts\\python.exe" -u run_monthly.py %*')
    assert code[run + 2] == "if not defined ml_interactive exit /b %rc%"
    for i, c in enumerate(code):
        if "timeout /t" in c and i < code.index(":novenv"):
            assert i > run + 2, c
        elif "timeout /t" in c:
            assert c.startswith("if defined ml_interactive"), c
    assert 'if "%~1"=="" set "ML_INTERACTIVE=1"' in text


def test_one_powershell_call_reads_every_task():
    out = ("BEGIN=ML月度报表\nEXISTS=0\n"
           "BEGIN=MX月度报表\nEXISTS=1\nSTATE=Ready\nNEXT=2026/11/15 2:00:00\n"
           "LAST=2026/10/15 2:00:00\nRESULT=1\nXML=<Task> <CalendarTrigger>"
           "<StartBoundary>2026-09-01T02:00:00</StartBoundary> <Day>15</Day>"
           "</CalendarTrigger> <Arguments>--entrance scheduled</Arguments> </Task>\n")
    got = win_tasks.split_states(out)
    assert got["ML月度报表"] is None
    mx = got["MX月度报表"]
    assert (mx["DAY"], mx["TIME"], mx["RESULT"]) == ("15", "02:00", "1")
    assert RC.args_region(mx["ARGS"]) == "MX"            # 旧任务没有 --region
    script = []
    orig = win_tasks.ps
    try:
        win_tasks.ps = lambda s, timeout=120: (script.append(s) or (0, out))
        states = win_tasks.task_states(["ML月度报表", "MX月度报表", "CL月度报表"])
    finally:
        win_tasks.ps = orig
    assert len(script) == 1 and script[0].count("BEGIN=") == 3
    assert states["CL月度报表"] is None and states["MX月度报表"]["DAY"] == "15"
