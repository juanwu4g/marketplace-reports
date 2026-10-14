# -*- coding: utf-8 -*-
"""runlog：统一日志格式、逐行分类、跟读、运行记录。全部不需要 tkinter。"""
import datetime
import io
import os
import time

import pytest

import runlog
from runlog import INFO, WARN, ERROR


# ------------------------------------------------------------ 格式 / 解析

def test_format_matches_the_documented_example():
    line = runlog.format_line(
        datetime.datetime(2026, 10, 13, 14, 5, 1), "scheduled", "下载",
        "BOCINA_TA02", "警告",
        "没有出现新行，将按 '9 agosto 2026 - 8 octubre 2026' 匹配")
    assert line == ("2026-10-13 14:05:01 | 定时 | 下载 | BOCINA_TA02 | 警告 | "
                    "没有出现新行，将按 '9 agosto 2026 - 8 octubre 2026' 匹配")


@pytest.mark.parametrize("entrance,zh", [("manual", "手动"),
                                         ("scheduled", "定时"),
                                         ("cli", "命令行")])
def test_round_trip_keeps_every_field(entrance, zh):
    # 正文里自己带 " | "，还有前导缩进 —— 都不能丢。
    msg = "    路线：all | 白名单：12 家 | x"
    line = runlog.format_line("2026-10-13 02:00:00", entrance, "启动", None,
                              INFO, msg)
    rec = runlog.parse_line(line + "\r\n")
    assert rec == {"time": "2026-10-13 02:00:00", "entrance": zh,
                   "step": "启动", "store": "-", "level": "信息",
                   "message": msg, "raw": line}


def test_parse_rejects_foreign_lines_but_viewer_still_shows_them():
    assert runlog.parse_line("Traceback (most recent call last):") is None
    assert runlog.parse_line("a | b | c | d | e | f") is None
    rec = runlog.as_record("随便一行\n")
    assert rec["message"] == "随便一行" and rec["level"] == INFO


def test_filter_by_store_and_level():
    recs = [runlog.as_record(runlog.format_line("2026-10-13 02:00:00", "cli",
                                                "下载", s, lv, "x"))
            for s, lv in [("A", INFO), ("A", WARN), ("B", ERROR), (None, INFO)]]
    assert [runlog.matches(r, "A") for r in recs] == [True, True, False, False]
    assert [runlog.matches(r, None, WARN) for r in recs] == [False, True, True, False]
    assert [runlog.matches(r, "A", ERROR) for r in recs] == [False] * 4


# ------------------------------------------------------------ run id / 文件布局

def test_run_id_and_file_layout(tmp_path):
    when = datetime.datetime(2026, 10, 13, 14, 5, 1)
    rid = runlog.make_run_id("scheduled", when, str(tmp_path))
    assert rid == "20261013_140501_scheduled"
    assert rid.encode("ascii")
    p = runlog.run_paths(rid, str(tmp_path))
    runs = os.path.join(str(tmp_path), "runs")
    assert p == {"log": os.path.join(runs, rid + ".log"),
                 "json": os.path.join(runs, rid + ".json"),
                 "stop": os.path.join(runs, rid + ".stop")}
    assert runlog.current_path(str(tmp_path)) == os.path.join(str(tmp_path),
                                                              "current.json")
    assert runlog.parse_run_id(rid) == (when, "scheduled")
    assert runlog.parse_run_id("monthly_202609") is None
    with pytest.raises(ValueError):
        runlog.make_run_id("gui", when, str(tmp_path))


def test_run_id_does_not_collide_within_one_second(tmp_path):
    when = datetime.datetime(2026, 10, 13, 14, 5, 1)
    a = runlog.make_run_id("manual", when, str(tmp_path))
    runlog.RunLog(a, "manual", str(tmp_path), echo=False).close()
    b = runlog.make_run_id("manual", when, str(tmp_path))
    assert b == a + "_2"
    assert runlog.parse_run_id(b) == (when, "manual")


def test_runlog_writes_utf8_lines_flushed_per_line(tmp_path, capsys):
    log = runlog.RunLog("20261013_140501_manual", "manual", str(tmp_path))
    log.warn("下载", "清货1_PW03 没下到\n第二行", store="清货1_PW03")
    # 还没 close 就能读到：每行都 flush 了。
    with io.open(log.path, "rb") as f:
        raw = f.read()
    log.close()
    lines = raw.decode("utf-8").split("\n")
    assert lines[-1] == "" and len(lines) == 3
    assert b"\r" not in raw
    recs = [runlog.parse_line(x) for x in lines[:2]]
    assert [r["message"] for r in recs] == ["清货1_PW03 没下到", "第二行"]
    assert all(r["entrance"] == "手动" and r["store"] == "清货1_PW03"
               and r["level"] == "警告" for r in recs)
    # 同样的内容回显到 stdout
    assert capsys.readouterr().out.splitlines() == lines[:2]


# ------------------------------------------------------------ 逐行分类

BATCH_OUTPUT = """\
============================================================
批次 20261013_020001_m202609 —— 2 家店铺：BOCINA_TA02, 清货1_PW03
路线：all | 白名单：2 家
============================================================

############################################################
# [1/2] BOCINA_TA02
############################################################
输出目录：D:\\x\\downloads\\BOCINA_TA02\\20261013_020001_m202609
  [警告] 没有出现新行，将按 '9 agosto 2026 - 8 octubre 2026' 匹配
    [剩余  230 秒] 列表中尚未出现该行

  [待取] 本次未能及时就绪，后续运行可取回：
      mercadopago/liberaciones (2026-09)
  [无数据] 该区间没有可报的内容，这不是错误：
  [错误] 1 条路线失败：
  本次下载合计 ： 11

############################################################
# [2/2] 清货1_PW03
############################################################
  [失败] TimeoutException: Message: timeout
Traceback (most recent call last):
  File "run_batch.py", line 342, in main
    result = run_downloads.run_store(client, name, args,
selenium.common.exceptions.TimeoutException: Message: timeout
  界面语言：es ✓

============================================================
补收阶段 —— 1 张报表，涉及 1 家店铺
============================================================

--- 补收：BOCINA_TA02（1 张报表）
    已取回 0 张，仍欠 1 张

>>> 清洗：-m mx_sales build
    工作目录：D:\\x\\downloads
      ERROR Ventas_MX_x.xlsx: bad header
    UserWarning: Workbook contains no default style   (x22)
    清洗 部分文件失败（其余已写入），用时 1 分 02 秒

============================================================
批次汇总 —— 用时 31 分 07 秒
============================================================
  店铺             状态          文件数  备注
  BOCINA_TA02      partial          11  1 处错误：x | 1 张仍在生成
  清货1_PW03       failed            0  1 处错误：TimeoutException
  2/2 家店铺需要关注
[BATCH_SUMMARY_JSON] D:\\x\\logs\\batch_20261013_020001_m202609.json
"""


def classify(child, text, stores=None):
    c = runlog.LineClassifier(child, stores)
    out = [c.feed(line) for line in text.splitlines()]
    return c, [o for o in out if o is not None]


def test_classifier_on_real_run_batch_output():
    c, got = classify("download", BATCH_OUTPUT)
    by_msg = {m.strip(): (step, store, level) for step, store, level, m in got}

    # 装饰线、空行、机器可读行都不进日志
    assert not any(set(m.strip()) <= set("=#") for _, _, _, m in got)
    assert not any("BATCH_SUMMARY_JSON" in m for _, _, _, m in got)
    assert c.markers["BATCH_SUMMARY_JSON"] == (
        "D:\\x\\logs\\batch_20261013_020001_m202609.json")
    assert c.progress == (2, 2, "清货1_PW03")

    assert by_msg["路线：all | 白名单：2 家"] == ("下载", "-", INFO)
    assert by_msg["# [1/2] BOCINA_TA02"] == ("下载", "BOCINA_TA02", INFO)
    assert by_msg["[警告] 没有出现新行，将按 '9 agosto 2026 - 8 octubre 2026' 匹配"] \
        == ("下载", "BOCINA_TA02", WARN)
    assert by_msg["[剩余  230 秒] 列表中尚未出现该行"] == ("下载", "BOCINA_TA02", INFO)
    assert by_msg["[待取] 本次未能及时就绪，后续运行可取回："][2] == WARN
    assert by_msg["[无数据] 该区间没有可报的内容，这不是错误："][2] == INFO
    assert by_msg["[错误] 1 条路线失败："][2] == ERROR
    assert by_msg["本次下载合计 ： 11"] == ("下载", "BOCINA_TA02", INFO)

    # 第二家店：失败 + 整段 Traceback 都是错误，Traceback 结束后恢复正常
    assert by_msg["[失败] TimeoutException: Message: timeout"] == (
        "下载", "清货1_PW03", ERROR)
    assert by_msg["Traceback (most recent call last):"][2] == ERROR
    assert by_msg['File "run_batch.py", line 342, in main'][2] == ERROR
    assert by_msg["result = run_downloads.run_store(client, name, args,"][2] == ERROR
    assert by_msg["selenium.common.exceptions.TimeoutException: Message: timeout"][2] \
        == ERROR
    assert by_msg["界面语言：es ✓"] == ("下载", "清货1_PW03", INFO)

    # 步骤跟着横幅走：补收 → 清洗 → 回到下载（汇总表）
    assert by_msg["补收阶段 —— 1 张报表，涉及 1 家店铺"] == ("补收", "-", INFO)
    assert by_msg["--- 补收：BOCINA_TA02（1 张报表）"] == ("补收", "BOCINA_TA02", INFO)
    assert by_msg["已取回 0 张，仍欠 1 张"] == ("补收", "BOCINA_TA02", WARN)
    assert by_msg[">>> 清洗：-m mx_sales build"] == ("清洗", "-", INFO)
    assert by_msg["ERROR Ventas_MX_x.xlsx: bad header"] == ("清洗", "-", ERROR)
    assert by_msg["UserWarning: Workbook contains no default style   (x22)"][2] == WARN
    assert by_msg["清洗 部分文件失败（其余已写入），用时 1 分 02 秒"] == ("清洗", "-", WARN)

    # 汇总表里按行首认店铺，状态列决定级别
    assert by_msg["BOCINA_TA02      partial          11  1 处错误：x | 1 张仍在生成"] \
        == ("下载", "BOCINA_TA02", WARN)
    assert by_msg["清货1_PW03       failed            0  1 处错误：TimeoutException"] \
        == ("下载", "清货1_PW03", ERROR)
    assert by_msg["2/2 家店铺需要关注"] == ("下载", "-", WARN)


def test_classifier_keeps_indentation_in_message():
    _, got = classify("download", "# [1/1] A\n      file.xlsx\n")
    assert got[1] == ("下载", "A", INFO, "      file.xlsx")


REPORT_OUTPUT = """\
==============================================================
月度财务报表 —— 2026-09，2 家店铺
==============================================================
  BOCINA_TA02            20261013_020001_m202609（账单齐全）
  UNIT_TA04              [跳过] 没有下载目录

[提示] 未检测到 LibreOffice，跳过公式重算。

--------------------------------------------------------------
单店报表：BOCINA_TA02
--------------------------------------------------------------
    [校验]
      ✓  BOCINA_TA02  桥A 逐笔勾稽          残差 0.00
      ✗  BOCINA_TA02  桥B 账单应付          差异 12.30
    [诊断] error 级未通过 0 项；warn 级未通过 0 项
    [诊断] error 级未通过 1 项；warn 级未通过 0 项
    [诊断日志] D:\\x\\diagnostics\\a.log
    [错误] 未能写入 BOCINA_TA02_2026-09.xlsx
[REPORT_CHECKS_JSON] []

--------------------------------------------------------------
汇总报表：1 家店铺
--------------------------------------------------------------
    [剔除] UNIT_TA04 —— 没有账单

==============================================================
报表生成结果
==============================================================
  BOCINA_TA02            有校验未过      41s  BOCINA_TA02_2026-09.xlsx
  全店汇总               成功            63s  MX_ML_全店汇总_2026-09.xlsx
    [error][bridge_b] BOCINA_TA02 桥B：差异 12.30
  1/2 份报表需要关注
"""


def test_classifier_on_real_run_reports_output():
    c, got = classify("report", REPORT_OUTPUT, stores=["BOCINA_TA02", "UNIT_TA04"])
    by_msg = {m.strip(): (step, store, level) for step, store, level, m in got}
    assert all(step == "报表" for step, _, _, _ in got)
    assert "REPORT_CHECKS_JSON" in c.markers
    assert not any("REPORT_CHECKS_JSON" in m for _, _, _, m in got)

    assert by_msg["月度财务报表 —— 2026-09，2 家店铺"][1] == "-"
    assert by_msg["BOCINA_TA02            20261013_020001_m202609（账单齐全）"] \
        == ("报表", "BOCINA_TA02", INFO)
    assert by_msg["UNIT_TA04              [跳过] 没有下载目录"] == ("报表", "UNIT_TA04", WARN)
    assert by_msg["[提示] 未检测到 LibreOffice，跳过公式重算。"] == ("报表", "-", INFO)
    assert by_msg["单店报表：BOCINA_TA02"] == ("报表", "BOCINA_TA02", INFO)
    assert by_msg["✓  BOCINA_TA02  桥A 逐笔勾稽          残差 0.00"][2] == INFO
    # 校验未过是"要人看"，不是程序出错 —— 桥B 每家店每个月都不过
    assert by_msg["✗  BOCINA_TA02  桥B 账单应付          差异 12.30"][2] == WARN
    assert by_msg["[诊断] error 级未通过 0 项；warn 级未通过 0 项"][2] == INFO
    assert by_msg["[诊断] error 级未通过 1 项；warn 级未通过 0 项"][2] == WARN
    assert by_msg["[错误] 未能写入 BOCINA_TA02_2026-09.xlsx"] == (
        "报表", "BOCINA_TA02", ERROR)
    assert by_msg["汇总报表：1 家店铺"][1] == "全店汇总"
    assert by_msg["[剔除] UNIT_TA04 —— 没有账单"] == ("报表", "全店汇总", WARN)
    assert by_msg["BOCINA_TA02            有校验未过      41s  BOCINA_TA02_2026-09.xlsx"] \
        == ("报表", "BOCINA_TA02", WARN)
    assert by_msg["全店汇总               成功            63s  MX_ML_全店汇总_2026-09.xlsx"] \
        == ("报表", "全店汇总", INFO)
    assert by_msg["[error][bridge_b] BOCINA_TA02 桥B：差异 12.30"][2] == ERROR
    assert by_msg["1/2 份报表需要关注"] == ("报表", "-", WARN)


def test_classifier_collect_child_and_hard_errors():
    _, got = classify("collect", "仅补收模式：1 家店铺有待取回的报表：A\n"
                                 "--- 补收：A（2 张报表）\n"
                                 "    已取回 2 张，仍欠 0 张\n"
                                 "\n紫鸟客户端错误：client did not start\n")
    assert got[0][:3] == ("补收", "-", INFO)
    assert got[2] == ("补收", "A", INFO, "    已取回 2 张，仍欠 0 张")
    assert got[3][2] == ERROR


def test_indented_traceback_ends_at_the_exception_line():
    # financial_report 把 report_build 的输出整体缩进 4 格再打出来
    text = ("    Traceback (most recent call last):\n"
            "      File \"report_build.py\", line 1, in <module>\n"
            "    KeyError: 'x'\n"
            "    正常的一行\n")
    _, got = classify("report", text)
    assert [g[2] for g in got] == [ERROR, ERROR, ERROR, INFO]


# ------------------------------------------------------------ 跟读

def test_tailer_handles_partial_lines_and_split_multibyte(tmp_path):
    p = str(tmp_path / "a.log")
    t = runlog.Tailer(p)
    assert t.poll() == ([], False)                 # 文件还不存在
    data = "第一行\n第二行还没写完".encode("utf-8")
    with io.open(p, "ab") as f:
        f.write(data[:-4])                         # 切在一个汉字中间
    assert t.poll() == (["第一行"], False)
    with io.open(p, "ab") as f:
        f.write(data[-4:] + "了\r\n".encode("utf-8"))
    assert t.poll() == (["第二行还没写完了"], False)
    assert t.poll() == ([], False)
    # 文件被换成更短的一份：从头读
    with io.open(p, "wb") as f:
        f.write(b"new\n")
    assert t.poll() == (["new"], False)


def test_tailer_reads_big_files_in_chunks(tmp_path):
    p = str(tmp_path / "a.log")
    with io.open(p, "wb") as f:
        f.write(("行\n" * 50).encode("utf-8"))
    t = runlog.Tailer(p)
    t.CHUNK = 64
    n, more, rounds = 0, True, 0
    while more:
        lines, more = t.poll()
        assert all(x == "行" for x in lines)
        n += len(lines)
        rounds += 1
    assert n == 50 and rounds > 1


# ------------------------------------------------------------ 正在跑的那一轮

def test_active_run_needs_live_pid_and_fresh_heartbeat(tmp_path):
    d = str(tmp_path)
    assert runlog.active_run(d) is None
    runlog.write_current({"run_id": "20261013_140501_manual",
                          "pid": os.getpid(), "entrance": "manual"}, d)
    assert runlog.active_run(d)["run_id"] == "20261013_140501_manual"
    # 心跳断了但进程还在：是"卡住了"，不是"没在跑"（见 test_runlog_stall.py）
    assert runlog.active_run(
        d, now=time.time() + runlog.HEARTBEAT_STALE + 5)["stalled"]
    # 断了很久（断电重启后 PID 被别的程序占用的情形）：不再信那个 PID
    assert runlog.active_run(d, now=time.time() + runlog.STALL_GIVEUP + 5) is None
    runlog.touch_current(d)
    assert runlog.active_run(d) is not None
    # 别人的 run_id 不会被误删
    runlog.clear_current("other", d)
    assert runlog.active_run(d) is not None
    runlog.clear_current("20261013_140501_manual", d)
    assert runlog.active_run(d) is None


def test_active_run_ignores_dead_pid(tmp_path):
    import subprocess
    import sys
    p = subprocess.Popen([sys.executable, "-c", "pass"])
    p.wait()
    runlog.write_current({"run_id": "20261013_140501_cli", "pid": p.pid}, str(tmp_path))
    assert runlog.pid_alive(p.pid) is False
    assert runlog.active_run(str(tmp_path)) is None
    assert runlog.pid_alive(os.getpid()) is True
    assert runlog.pid_alive(0) is True            # 判断不了：宁可当它在


def test_stop_request_round_trip(tmp_path):
    d, rid = str(tmp_path), "20261013_140501_scheduled"
    assert not runlog.stop_requested(rid, d)
    p = runlog.request_stop(rid, "启动窗口", d)
    assert p == runlog.run_paths(rid, d)["stop"] and runlog.stop_requested(rid, d)
    runlog.clear_stop(rid, d)
    assert not runlog.stop_requested(rid, d)


# ------------------------------------------------------------ 状态 / 列表

def _fake_run(d, rid, **summary):
    log = runlog.RunLog(rid, runlog.parse_run_id(rid)[1], d, echo=False)
    log.info("启动", "运行编号：%s" % rid)
    log.close()
    if summary:
        runlog.write_json(runlog.run_paths(rid, d)["json"],
                          dict({"run_id": rid, "month": "2026-09"}, **summary))


def test_exit_meaning_covers_the_contract():
    kinds = {c: runlog.exit_meaning(c)[1] for c in range(8)}
    assert kinds == {0: "good", 1: "warn", 2: "bad", 3: "bad", 4: "bad",
                     5: "warn", 6: "warn", 7: "bad"}
    assert runlog.exit_meaning(99)[1] == "bad"
    assert runlog.exit_meaning(None)[1] == "bad"
    assert "中止" in runlog.exit_meaning("6")[0]


def test_run_state_and_listing(tmp_path):
    d = str(tmp_path)
    _fake_run(d, "20261011_020000_scheduled", status="done", exit=1, pid=1)
    _fake_run(d, "20261012_090000_manual", status="aborted", exit=6, pid=1)
    _fake_run(d, "20261013_020000_scheduled", status="running", exit=None,
              pid=os.getpid())
    _fake_run(d, "20261013_030000_cli")                       # 没有摘要
    io.open(os.path.join(d, "runs", "notes.log"), "w").close()  # 不是运行记录

    ids = [r["run_id"] for r in runlog.list_runs(d)]
    assert ids == ["20261013_030000_cli", "20261013_020000_scheduled",
                   "20261012_090000_manual", "20261011_020000_scheduled"]

    st = runlog.run_state("20261011_020000_scheduled", d)
    assert (st["state"], st["kind"], st["entrance"], st["month"]) == (
        "finished", "warn", "scheduled", "2026-09")
    assert runlog.run_state("20261012_090000_manual", d)["state"] == "aborted"
    assert runlog.run_state("20261013_020000_scheduled", d)["state"] == "running"
    assert runlog.run_state("20261013_030000_cli", d)["state"] == "unknown"
    label = runlog.run_label(runlog.list_runs(d)[3], d)
    assert label.startswith("10-11 02:00  定时  2026-09  ")


def test_run_state_running_with_dead_pid_is_interrupted(tmp_path, monkeypatch):
    d = str(tmp_path)
    _fake_run(d, "20261013_020000_scheduled", status="running", pid=424242)
    monkeypatch.setattr(runlog, "pid_alive", lambda pid: False)
    st = runlog.run_state("20261013_020000_scheduled", d)
    assert st["state"] == "interrupted" and st["kind"] == "bad"


# ------------------------------------------------------------ Follower

def test_follower_shows_latest_then_switches_to_a_new_live_run(tmp_path):
    d = str(tmp_path)
    f = runlog.Follower(d)
    r = f.tick()
    assert r["state"] is None and r["new"] == []

    _fake_run(d, "20261011_020000_scheduled", status="done", exit=0, pid=1)
    r = f.tick()
    assert r["reset"] and f.run_id == "20261011_020000_scheduled"
    assert len(r["new"]) == 1 and r["state"]["state"] == "finished"
    assert f.tick()["reset"] is False

    # 窗口开着的时候，另一个入口起了一轮：自动切过去
    rid = "20261013_020000_scheduled"
    log = runlog.RunLog(rid, "scheduled", d, echo=False)
    log.info("下载", "# [1/2] A", store="A")
    runlog.write_json(runlog.run_paths(rid, d)["json"],
                      {"status": "running", "pid": os.getpid(), "month": "2026-09"})
    runlog.write_current({"run_id": rid, "pid": os.getpid()}, d)
    r = f.tick()
    assert r["reset"] and f.run_id == rid and r["state"]["state"] == "running"
    log.warn("下载", "[警告] x", store="B")
    r = f.tick(now=time.time() + 5)
    assert [x["store"] for x in r["new"]] == ["B"] and not r["reset"]
    assert f.stores == ["A", "B"]
    assert [x["level"] for x in f.filtered(min_level=WARN)] == [WARN]
    assert [x["store"] for x in f.filtered(store="A")] == ["A"]

    # 跑完：current.json 删掉、摘要落定 —— 仍然看这一轮，状态变成结束
    runlog.clear_current(rid, d)
    runlog.write_json(runlog.run_paths(rid, d)["json"],
                      {"status": "done", "exit": 7, "pid": os.getpid()})
    log.close()
    r = f.tick(now=time.time() + 10)
    assert f.run_id == rid and r["state_changed"]
    assert (r["state"]["state"], r["state"]["exit"]) == ("finished", 7)


def test_follower_pinned_run_is_not_stolen_by_a_live_one(tmp_path):
    d = str(tmp_path)
    _fake_run(d, "20261011_020000_scheduled", status="done", exit=0, pid=1)
    _fake_run(d, "20261013_020000_manual", status="running", pid=os.getpid())
    runlog.write_current({"run_id": "20261013_020000_manual",
                          "pid": os.getpid()}, d)
    f = runlog.Follower(d)
    f.pin("20261011_020000_scheduled")
    f.tick()
    assert f.run_id == "20261011_020000_scheduled"
    f.follow_current()
    assert f.tick()["reset"] and f.run_id == "20261013_020000_manual"


# ------------------------------------------------------------ 起一轮的参数

def test_runner_argv_differs_only_in_entrance():
    a = runlog.runner_argv("manual", "both", "2026-09", ["A", "B"], "D:\\out")
    b = runlog.runner_argv("scheduled", "both", "2026-09", ["A", "B"], "D:\\out")
    assert a == ["--entrance", "manual", "--month", "2026-09",
                 "--stores", "A", "B", "--out-dir", "D:\\out"]
    assert [x for x in a if x != "manual"] == [x for x in b if x != "scheduled"]
    # 计划任务登记的就是这一句
    assert runlog.runner_argv("scheduled") == ["--entrance", "scheduled"]
    assert runlog.runner_argv("scheduled", step="report") == [
        "--entrance", "scheduled", "--step", "report"]


def test_git_commit_reads_this_checkout():
    c = runlog.git_commit()
    assert c is None or (len(c) == 7 and int(c, 16) >= 0)


def test_spawn_detached_survives_and_captures_stderr(tmp_path):
    import sys
    err = str(tmp_path / "runs" / "_launch.err")
    p = runlog.spawn_detached(
        [sys.executable, "-c", "import sys; print('o'); sys.stderr.write('boom'); "
                               "sys.exit(3)"], cwd=str(tmp_path), stderr_path=err)
    assert p.wait(timeout=30) == 3
    with io.open(err, "rb") as f:
        assert f.read() == b"boom"
