# -*- coding: utf-8 -*-
"""run_monthly：唯一的运行器。子进程换成几行 Python，流程本身照真的走。"""
import io
import json
import os
import sys
import threading
import time

import pytest

from core import run_monthly
from core import runlog

PY = sys.executable
MONTH = "2026-09"


@pytest.fixture
def env(tmp_path, monkeypatch):
    """把日志、锁、输出目录都指到临时目录，并关掉 webhook 和白名单读取。"""
    logs = tmp_path / "logs"
    out = tmp_path / "out"
    monkeypatch.setattr(run_monthly, "LOG_DIR", str(logs))
    monkeypatch.setattr(run_monthly, "LOCK", str(logs / "monthly.lock"))
    monkeypatch.setattr(run_monthly, "config", lambda: {})
    monkeypatch.setattr(run_monthly, "known_stores", lambda: [])
    monkeypatch.setattr(run_monthly, "POLL", 0.1)

    class Env(object):
        log_dir = str(logs)
        out_dir = str(out)
        # 整个月只有一份「！需要处理_<年月>.txt」（core/month_overview），这一轮 urgent 的
        # 每一句都原样写在里面。
        alert = str(out / ("！需要处理_%s.txt" % MONTH.replace("-", "")))

        def run(self, argv, plan):
            args = run_monthly.parse_args(
                ["--month", MONTH, "--out-dir", self.out_dir] + argv)
            return run_monthly.execute(args, plan=plan)

        def last(self):
            rid = runlog.list_runs(self.log_dir)[0]["run_id"]
            p = runlog.run_paths(rid, self.log_dir)
            with io.open(p["log"], encoding="utf-8") as f:
                recs = [runlog.parse_line(x) for x in f.read().splitlines()]
            assert all(recs), "每一行都必须是标准格式"
            return rid, runlog.read_json(p["json"]), recs

        def alert_text(self):
            with io.open(self.alert, encoding="utf-8") as f:
                return f.read()

    return Env()


def child(code, text="", sleep=0):
    """一个假子进程：打几行、睡一会儿、按给定退出码结束。"""
    src = ("import sys, time\n"
           "sys.stdout.reconfigure(encoding='utf-8')\n"
           "print(%r, flush=True)\n"
           "time.sleep(%r)\n"
           "sys.exit(%d)\n" % (text, sleep, code))
    return [PY, "-c", src]


def plan(*steps):
    return [{"key": k, "label": k, "cmd": c} for k, c in steps]


# ------------------------------------------------------------ 命令行

def strip_py(p):
    return [(s["key"], [os.path.basename(c) if c.endswith(".py") else c
                        for c in s["cmd"][1:]]) for s in p]


def test_plan_for_each_step_always_passes_month():
    full = strip_py(run_monthly.build_plan("both", MONTH, None, "OUT", py="py"))
    assert full == [
        ("download", ["-u", "-m", "platforms.mercadolibre.run_batch", "--month", MONTH]),
        ("report", ["-u", "-m", "platforms.mercadolibre.run_reports", "--month", MONTH, "--out-dir", "OUT"]),
    ]
    assert strip_py(run_monthly.build_plan("download", MONTH, None, "OUT")) == full[:1]
    assert strip_py(run_monthly.build_plan("report", MONTH, None, "OUT")) == full[1:]
    assert strip_py(run_monthly.build_plan("collect", MONTH, None, "OUT")) == [
        ("collect", ["-u", "-m", "platforms.mercadolibre.run_batch", "--collect-only", "--month", MONTH])]


def test_plan_with_store_subset_never_writes_the_combined_workbook():
    p = strip_py(run_monthly.build_plan("both", MONTH, ["A", "清货1_PW03"], "OUT"))
    assert p == [
        ("download", ["-u", "-m", "platforms.mercadolibre.run_batch", "--month", MONTH,
                      "--stores", "A", "清货1_PW03"]),
        ("report", ["-u", "-m", "platforms.mercadolibre.run_reports", "--month", MONTH,
                    "--stores", "A", "清货1_PW03", "--no-combined",
                    "--out-dir", "OUT"]),
    ]
    c = strip_py(run_monthly.build_plan("collect", MONTH, ["A"], "OUT"))
    assert c[0][1][-2:] == ["--stores", "A"]
    # 全部店铺时不能带 --no-combined
    assert "--no-combined" not in run_monthly.build_plan(
        "report", MONTH, None, "OUT")[0]["cmd"]


@pytest.mark.parametrize("extra", [[], ["--stores", "A", "B"],
                                   ["--step", "report"], ["--step", "collect"]])
def test_manual_and_scheduled_build_identical_child_commands(extra, env, capsys):
    """两个入口跑的是同一个过程：同样的参数，子进程命令一字不差。"""
    seen = {}
    for entrance in ("manual", "scheduled", "cli"):
        args = run_monthly.parse_args(
            ["--entrance", entrance, "--month", MONTH, "--out-dir", env.out_dir]
            + extra)
        out_dir, _ = run_monthly.resolve_out_dir(args.out_dir)
        seen[entrance] = [s["cmd"] for s in run_monthly.build_plan(
            args.step, args.month, args.stores, out_dir)]
        # --dry-run 打出来的也是这些命令
        args.dry_run = True
        assert run_monthly.execute(args) == 0
        printed = capsys.readouterr().out
        for cmd in seen[entrance]:
            assert run_monthly.show(cmd) in printed
    assert seen["manual"] == seen["scheduled"] == seen["cli"]
    assert all("--month" in cmd for cmd in seen["manual"])
    assert not os.path.exists(env.log_dir)            # dry-run 不留任何文件


def test_scheduled_task_arguments_are_just_the_entrance():
    """计划任务只带 --entrance scheduled：月份默认上个月，和手动默认选的一样。"""
    a = run_monthly.parse_args(runlog.runner_argv("scheduled"))
    assert (a.entrance, a.step, a.month, a.stores) == ("scheduled", "both", None, None)
    assert run_monthly.parse_args([]).entrance == "cli"
    assert run_monthly.parse_args([]).max_hours == 5


# ------------------------------------------------------------ 正常一轮

def test_full_run_writes_standard_log_summary_and_cleans_up(env):
    code = env.run(["--entrance", "scheduled"], plan(
        ("download", child(0, "# [1/1] A\n  [警告] x\n=====\n\n"
                              "[BATCH_SUMMARY_JSON] /nope.json")),
        ("report", child(0, "单店报表：A"))))
    assert code == 0
    rid, doc, recs = env.last()
    assert rid.endswith("_scheduled")
    assert {r["entrance"] for r in recs} == {"定时"}

    head = "\n".join(r["message"] for r in recs if r["step"] == "启动")
    for needle in ("入口　　：定时", "运行编号：" + rid, "目标月份：" + MONTH,
                   "店铺　　：全部", "命令 1　：", "命令 2　：", "输出目录：",
                   "机器　　：", "用户：", "Python　：", "代码版本："):
        assert needle in head, needle

    msgs = [(r["step"], r["store"], r["level"], r["message"].strip()) for r in recs]
    assert ("下载", "A", "警告", "[警告] x") in msgs
    assert ("报表", "A", "信息", "单店报表：A") in msgs
    # 机器可读行和装饰线不进日志（文件头里的"命令"那几行是测试命令本身，不算）
    assert not any("BATCH_SUMMARY_JSON" in m[3] or m[3] == "====="
                   for m in msgs if m[0] != "启动")
    assert recs[-1]["step"] == "收尾" and "退出码 0" in recs[-1]["message"]

    assert doc["status"] == "done" and doc["exit"] == 0
    assert doc["entrance"] == "scheduled" and doc["run_id"] == rid
    assert doc["pid"] == os.getpid() and doc["stores"] == "all"
    assert [s["step"] for s in doc["steps"]] == ["download", "report"]
    assert [c["step"] for c in doc["commands"]] == ["download", "report"]
    assert doc["batch_summary"] == "/nope.json"

    # 跑完：锁放了、current.json 没了、没有告警、没有旧的 monthly_*.log
    assert not os.path.exists(run_monthly.LOCK)
    assert runlog.active_run(env.log_dir) is None
    assert not os.path.exists(runlog.current_path(env.log_dir))
    assert not os.path.exists(env.alert)
    assert not [n for n in os.listdir(env.log_dir) if n.startswith("monthly_")]


def test_current_json_points_at_the_live_run(env):
    seen = {}

    def peek():
        for _ in range(100):
            cur = runlog.active_run(env.log_dir)
            if cur:
                seen.update(cur)
                seen["state"] = runlog.run_state(cur["run_id"], env.log_dir)["state"]
                return
            time.sleep(0.05)

    t = threading.Thread(target=peek)
    t.start()
    env.run(["--entrance", "manual", "--step", "download"],
            plan(("download", child(0, "x", sleep=1.5))))
    t.join()
    rid, _, _ = env.last()
    assert seen["run_id"] == rid and seen["entrance"] == "manual"
    assert seen["pid"] == os.getpid() and seen["month"] == MONTH
    assert seen["log"] == runlog.run_paths(rid, env.log_dir)["log"]
    assert seen["state"] == "running"


# ------------------------------------------------------------ 闸门 / 告警

def test_download_exit_2_skips_reports_and_alerts(env, tmp_path):
    marker = tmp_path / "report_ran"
    code = env.run(["--entrance", "manual"], plan(
        ("download", child(2, "紫鸟客户端错误：x")),
        ("report", [PY, "-c", "open(%r, 'w').close()" % str(marker)])))
    assert code == 2
    assert not marker.exists(), "下载硬失败后不能出报表"
    _, doc, recs = env.last()
    assert doc["status"] == "failed" and [s["step"] for s in doc["steps"]] == ["download"]
    text = env.alert_text()
    assert "下载失败（退出码 2" in text and "报表未生成" in text
    assert "这一轮是：手动" in text
    assert runlog.run_paths(doc["run_id"], env.log_dir)["log"] in text
    assert "monthly_" not in text
    assert any("已写出待处理清单" in r["message"] for r in recs)
    assert not os.path.exists(run_monthly.LOCK)


@pytest.mark.parametrize("code", [3, 4])
def test_download_exit_3_and_4_also_stop(env, code):
    assert env.run([], plan(("download", child(code)),
                            ("report", child(0)))) == code
    assert len(env.last()[1]["steps"]) == 1 and os.path.exists(env.alert)


def test_download_exit_1_continues_to_reports(env, tmp_path):
    marker = tmp_path / "report_ran"
    code = env.run([], plan(
        ("download", child(1, "  1/12 家店铺需要关注")),
        ("report", [PY, "-c", "open(%r, 'w').close()" % str(marker)])))
    assert code == 1 and marker.exists()
    assert env.last()[1]["status"] == "done"
    # 没有批次汇总可读时退回一句笼统的话，照旧告警
    assert "没读到批次汇总" in env.alert_text()


def _summary(tmp_path, **doc):
    p = tmp_path / "batch.json"
    p.write_text(json.dumps(doc, ensure_ascii=False), encoding="utf-8")
    return str(p)


def test_alert_says_what_actually_happened(env, tmp_path):
    s = _summary(tmp_path, stores=[
        {"store": "A", "status": "ok", "pending": [], "errors": []},
        {"store": "B", "status": "failed", "pending": [],
         "errors": ["TimeoutException: x"]},
        {"store": "C", "status": "partial", "pending": [],
         "errors": ["stock: x"],
         "missing": [{"form": "stock/general", "cause": "UI_NOT_FOUND"},
                     {"form": "mercadopago/release", "cause": "STILL_GENERATING"}]},
        {"store": "WAITSHOP", "status": "partial", "pending": ["mp/x", "mp/y"],
         "errors": [],
         "missing": [{"form": "mercadopago/settlement", "cause": "PLATFORM_DELAYED"},
                     {"form": "mercadopago/release", "cause": "STILL_GENERATING"}]}],
        transform={"ok": False, "partial": True, "skipped": None})
    code = env.run([], plan(
        ("download", child(1, "[BATCH_SUMMARY_JSON] " + s.replace("\\", "\\\\"))),
        ("report", child(0))))
    assert code == 1
    text = env.alert_text()
    assert "1/4 家店铺没下到数据：B" in text and "TimeoutException" in text
    assert "1 家店铺有部分报表没下成：C" in text
    assert "缺 1 张：页面是对的，但要点的控件找不到或点不了（UI_NOT_FOUND）" in text
    assert "部分店铺下载失败" not in text
    # 只是平台还没生成完的那家店（WAITSHOP）不进告警，清洗的事也不进；都在日志和摘要里
    assert "WAITSHOP" not in text and "还欠" not in text
    assert "清洗" not in text
    att = env.last()[1]["attention"]
    assert any("清洗" in a for a in att)
    assert any("2 张不影响报表的文件没下到" in a and "WAITSHOP" in a for a in att)


def test_reports_the_platform_is_still_building_do_not_alert(env, tmp_path):
    """2026-10-09 那轮 12 家店里 7 家有 MercadoPago 报表当晚没生成完 —— 常态。
    为它每个月写一个"报表异常"，那个文件很快就没人看了。"""
    s = _summary(tmp_path, stores=[
        {"store": "A", "status": "partial", "pending": ["mp/x"], "errors": [],
         "missing": [{"form": "mercadopago/settlement",
                      "cause": "PLATFORM_DELAYED"}]}])
    code = env.run([], plan(
        ("download", child(1, "[BATCH_SUMMARY_JSON] " + s.replace("\\", "\\\\"))),
        ("report", child(0))))
    assert code == 1
    assert not os.path.exists(env.alert)
    _, doc, recs = env.last()
    assert doc["urgent"] == []
    assert "1 张不影响报表的文件没下到" in doc["attention"][0] and "：A。" in doc["attention"][0]
    assert any("补取欠的报表" in r["message"] for r in recs if r["step"] == "收尾")


def test_cleaning_only_failure_does_not_alert(env, tmp_path):
    s = _summary(tmp_path, stores=[{"store": "A", "status": "ok",
                                    "pending": [], "errors": []}],
                 transform={"ok": False, "partial": True})
    code = env.run([], plan(
        ("download", child(1, "[BATCH_SUMMARY_JSON] " + s.replace("\\", "\\\\"))),
        ("report", child(0))))
    assert code == 1
    assert not os.path.exists(env.alert)
    _, doc, recs = env.last()
    assert doc["urgent"] == [] and "清洗" in doc["attention"][0]
    assert any("清洗" in r["message"] for r in recs if r["step"] == "收尾")


def test_describe_batch_tolerates_missing_or_broken_summary(tmp_path):
    assert run_monthly.describe_batch(None) == ([], [])
    assert run_monthly.describe_batch(str(tmp_path / "nope.json")) == ([], [])
    bad = tmp_path / "bad.json"
    bad.write_text("{", encoding="utf-8")
    assert run_monthly.describe_batch(str(bad)) == ([], [])
    # 旧版汇总：没有逐张的 missing 记录，partial 就按"有报表没下成"报
    old = _summary(tmp_path, stores=[{"store": "A", "status": "partial",
                                      "pending": ["mp/x"]}],
                   missing_by_cause={"UI_NOT_FOUND": 1})
    u, a = run_monthly.describe_batch(old)
    assert "1 家店铺有部分报表没下成：A" in u[0] and "UI_NOT_FOUND" in u[1]
    assert a == []


def test_billing_summary_error_about_bill_mp_v3_is_harmless(tmp_path):
    # 2026-10-15 智利 DAB_CL：账单路线报错是一行汇总，报表名在中间；缺的全是
    # 报表不读的（BILL_MP_V3、account statement），不该算"部分报表没下成"
    s = _summary(tmp_path, stores=[{
        "store": "D", "status": "partial", "pending": [],
        "errors": ["mercadopago/account_statement/request: requested nothing",
                   "billing: 2 张未下载 —— billing/2026-09/BILL_MP_V3 "
                   "[UI_NOT_FOUND]；billing/2026-10/BILL_MP_V3 [UI_NOT_FOUND]"],
        "missing": [{"form": "mercadopago/account_statement", "cause": "UI_NOT_FOUND"},
                    {"form": "billing/2026-09/BILL_MP_V3", "cause": "UI_NOT_FOUND"},
                    {"form": "billing/2026-10/BILL_MP_V3", "cause": "UI_NOT_FOUND"}]}])
    u, a = run_monthly.describe_batch(s)
    assert u == [] and "3 张不影响报表的文件" in a[0]


def test_report_exit_1_does_not_alert_but_exit_2_does(env):
    assert env.run(["--step", "report"], plan(("report", child(1)))) == 1
    assert not os.path.exists(env.alert)
    assert env.last()[1]["attention"]

    assert env.run(["--step", "report"], plan(("report", child(2)))) == 2
    text = env.alert_text()
    assert "报表步骤失败（退出码 2）" in text and "没有生成" in text
    assert env.last()[1]["status"] == "failed"


def test_crashed_child_exit_code_is_normalised(env):
    # 被信号杀掉（负数）或 Windows 上的崩溃码：都算硬失败 2，摘要里留原值
    assert env.run(["--step", "report"], plan(("report", child(77)))) == 2
    assert env.last()[1]["steps"][0]["code"] == 77
    assert [run_monthly.norm(c) for c in (0, 1, 4, 5, -9, 3221225477)] == [
        0, 1, 4, 2, 2, 2]


# ------------------------------------------------------------ 锁

def test_lock_acquire_release_and_refusal(env):
    assert run_monthly.acquire() is True
    with io.open(run_monthly.LOCK, encoding="utf-8") as f:
        assert ("pid=%d " % os.getpid()) in f.read()
    said = []
    assert run_monthly.acquire(lambda lv, t: said.append((lv, t))) is False
    assert said[0][0] == runlog.WARN and "还在跑" in said[0][1]
    run_monthly.release()
    assert not os.path.exists(run_monthly.LOCK)
    run_monthly.release()                              # 再放一次不报错
    assert run_monthly.acquire() is True
    run_monthly.release()


def test_stale_lock_is_taken_over(env, monkeypatch):
    os.makedirs(env.log_dir)
    with io.open(run_monthly.LOCK, "w", encoding="utf-8") as f:
        f.write("pid=424242 起于 2026-10-13 10:01:00")
    # 进程已经不在：立刻接管
    monkeypatch.setattr(run_monthly, "alive", lambda pid: pid != 424242)
    said = []
    assert run_monthly.acquire(lambda lv, t: said.append(t)) is True
    assert "已经不在了" in said[0]
    with io.open(run_monthly.LOCK, encoding="utf-8") as f:
        assert ("pid=%d " % os.getpid()) in f.read()

    # 进程号还在，但锁的心跳停了太久（PID 被复用 / 那一轮卡死）：也接管
    monkeypatch.setattr(run_monthly, "alive", lambda pid: True)
    old = time.time() - run_monthly.LOCK_QUIET - 60
    os.utime(run_monthly.LOCK, (old, old))
    assert run_monthly.acquire(lambda lv, t: said.append(t)) is True
    assert "判为异常" in said[-1]
    run_monthly.release()


def test_second_run_exits_5_without_touching_the_live_one(env, tmp_path):
    marker = tmp_path / "ran"
    assert run_monthly.acquire() is True
    runlog.write_current({"run_id": "20260101_000000_scheduled",
                          "pid": os.getpid()}, env.log_dir)
    try:
        code = env.run(["--entrance", "manual"], plan(
            ("download", [PY, "-c", "open(%r, 'w').close()" % str(marker)])))
    finally:
        pass
    assert code == 5 and not marker.exists()
    _, doc, recs = env.last()
    assert doc["status"] == "skipped" and doc["exit"] == 5
    assert runlog.run_state(doc["run_id"], env.log_dir)["state"] == "skipped"
    assert any("另一轮还在跑" in r["message"] for r in recs)
    # 正在跑的那一轮的锁和 current.json 原封不动
    assert os.path.exists(run_monthly.LOCK)
    assert runlog.active_run(env.log_dir)["run_id"] == "20260101_000000_scheduled"
    assert not os.path.exists(env.alert)
    run_monthly.release()


# ------------------------------------------------------------ 中止 / 时限

# 子进程再起一个孙进程，两个都长睡：用来确认杀的是整棵树。
TREE = ("import subprocess, sys, time\n"
        "g = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(120)'])\n"
        "print('pids', __import__('os').getpid(), g.pid, flush=True)\n"
        "time.sleep(120)\n")


def _pids(recs):
    line = [r["message"] for r in recs if r["message"].startswith("pids ")][0]
    return [int(x) for x in line.split()[1:]]


def _gone(pid, within=10):
    end = time.time() + within
    while time.time() < end:
        if not runlog.pid_alive(pid):
            return True
        try:                       # 孙进程成了僵尸也算结束
            with open("/proc/%d/stat" % pid) as f:
                if f.read().split(")")[-1].split()[0] == "Z":
                    return True
        except (IOError, OSError):
            return True
        time.sleep(0.1)
    return False


def test_stop_request_kills_the_child_tree_and_exits_6(env, tmp_path):
    marker = tmp_path / "report_ran"

    def stopper():
        for _ in range(200):
            cur = runlog.active_run(env.log_dir)
            if cur:
                time.sleep(1.0)                    # 等子进程把 pid 打出来
                runlog.request_stop(cur["run_id"], "测试", env.log_dir)
                return
            time.sleep(0.05)

    t = threading.Thread(target=stopper)
    t.start()
    began = time.time()
    code = env.run(["--entrance", "scheduled"], plan(
        ("download", [PY, "-c", TREE]),
        ("report", [PY, "-c", "open(%r, 'w').close()" % str(marker)])))
    t.join()
    assert code == 6
    assert time.time() - began < 30
    assert not marker.exists(), "中止后不能接着出报表"
    rid, doc, recs = env.last()
    assert doc["status"] == "aborted" and doc["exit"] == 6
    assert doc["steps"][0]["ended"] == "aborted"
    assert any("收到中止请求" in r["message"] and r["level"] == "警告" for r in recs)
    child_pid, grandchild_pid = _pids(recs)
    assert _gone(child_pid) and _gone(grandchild_pid), "子进程树必须全部结束"
    # 收尾干净：锁、current.json、中止请求文件都没了；人为中止不告警
    assert not os.path.exists(run_monthly.LOCK)
    assert runlog.active_run(env.log_dir) is None
    assert not runlog.stop_requested(rid, env.log_dir)
    assert not os.path.exists(env.alert)
    assert runlog.run_state(rid, env.log_dir)["state"] == "aborted"


def test_deadline_cuts_download_then_still_reports_and_alerts(env, tmp_path):
    marker = tmp_path / "report_ran"
    began = time.time()
    code = env.run(["--max-hours", "0.0003"], plan(     # 约 1 秒
        ("download", [PY, "-c", "print('# [3/12] STORE_C', flush=True)\n" + TREE]),
        ("report", [PY, "-c", "open(%r, 'w').close()" % str(marker)])))
    assert code == 7
    assert time.time() - began < 30
    assert marker.exists(), "超时后仍然要出报表"
    _, doc, recs = env.last()
    assert doc["status"] == "failed" and doc["exit"] == 7
    assert [(s["step"], s["ended"]) for s in doc["steps"]] == [
        ("download", "deadline"), ("report", "exit")]
    child_pid, grandchild_pid = _pids(recs)
    assert _gone(child_pid) and _gone(grandchild_pid)
    text = env.alert_text()
    assert "被掐断" in text and "第 3/12 家（STORE_C）" in text
    assert any(r["level"] == "错误" and "被掐断" in r["message"] for r in recs)
    assert "退出码 7" in recs[-1]["message"]
    assert not os.path.exists(run_monthly.LOCK)


def test_deadline_gives_way_to_a_worse_report_failure(env):
    code = env.run(["--max-hours", "0.0003"], plan(
        ("download", child(0, "x", sleep=60)), ("report", child(2))))
    assert code == 2
    text = env.alert_text()
    assert "被掐断" in text and "报表步骤失败" in text


def test_report_step_has_no_deadline(env):
    code = env.run(["--step", "report", "--max-hours", "0.0001"],
                   plan(("report", child(0, "x", sleep=1.5))))
    assert code == 0 and env.last()[1]["steps"][0]["ended"] == "exit"


def test_runner_bug_is_logged_and_lock_released(env, monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("内部错误")
    monkeypatch.setattr(run_monthly, "run_child", boom)
    assert env.run([], plan(("download", child(0)))) == 2
    _, doc, recs = env.last()
    assert doc["status"] == "failed"
    assert any("RuntimeError: 内部错误" in r["message"] and r["level"] == "错误"
               for r in recs)
    assert not os.path.exists(run_monthly.LOCK)
    assert runlog.active_run(env.log_dir) is None
    assert os.path.exists(env.alert)


# ------------------------------------------------- 孙进程攥着管道不放

# 子进程正常结束，但它起的孙进程继承了 stdout 并且还活着 —— chromedriver 在
# Windows 上就是这样（selenium 起它时 close_fds=False）。
LINGER = ("import subprocess, sys\n"
          "g = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(40)'])\n"
          "print('pids', __import__('os').getpid(), g.pid, flush=True)\n"
          "sys.exit(1)\n")


def test_lingering_grandchild_does_not_hang_the_runner(env, tmp_path):
    """2026-10-13 实测：下载结束后 chromedriver 没退，管道没关，本脚本卡在
    关管道上二十多分钟，报表那一步始终没开始。"""
    marker = tmp_path / "report_ran"
    began = time.time()
    code = env.run(["--entrance", "manual"], plan(
        ("download", [PY, "-c", LINGER]),
        ("report", [PY, "-c", "open(%r, 'w').close()" % str(marker)])))
    assert time.time() - began < 25, "不能等孙进程自己退出"
    assert code == 1
    assert marker.exists(), "下载那步退出码 1，报表必须接着跑"
    rid, doc, recs = env.last()
    assert doc["status"] == "done"
    assert [s["step"] for s in doc["steps"]] == ["download", "report"]
    assert not os.path.exists(run_monthly.LOCK)
    _, grandchild_pid = _pids(recs)
    try:                                   # 测试自己起的，自己收拾
        os.kill(grandchild_pid, 9)
    except OSError:
        pass


def test_files_the_report_never_reads_do_not_alert(env, tmp_path):
    """2026-10-14：Retiros 的 Generar 点不了、Poscobro 碰上 503 —— 报表一个数都不受影响。"""
    s = _summary(tmp_path, stores=[
        {"store": "TANKE_EE", "status": "partial", "pending": [],
         "errors": ["mercadopago/withdraw/request: requested nothing"],
         "missing": [{"form": "mercadopago/withdraw", "cause": "UI_NOT_FOUND"},
                     {"form": "mercadopago/after_collection", "cause": "PLATFORM_ERROR"}]}])
    code = env.run([], plan(
        ("download", child(1, "[BATCH_SUMMARY_JSON] " + s.replace("\\", "\\\\"))),
        ("report", child(0))))
    assert code == 1
    assert not os.path.exists(env.alert)
    assert "TANKE_EE" in env.last()[1]["attention"][0]


def test_a_missing_report_input_still_alerts(env, tmp_path):
    s = _summary(tmp_path, stores=[
        {"store": "TRQWH_TA01", "status": "partial", "pending": [],
         "errors": ["stock: 1 张未下载"],
         "missing": [{"form": "stock/general", "cause": "DOWNLOAD_NOT_LANDED"}]}])
    env.run([], plan(
        ("download", child(1, "[BATCH_SUMMARY_JSON] " + s.replace("\\", "\\\\"))),
        ("report", child(0))))
    assert "TRQWH_TA01" in env.alert_text()


def test_report_failure_names_the_store_without_a_download_folder(env):
    """run_reports says which store had no download folder (its last line is a
    machine-readable summary); the alert repeats that instead of guessing
    'crash / timeout'."""
    import json
    summary = {"written": ["BOCINA_SM", "全店汇总"], "missing": [],
               "unusable": [["AIS_CL", "没有下载目录"]], "attention": [], "incomplete": []}
    line = "[REPORT_SUMMARY_JSON] " + json.dumps(summary, ensure_ascii=False)
    assert env.run(["--step", "report"], plan(("report", child(2, line)))) == 2
    text = env.alert_text()
    assert "AIS_CL（没有下载目录）" in text and "先给这几家跑一次下载" in text
    assert "其余 2 份报表已写出" in text
    assert "崩溃 / 超时" not in text
    assert env.last()[1]["report_summary"] == summary
    # a report that was not written is named too
    summary["missing"] = ["EWTTO_SM"]
    line = "[REPORT_SUMMARY_JSON] " + json.dumps(summary, ensure_ascii=False)
    env.run(["--step", "report"], plan(("report", child(2, line))))
    assert "1 份报表没有写出来" in env.alert_text() and "EWTTO_SM" in env.alert_text()
