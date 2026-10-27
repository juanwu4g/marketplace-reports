# -*- coding: utf-8 -*-
"""运行器的边角：config.json 写坏、数据盘掉线、子进程崩溃、锁的心跳、卡死的
LibreOffice、补取那一轮、旧告警的清理。子进程都是几行 Python，不碰紫鸟。"""
import io
import json
import os
import subprocess
import sys
import time

import pytest

from core import platform_batch
from core import report_kit
from core import run_monthly
from core import runlog
from core import store_config

from test_runner import env, child, plan, _summary, MONTH  # noqa: F401  复用夹具

PY = sys.executable
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ------------------------------------------------------------ F4 config.json 写坏

def test_a_broken_config_does_not_break_the_import(tmp_path):
    bad = tmp_path / "config.json"
    bad.write_text("{ broken", encoding="utf-8")
    src = ("import sys\nsys.path.insert(0, %r)\n"
           "from core import store_config\nstore_config.CONFIG_PATH = %r\n"
           "import core.run_monthly\n"
           "from platforms.mercadolibre import run_reports, financial_report\n"
           "assert run_reports.OUT_DIR is None\n"
           "try:\n    financial_report.DOWNLOADS\nexcept store_config.StoreConfigError:\n"
           "    print('LAZY-OK')\n" % (ROOT, str(bad)))
    r = subprocess.run([PY, "-c", src], cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", timeout=120)
    assert r.returncode == 0, r.stderr
    assert "LAZY-OK" in r.stdout


def test_downloads_is_still_resolvable_for_old_callers():
    from core import layout
    from platforms.mercadolibre import financial_report
    assert financial_report.DOWNLOADS == layout.raw_root()
    with pytest.raises(AttributeError):
        financial_report.NOT_A_THING


def test_a_broken_config_exits_4_with_an_alert_in_the_old_report_folder(tmp_path, monkeypatch):
    bad = tmp_path / "config.json"
    bad.write_text("{ broken", encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(bad))
    monkeypatch.setattr(run_monthly, "ROOT", str(tmp_path))
    monkeypatch.setattr(run_monthly, "GUI_SETTINGS", str(tmp_path / "nope.json"))
    monkeypatch.setattr(run_monthly, "LOG_DIR", str(tmp_path / "logs"))
    monkeypatch.setattr(run_monthly, "LOCK", str(tmp_path / "logs" / "monthly.lock"))
    monkeypatch.setattr(run_monthly, "config", lambda: {})
    assert run_monthly.main(["--month", MONTH]) == 4
    p = tmp_path / "downloads" / "data" / "reports" / "financial" / "！报表异常_202609.txt"
    text = p.read_text(encoding="utf-8")
    assert "config.json 有问题" in text and "StoreConfigError" in text


def test_ml_run_reports_crash_exits_2_and_config_error_4(monkeypatch, capsys):
    from platforms.mercadolibre import run_reports

    def boom():
        raise RuntimeError("pandas 坏了")
    with pytest.raises(SystemExit) as e:
        run_reports.run_as_script(boom)
    assert e.value.code == 2
    assert "RuntimeError: pandas 坏了" in capsys.readouterr().err

    # import 时没算出输出目录：真跑时再算一次，还是坏的就按配置错误退 4
    def bad_root(*a, **k):
        raise store_config.StoreConfigError("config.json 读不出来")
    monkeypatch.setattr(run_reports, "OUT_DIR", None)
    monkeypatch.setattr(run_reports.layout, "reports_root", bad_root)
    with pytest.raises(SystemExit) as e:
        run_reports.run_as_script(lambda: 0)
    assert e.value.code == 4


# ------------------------------------------------------------ F5 平台下载器崩溃

def test_platform_batch_crash_exits_2_not_1(monkeypatch, capsys):
    r = platform_batch.BatchRunner("FA", "FA-CL", forms=None, periods=None,
                                   run_downloads=None, site=None, name_zh="Falabella")

    def boom(argv=None):
        raise KeyError("pending")
    monkeypatch.setattr(r, "_main", boom)
    assert r.main([]) == 2
    out = capsys.readouterr()
    assert "KeyError" in out.err and "按硬失败退出" in out.out


# ------------------------------------------------------------ F16 数据盘掉线

def test_unreachable_data_folder_exits_2_with_alert_in_code_logs(tmp_path, monkeypatch):
    blocker = tmp_path / "offline"
    blocker.write_text("x")                      # 一个文件：它下面建不了文件夹
    monkeypatch.setattr(run_monthly, "ROOT", str(tmp_path / "code"))
    monkeypatch.setattr(run_monthly, "LOG_DIR", str(blocker / "logs"))
    monkeypatch.setattr(run_monthly, "LOCK", str(blocker / "logs" / "monthly.lock"))
    monkeypatch.setattr(run_monthly, "config", lambda: {})
    monkeypatch.setattr(run_monthly, "known_stores", lambda: [])
    code = run_monthly.main(["--month", MONTH, "--out-dir", str(blocker / "reports"),
                             "--step", "report"])
    assert code == 2
    p = tmp_path / "code" / "logs" / "！报表异常_202609.txt"
    assert "文件夹打不开或写不进去" in p.read_text(encoding="utf-8")


# ------------------------------------------------------------ F6 锁

def _write_lock(body, age=0):
    os.makedirs(os.path.dirname(run_monthly.LOCK), exist_ok=True)
    with io.open(run_monthly.LOCK, "w", encoding="utf-8") as f:
        f.write(body)
    if age:
        t = time.time() - age
        os.utime(run_monthly.LOCK, (t, t))


def test_live_pid_with_a_quiet_heartbeat_is_taken_over(env, monkeypatch):
    monkeypatch.setattr(run_monthly, "alive", lambda pid: True)
    _write_lock("pid=424242 起于 x", age=run_monthly.LOCK_QUIET - 120)
    assert run_monthly.acquire(lambda lv, t: None) is False      # 心跳还新：在跑
    _write_lock("pid=424242 起于 x", age=run_monthly.LOCK_QUIET + 60)
    said = []
    assert run_monthly.acquire(lambda lv, t: said.append(t)) is True
    assert "没有心跳" in said[-1]
    run_monthly.release()


def test_empty_lock_is_stale_after_a_grace(env):
    _write_lock("")                              # 对方刚建好、还没写内容
    assert run_monthly.acquire(lambda lv, t: None) is False
    _write_lock("", age=run_monthly.LOCK_EMPTY_GRACE + 5)
    said = []
    assert run_monthly.acquire(lambda lv, t: said.append(t)) is True
    assert "没有进程号" in said[-1]
    run_monthly.release()
    assert not os.path.exists(run_monthly.LOCK)


def test_release_never_removes_another_runs_lock(env):
    _write_lock("pid=424242 起于 x")
    run_monthly.release()
    assert os.path.exists(run_monthly.LOCK)


def test_the_heartbeat_refreshes_the_lock(env, monkeypatch):
    monkeypatch.setattr(run_monthly, "HEARTBEAT", 0.2)
    assert run_monthly.acquire() is True
    old = time.time() - 3600
    os.utime(run_monthly.LOCK, (old, old))

    class Log(object):
        def write(self, *a, **k):
            pass
    run_monthly.run_child(Log(), "report", child(0, "x", sleep=0.8), "rid",
                          log_dir=env.log_dir)
    assert time.time() - os.path.getmtime(run_monthly.LOCK) < 60
    run_monthly.release()


# ------------------------------------------------------------ F7 LibreOffice 卡死

def _fake_soffice(tmp_path, body):
    p = tmp_path / "soffice"
    p.write_text("#!%s\n%s" % (PY, body), encoding="utf-8")
    os.chmod(str(p), 0o755)
    return str(p)


@pytest.mark.skipif(os.name == "nt", reason="fake soffice is a shebang script")
def test_recalc_kills_a_stuck_soffice_and_reports_not_recalculated(tmp_path, monkeypatch):
    pidfile = tmp_path / "grandchild.pid"
    exe = _fake_soffice(tmp_path, (
        "import subprocess, sys, time\n"
        "g = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(60)'])\n"
        "open(%r, 'w').write(str(g.pid))\n"
        "time.sleep(60)\n" % str(pidfile)))
    monkeypatch.setenv("SOFFICE", exe)
    assert report_kit.soffice_path() == exe
    wb = tmp_path / "x.xlsx"
    wb.write_bytes(b"data")
    began = time.time()
    ok, msg = report_kit.recalc(str(wb), timeout=1.5)
    assert not ok and "超过 1.5 秒" in msg
    assert time.time() - began < 20
    assert wb.read_bytes() == b"data"            # 原文件不动
    gpid = int(pidfile.read_text())
    for _ in range(50):
        if not runlog.pid_alive(gpid):
            break
        time.sleep(0.1)
    assert not runlog.pid_alive(gpid), "soffice 起的子进程也要结束"


@pytest.mark.skipif(os.name == "nt", reason="fake soffice is a shebang script")
def test_recalc_uses_the_soffice_env_var(tmp_path, monkeypatch):
    exe = _fake_soffice(tmp_path, (
        "import os, shutil, sys\n"
        "a = sys.argv\nout = a[a.index('--outdir') + 1]\nos.makedirs(out, exist_ok=True)\n"
        "open(os.path.join(out, os.path.basename(a[-1])), 'wb').write(b'recalculated')\n"))
    monkeypatch.setenv("SOFFICE", exe)
    monkeypatch.setenv("PATH", "")
    wb = tmp_path / "x.xlsx"
    wb.write_bytes(b"data")
    assert report_kit.recalc(str(wb)) == (True, "已重算")
    assert wb.read_bytes() == b"recalculated"
    from platforms.mercadolibre import financial_report
    assert financial_report.has_libreoffice() is True


# ------------------------------------------------------------ F8 补取那一轮

def _write_run(log_dir, run_id, stores, rows, step=None):
    os.makedirs(runlog.runs_dir(log_dir), exist_ok=True)
    batch = os.path.join(log_dir, "batch_%s.json" % run_id)
    with open(batch, "w", encoding="utf-8") as f:
        json.dump({"stores": rows}, f, ensure_ascii=False)
    doc = {"run_id": run_id, "month": "2026-09", "stores": stores,
           "status": "done", "exit": 1, "batch_summary": batch}
    if step:
        doc["step"] = step
    runlog.write_json(runlog.run_paths(run_id, log_dir)["json"], doc)
    open(runlog.run_paths(run_id, log_dir)["log"], "w").close()


BLOCKED = [{"store": "A", "status": "partial",
            "missing": [{"form": "ventas", "cause": "PAGE_BLOCKED", "period": None,
                         "explain": "", "evidence": {}}]},
           {"store": "B", "status": "ok", "missing": []}]


def test_a_collect_run_does_not_hide_earlier_human_actions(tmp_path):
    d = str(tmp_path / "logs")
    _write_run(d, "20261014_025744_manual", "all", BLOCKED, step="both")
    assert [a["store"] for a in runlog.pending_human_actions(d)["actions"]] == ["A"]
    # 有人点了「补取欠的报表」：只碰了欠着报表的 B
    _write_run(d, "20261014_120000_manual", "all",
               [{"store": "B", "status": "ok", "missing": []}], step="collect")
    assert [a["store"] for a in runlog.pending_human_actions(d)["actions"]] == ["A"]
    # 一轮真正的全店下载照旧让更早的过时
    _write_run(d, "20261014_130000_manual", "all",
               [{"store": "A", "status": "ok", "missing": []}], step="both")
    assert runlog.pending_human_actions(d)["actions"] == []


# ------------------------------------------------------------ F10 旧告警清理

def test_a_clean_rerun_removes_the_old_alert(env):
    assert env.run([], plan(("download", child(2)), ("report", child(0)))) == 2
    assert os.path.exists(env.alert)
    assert env.run([], plan(("download", child(0)), ("report", child(0)))) == 0
    assert not os.path.exists(env.alert)
    _, _, recs = env.last()
    assert any("删掉了之前的告警文件" in r["message"] for r in recs)


def test_a_rerun_with_urgent_items_keeps_the_alert(env):
    assert env.run([], plan(("download", child(2)), ("report", child(0)))) == 2
    assert env.run([], plan(("download", child(1)), ("report", child(0)))) == 1
    assert os.path.exists(env.alert)             # 笼统的"下载有要关注的项"仍是告警


def test_a_subset_run_does_not_clear_the_month_alert(env):
    os.makedirs(env.out_dir)
    with io.open(env.alert, "w", encoding="utf-8") as f:
        f.write("x")
    msgs = []

    class Log(object):
        def info(self, step, m, store=None):
            msgs.append(m)

        def warn(self, step, m, store=None):
            msgs.append(m)
    args = run_monthly.parse_args(["--month", MONTH, "--out-dir", env.out_dir])
    args.stores = ["A"]
    assert run_monthly.clear_alert(Log(), args, MONTH, env.out_dir, "MX") is None
    assert os.path.exists(env.alert)
    args.reports_for = "all"                     # 重跑这些店并重出全部报表：算数
    assert run_monthly.clear_alert(Log(), args, MONTH, env.out_dir, "MX") == env.alert
    assert not os.path.exists(env.alert)


# ------------------------------------------------------------ F12 补取后还欠着

def test_collect_with_reports_still_owed_is_attention_not_alert(env, tmp_path):
    s = _summary(tmp_path, platform="FA", collect_only=True, stores=[],
                 late={"done": 1, "still_owed": 2}, exit_code=1)
    code = env.run(["--step", "collect"], plan(
        ("collect", child(1, "[BATCH_SUMMARY_JSON] " + s.replace("\\", "\\\\")))))
    assert code == 1
    assert not os.path.exists(env.alert)
    doc = env.last()[1]
    assert doc["urgent"] == []
    assert any("还有 2 张" in a for a in doc["attention"])


# ------------------------------------------------------------ F14 期间报表的时机

def test_period_hook_runs_under_the_lock_and_only_after_a_report(env, monkeypatch):
    seen = []
    monkeypatch.setattr(run_monthly, "period_hook",
                        lambda *a, **k: seen.append(os.path.exists(run_monthly.LOCK)))
    assert env.run([], plan(("download", child(0)), ("report", child(0)))) == 0
    assert seen == [True]
    assert not os.path.exists(run_monthly.LOCK)
    # 下载硬失败、报表那一步没跑：不生成期间报表
    assert env.run([], plan(("download", child(2)), ("report", child(0)))) == 2
    assert seen == [True]
