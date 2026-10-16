# -*- coding: utf-8 -*-
"""卡住的那一轮：要看得见、停得掉。

2026-10-13 在 Windows 上实测到的情况：运行器挂死（进程还在、不再有任何动静），
界面一边放开了按钮，一边还显示"正在跑"，"中止"按了没反应。
"""
import os
import subprocess
import sys
import time

from core import runlog


def _sleeper():
    return subprocess.Popen([sys.executable, "-c", "import time; time.sleep(120)"])


def _setup(tmp_path, pid, quiet=0, run_id="20261013_153250_manual"):
    log_dir = str(tmp_path / "logs")
    paths = runlog.run_paths(run_id, log_dir)
    os.makedirs(os.path.dirname(paths["log"]))
    with open(paths["log"], "w", encoding="utf-8") as f:
        f.write("")
    runlog.write_json(paths["json"], {"run_id": run_id, "entrance": "manual",
                                      "pid": pid, "month": "2026-09",
                                      "status": "running", "exit": None})
    runlog.write_current({"run_id": run_id, "pid": pid, "entrance": "manual",
                          "log": paths["log"], "month": "2026-09"}, log_dir)
    if quiet:
        t = time.time() - quiet
        os.utime(runlog.current_path(log_dir), (t, t))
    return log_dir, run_id, paths


def _dead(p, within=10):
    try:
        p.wait(timeout=within)
        return True
    except subprocess.TimeoutExpired:
        return False


def test_live_pid_with_stale_heartbeat_is_stalled_not_absent(tmp_path):
    p = _sleeper()
    try:
        log_dir, run_id, _ = _setup(tmp_path, p.pid,
                                    quiet=runlog.HEARTBEAT_STALE + 600)
        cur = runlog.active_run(log_dir)
        assert cur and cur["stalled"] and cur["quiet_seconds"] >= 600
        st = runlog.run_state(run_id, log_dir)
        # 仍然占着锁和紫鸟：界面要继续锁按钮，所以 state 还是 running
        assert st["state"] == "running" and st["stalled"] and st["kind"] == "bad"
        assert "没有响应" in st["text"] and "中止" in st["text"]
    finally:
        p.kill()


def test_healthy_run_is_not_stalled(tmp_path):
    p = _sleeper()
    try:
        log_dir, run_id, _ = _setup(tmp_path, p.pid)
        assert "stalled" not in runlog.active_run(log_dir)
        assert not runlog.run_state(run_id, log_dir).get("stalled")
    finally:
        p.kill()


def test_dead_pid_is_never_active(tmp_path):
    p = _sleeper()
    p.kill()
    p.wait()
    log_dir, run_id, _ = _setup(tmp_path, p.pid, quiet=runlog.HEARTBEAT_STALE + 5)
    assert runlog.active_run(log_dir) is None
    assert runlog.run_state(run_id, log_dir)["state"] == "interrupted"


def test_very_old_heartbeat_stops_trusting_the_pid(tmp_path):
    p = _sleeper()
    try:
        log_dir, _, _ = _setup(tmp_path, p.pid, quiet=runlog.STALL_GIVEUP + 60)
        assert runlog.active_run(log_dir) is None
    finally:
        p.kill()


def test_stop_on_a_stalled_run_is_forced_at_once(tmp_path):
    p = _sleeper()
    log_dir, run_id, paths = _setup(tmp_path, p.pid,
                                    quiet=runlog.HEARTBEAT_STALE + 60)
    assert runlog.enforce_stop(log_dir) is None, "没人点中止就不该动它"
    runlog.request_stop(run_id, "启动窗口", log_dir)
    assert runlog.enforce_stop(log_dir) == run_id
    assert _dead(p), "卡住的运行器必须被结束"
    assert runlog.active_run(log_dir) is None
    assert not runlog.stop_requested(run_id, log_dir)
    st = runlog.run_state(run_id, log_dir)
    assert st["state"] == "aborted" and st["exit"] == 6
    rec = runlog.parse_line(open(paths["log"], encoding="utf-8").read().strip())
    assert rec["step"] == "收尾" and rec["level"] == "错误"
    assert "强行结束" in rec["message"] and "启动窗口" in rec["message"]


def test_responsive_run_gets_its_grace_period_first(tmp_path):
    p = _sleeper()
    try:
        log_dir, run_id, paths = _setup(tmp_path, p.pid)
        runlog.request_stop(run_id, "定时设置窗口", log_dir)
        assert runlog.enforce_stop(log_dir) is None, "刚点中止：先让它自己停"
        assert p.poll() is None
        t = time.time() - runlog.STOP_GRACE - 5
        os.utime(paths["stop"], (t, t))
        assert runlog.enforce_stop(log_dir) == run_id
        assert _dead(p)
    finally:
        if p.poll() is None:
            p.kill()


def test_force_stop_takes_the_children_too(tmp_path):
    code = ("import subprocess, sys, time\n"
            "g = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(120)'])\n"
            "print(g.pid, flush=True)\n"
            "time.sleep(120)\n")
    p = subprocess.Popen([sys.executable, "-c", code], stdout=subprocess.PIPE)
    child = int(p.stdout.readline())
    log_dir, run_id, _ = _setup(tmp_path, p.pid, quiet=runlog.HEARTBEAT_STALE + 60)
    runlog.request_stop(run_id, "测试", log_dir)
    assert runlog.enforce_stop(log_dir) == run_id
    assert _dead(p)
    end = time.time() + 10
    while time.time() < end and runlog.pid_alive(child):
        try:                                   # 成了僵尸也算结束
            with open("/proc/%d/stat" % child) as f:
                if f.read().split(")")[-1].split()[0] == "Z":
                    break
        except (IOError, OSError):
            break
        time.sleep(0.1)
    else:
        assert not runlog.pid_alive(child), "孙进程也要结束"


def test_force_stop_that_cannot_kill_does_not_pretend(tmp_path, monkeypatch):
    """实测：那一轮以管理员权限起、窗口不是，taskkill 杀不掉。不能把它标成已中止。"""
    p = _sleeper()
    try:
        log_dir, run_id, paths = _setup(tmp_path, p.pid,
                                        quiet=runlog.HEARTBEAT_STALE + 60)
        monkeypatch.setattr(runlog, "kill_pid_tree", lambda pid: False)
        runlog.request_stop(run_id, "启动窗口", log_dir)
        assert runlog.enforce_stop(log_dir) is None
        assert runlog.active_run(log_dir)["stalled"], "它还在，就还得显示它在"
        assert runlog.run_state(run_id, log_dir)["state"] == "running"
        assert not runlog.stop_requested(run_id, log_dir), "撤掉请求，不要每秒重试"
        text = open(paths["log"], encoding="utf-8").read()
        assert "结束不了进程" in text and text.count("\n") == 1
    finally:
        p.kill()
