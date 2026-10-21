# -*- coding: utf-8 -*-
"""tools/update.py: it refuses before changing anything when a monthly run is
live or code was edited locally, and every later failure names the way back.
git, pip and the checks are faked; nothing is pulled or installed."""
import pytest

from tools import update


@pytest.fixture
def fake(monkeypatch):
    calls = []
    state = {"commit": ["aaaaaaa", "bbbbbbb"], "codes": {}}

    def run(cmd, env=None, quiet=False):
        calls.append(cmd)
        key = cmd[1] if cmd[0] == "git" else (cmd[2] if cmd[1] == "-m" else cmd[1])
        if cmd[:2] == ["git", "status"]:
            return 0, state.get("status", "")
        return state["codes"].get(key, 0), ""

    def commit(root=None):
        c = state["commit"]
        return c[0] if len(c) == 1 or not any(x[:2] == ["git", "pull"] for x in calls) else c[1]
    monkeypatch.setattr(update, "run", run)
    monkeypatch.setattr(update.runlog, "git_commit", commit)
    monkeypatch.setattr(update, "monthly_running", lambda: None)
    monkeypatch.setattr(update.os.path, "exists", lambda p: True)
    state["calls"] = calls
    return state


def test_refuses_during_a_monthly_run(fake, monkeypatch, capsys):
    monkeypatch.setattr(update, "monthly_running", lambda: "pid=42 起于 今天")
    assert update.main([]) == 2
    assert fake["calls"] == []                         # nothing was touched
    assert "正在运行" in capsys.readouterr().out


def test_refuses_with_local_code_edits(fake, capsys):
    fake["status"] = " M core/run_monthly.py\n"
    assert update.main([]) == 2
    assert not any(c[:2] == ["git", "pull"] for c in fake["calls"])
    assert "core/run_monthly.py" in capsys.readouterr().out


def test_happy_path_runs_every_step_in_order(fake, capsys):
    assert update.main([]) == 0
    steps = [c[1] if c[0] == "git" else (c[2] if c[1] == "-m" else c[1]) for c in fake["calls"]]
    assert steps == ["status", "pull", "pip", "bootstrap.py", "run_monthly.py"]
    assert "aaaaaaa -> bbbbbbb" in capsys.readouterr().out


def test_failed_check_after_pull_names_the_way_back(fake, capsys):
    fake["codes"]["bootstrap.py"] = 1
    assert update.main([]) == 1
    assert "git checkout aaaaaaa" in capsys.readouterr().out


def test_failed_pull_changes_nothing(fake, capsys):
    fake["codes"]["pull"] = 1
    assert update.main([]) == 1
    assert not any("pip" in c for c in fake["calls"])
