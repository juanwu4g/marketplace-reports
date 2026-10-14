# -*- coding: utf-8 -*-
"""transform.build 怎么报清洗结果：本次目录里的失败 vs 以前目录里的失败。

清洗每次都重扫整棵下载树，所以一个洗不动的旧文件会让之后每一次都"部分失败"。
这里用桩替掉子进程，只测 transform 拿到 build 的输出后怎么判。
"""
import json

import pytest

import transform

OLD = {"report": "ventas_mx", "file": "A/20260815_090000/v.xlsx",
       "error": "v.xlsx: ventas_mx: no usable header", "current": False}
NEW = {"report": "ventas_mx", "file": "A/20261001_090000_m202609/v.xlsx",
       "error": "v.xlsx: boom", "current": True}


class Done(object):
    def __init__(self, returncode, stdout):
        self.returncode, self.stdout = returncode, stdout


@pytest.fixture
def fake(monkeypatch):
    seen = {}

    def install(returncode, failed=None, written=5, marker=True):
        def run(cmd, **kwargs):
            seen["cmd"] = cmd
            runs = [cmd[i + 1] for i, a in enumerate(cmd) if a == "--current-run"]
            out = "ventas_mx: 10 rows from %d file(s)\n" % written
            if marker:
                rows = [dict(f, current=(not runs) or f["current"]) for f in failed or []]
                out += transform.BUILD_RESULT_MARKER + json.dumps(
                    {"written": written, "failed": rows, "current_runs": runs}) + "\n"
            return Done(returncode, out)

        monkeypatch.setattr(transform, "available", lambda: True)
        monkeypatch.setattr(transform.subprocess, "run", run)
        return seen

    return install


def test_default_call_is_unchanged(fake, capsys):
    seen = fake(1, failed=[OLD])
    r = transform.build()

    assert "--current-run" not in seen["cmd"]
    # 没说哪些是本次的：和以前一样，失败就是部分失败。
    assert (r["ok"], r["partial"], r["returncode"]) == (False, True, 1)
    assert r["older_only"] is False and r["failed_older"] == []
    assert len(r["failed_current"]) == 1
    out = capsys.readouterr().out
    assert "部分文件失败" in out
    assert transform.BUILD_RESULT_MARKER.strip() not in out     # 机器行不给人看


def test_all_clean(fake):
    fake(0)
    r = transform.build(current_runs=["20261001_090000_m202609"])

    assert (r["ok"], r["partial"]) == (True, False)
    assert r["files_ok"] == 5 and r["files_failed"] == 0
    assert r["failed_current"] == [] and r["failed_older"] == []


def test_only_older_files_failed_is_ok(fake, capsys):
    seen = fake(1, failed=[OLD])
    r = transform.build(current_runs="20261001_090000_m202609")

    assert seen["cmd"][-2:] == ["--current-run", "20261001_090000_m202609"]
    assert (r["ok"], r["partial"], r["older_only"]) == (True, False, True)
    assert r["returncode"] == 1                 # build 自己的退出码原样留着
    assert [f["file"] for f in r["failed_older"]] == [OLD["file"]]
    assert r["failed_current"] == [] and r["files_failed"] == 1
    assert "以前的文件" in capsys.readouterr().out


def test_a_failure_from_this_run_is_still_partial(fake, capsys):
    fake(1, failed=[OLD, NEW])
    r = transform.build(current_runs=["20261001_090000_m202609"])

    assert (r["ok"], r["partial"], r["older_only"]) == (False, True, False)
    assert [f["file"] for f in r["failed_current"]] == [NEW["file"]]
    assert [f["file"] for f in r["failed_older"]] == [OLD["file"]]
    assert "本次 1 个，以前的 1 个" in capsys.readouterr().out


def test_run_folder_may_be_given_as_a_path(fake):
    seen = fake(0)
    transform.build(current_runs=["downloads/A/20261001_090000_m202609/"])

    assert seen["cmd"][-1] == "20261001_090000_m202609"


def test_nothing_written_is_a_failure_whatever_the_split(fake):
    fake(2, failed=[OLD], written=0)
    r = transform.build(current_runs=["20261001_090000_m202609"])

    assert (r["ok"], r["partial"], r["older_only"]) == (False, False, False)


def test_unreadable_marker_falls_back_to_partial(fake):
    fake(1, marker=False)
    r = transform.build(current_runs=["20261001_090000_m202609"])

    # 分不清是哪一堆，就不能替它说"都是旧的"。
    assert (r["ok"], r["partial"], r["older_only"]) == (False, True, False)
