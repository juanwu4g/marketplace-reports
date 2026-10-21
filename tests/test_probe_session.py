# -*- coding: utf-8 -*-
"""tools/probe_session: the command loop and the dump, with a fake driver.
No store is opened; main() is only tested for refusing before Ziniao starts."""
import io
import json
import os

import pytest

from core import store_config
from tools import probe_session as ps


class FakeDriver(object):
    current_url = "https://seller.example.com/x"
    title = "X"

    def __init__(self):
        self.scripts = []

    def execute_script(self, src, *a):
        self.scripts.append(src)
        return {"hosts": [], "controls": [], "text": ["[d0] hola"]}

    def save_screenshot(self, path):
        open(path, "wb").write(b"png")
        return True


def write(p, text):
    with io.open(str(p), "w", encoding="utf-8") as f:
        f.write(text)


def read(p):
    with io.open(str(p), encoding="utf-8") as f:
        return f.read()


def test_run_command_captures_output_and_traceback(tmp_path):
    ns = {"x": 1}
    write(tmp_path / "a.py", "print('你好', x)\ny = x + 1\n")
    write(tmp_path / "b.py", "print(y)\nraise ValueError('boom')\n")
    ps.run_command(str(tmp_path / "a.py"), ns)
    ps.run_command(str(tmp_path / "b.py"), ns)
    a, b = read(tmp_path / "a.out"), read(tmp_path / "b.out")
    assert a.splitlines() == ["你好 1", ps.DONE_MARK]
    assert b.splitlines()[0] == "2"                     # the namespace persists
    assert "ValueError: boom" in b and b.rstrip().endswith(ps.DONE_MARK)


def test_serve_runs_new_files_once_skips_old_ones_and_stops(tmp_path, monkeypatch):
    cmd = tmp_path / "cmd"
    cmd.mkdir()
    write(cmd / "000_old.py", "raise SystemExit('must not run')\n")
    write(cmd / ps.STOP, "")                            # stale STOP from last time
    ticks = {"n": 0}

    def sleep(s):
        ticks["n"] += 1
        if ticks["n"] == 1:
            assert read(cmd / ps.READY).startswith("ok")
            write(cmd / "002_b.py", "print('b')\n")
            write(cmd / "001_a.py", "print('a')\n")
        elif ticks["n"] == 2:
            write(cmd / ps.STOP, "")
    monkeypatch.setattr(ps.time, "sleep", sleep)
    ran = ps.serve(str(cmd), {}, max_seconds=3600, log=lambda *a: None)
    assert ran == ["001_a.py", "002_b.py"]
    assert not (cmd / "000_old.out").exists()
    assert read(cmd / "001_a.out").startswith("a\n")


def test_serve_stops_at_max_time(tmp_path, monkeypatch):
    clock = {"t": 0.0}
    monkeypatch.setattr(ps.time, "time", lambda: clock["t"])
    monkeypatch.setattr(ps.time, "sleep", lambda s: clock.__setitem__("t", clock["t"] + s))
    assert ps.serve(str(tmp_path), {}, max_seconds=5, log=lambda *a: None) == []
    assert clock["t"] >= 5


def test_dump_writes_screenshot_and_json(tmp_path):
    d = FakeDriver()
    ns = ps.namespace(d, str(tmp_path), str(tmp_path / "files"))
    p1 = ns["dump"](d, "pagos / inicio")
    p2 = ns["dump"](d, "next", {"note": 1})
    assert os.path.basename(p1) == "00_pagos_inicio.json"
    assert os.path.basename(p2) == "01_next.json"
    info = json.loads(read(p2))
    assert info["url"] == d.current_url and info["note"] == 1
    assert info["page"]["text"] == ["[d0] hola"]
    assert (tmp_path / "01_next.png").exists()
    assert "STYLE" in d.scripts[0]                       # style contents are skipped


def test_namespace_offers_the_browser_kit():
    ns = ps.namespace(FakeDriver(), "o", "f")
    for name in ("d", "driver", "find", "need", "wait_for", "js_click", "real_click",
                 "table_rows", "wait_new_files", "snapshot", "WALK", "By", "Keys",
                 "ActionChains", "js", "dump", "time", "os"):
        assert name in ns, name
    assert ps.js(type("Dead", (), {"execute_script": lambda self, *a: 1 / 0})(),
                 "x").startswith("JSERR")


def test_main_refuses_a_store_not_on_the_allowlist_before_ziniao(tmp_path, monkeypatch):
    monkeypatch.setattr(store_config, "_stores",
                        lambda: {"default": "A", "allowed": ["A"]})
    import core.ziniao_client as zc

    def boom(*a, **k):
        raise AssertionError("Ziniao must not be touched")
    monkeypatch.setattr(zc, "ZiniaoClient", boom)
    with pytest.raises(store_config.StoreNotAllowed):
        ps.main(["--store", "SOMEONE_ELSE", "--cmd-dir", str(tmp_path / "c"),
                 "--out", str(tmp_path / "o")])
    assert not (tmp_path / "c").exists()
