# -*- coding: utf-8 -*-
"""core/browser_kit: the parts that need no browser."""
import os
import threading
import time

import pytest

from core import browser_kit as bk
from platforms.falabella import forms
from platforms.falabella import page_kit as pk


@pytest.fixture
def fast(monkeypatch):
    """wait_new_files polls once a second; make the clock jump instead."""
    clock = {"t": 1000.0}
    monkeypatch.setattr(bk.time, "time", lambda: clock["t"])

    def sleep(s):
        clock["t"] += s
    monkeypatch.setattr(bk.time, "sleep", sleep)
    return clock


def test_wait_new_files_returns_complete_new_files_only(tmp_path, fast):
    (tmp_path / "old.xlsx").write_text("x")
    before = bk.snapshot(str(tmp_path))
    (tmp_path / "new.xlsx").write_text("x")
    (tmp_path / "other.csv").write_text("x")
    assert bk.wait_new_files(str(tmp_path), before, timeout=20, settle=2) == [
        "new.xlsx", "other.csv"]


def test_wait_new_files_waits_while_a_partial_file_is_there(tmp_path, fast):
    before = bk.snapshot(str(tmp_path))
    (tmp_path / "a.xlsx").write_text("x")
    (tmp_path / "b.xlsx.crdownload").write_text("x")
    # the partial never finishes: at timeout only the complete file is reported
    assert bk.wait_new_files(str(tmp_path), before, timeout=10, settle=2) == ["a.xlsx"]


def test_wait_new_files_gives_up_early_when_nothing_starts(tmp_path, fast):
    t0 = fast["t"]
    assert bk.wait_new_files(str(tmp_path), set(), timeout=90, start_timeout=5) == []
    assert fast["t"] - t0 < 10


def test_wait_new_files_with_a_real_clock(tmp_path):
    before = bk.snapshot(str(tmp_path))
    p = tmp_path / "late.csv.part"

    def write():
        p.write_text("x")
        time.sleep(1.2)
        os.replace(str(p), str(tmp_path / "late.csv"))
    th = threading.Thread(target=write)
    th.start()
    got = bk.wait_new_files(str(tmp_path), before, timeout=15, settle=1)
    th.join()
    assert got == ["late.csv"]


def test_snapshot_of_a_missing_folder_is_empty(tmp_path):
    assert bk.snapshot(str(tmp_path / "nope")) == set()


class _Ctx(object):
    def __init__(self):
        self.lines = []
        self.log = self.lines.append


class Stop(RuntimeError):
    pass


def test_retried_tries_again_after_a_page_error(monkeypatch):
    monkeypatch.setattr(bk.time, "sleep", lambda s: None)
    calls = []

    def fn(ctx, x):
        calls.append(x)
        if len(calls) == 1:
            raise bk.PageError("button not found")
        return "ok"
    ctx = _Ctx()
    assert bk.retried(ctx, "测试", fn, 7) == "ok"
    assert calls == [7, 7]
    assert len(ctx.lines) == 1 and "第 1 次没成功" in ctx.lines[0]


def test_retried_raises_the_last_failure(monkeypatch):
    monkeypatch.setattr(bk.time, "sleep", lambda s: None)
    calls = []

    def fn(ctx):
        calls.append(1)
        raise bk.PageError("still not found")
    with pytest.raises(bk.PageError):
        bk.retried(_Ctx(), "x", fn, attempts=3)
    assert len(calls) == 3


def test_retried_never_retries_the_listed_types(monkeypatch):
    monkeypatch.setattr(bk.time, "sleep", lambda s: None)

    class Fatal(bk.PageError):         # matches retry_on too: never_retry wins
        pass
    calls = []

    def fn(ctx):
        calls.append(1)
        raise Fatal("wrong language")
    with pytest.raises(Fatal):
        bk.retried(_Ctx(), "x", fn, never_retry=(Fatal,))
    assert calls == [1]


def test_retried_does_not_catch_other_errors(monkeypatch):
    monkeypatch.setattr(bk.time, "sleep", lambda s: None)
    calls = []

    def fn(ctx):
        calls.append(1)
        raise Stop("bug")
    with pytest.raises(Stop):
        bk.retried(_Ctx(), "x", fn)
    assert calls == [1]


def test_falabella_retried_keeps_raising_language_wrong(monkeypatch):
    monkeypatch.setattr(bk.time, "sleep", lambda s: None)
    calls = []

    def fn(ctx):
        calls.append(1)
        raise forms.LanguageWrong("界面语言是 en-CL")
    with pytest.raises(forms.LanguageWrong):
        forms.retried(_Ctx(), "x", fn)
    assert calls == [1]


def test_falabella_page_kit_re_exports_the_generic_helpers():
    for name in ("WALK", "PageError", "find", "need", "wait_for", "js_click",
                 "real_click", "shadow", "inside", "table_rows", "row_starting",
                 "row_containing", "PARTIAL", "set_download_dir", "snapshot",
                 "wait_new_files"):
        assert getattr(pk, name) is getattr(bk, name), name


def test_page_lookups_swallow_driver_errors():
    class Dead(object):
        def execute_script(self, *a):
            raise RuntimeError("browser gone")
    d = Dead()
    assert bk.find(d, "button") == []
    assert bk.table_rows(d) == []
    assert bk.row_starting(d, "x") is None
    assert bk.row_containing(d, "x") is None
    assert bk.shadow(d, "div#x", "a") is None
    assert bk.inside(d, object(), "a") is None
    with pytest.raises(bk.PageError) as e:
        bk.need(d, "button", "Aplicar", timeout=0, what="应用按钮")
    assert "not found" in str(e.value)


def test_core_browser_kit_names_no_platform():
    src = open(bk.__file__, encoding="utf-8").read().lower()
    for word in ("sellercenter", "mercadolibre", "español", "listo", "idioma"):
        assert word not in src, word


def test_wait_new_files_ignores_a_partial_that_was_already_downloading(tmp_path, fast):
    """A file still .crdownload before the click belongs to the previous route;
    once the browser renames it, it must not count as this route's download."""
    (tmp_path / "export.xlsx.crdownload").write_bytes(b"x")
    before = bk.snapshot(tmp_path)
    (tmp_path / "export.xlsx.crdownload").unlink()
    (tmp_path / "export.xlsx").write_bytes(b"late file of the previous route")
    (tmp_path / "ours.csv").write_bytes(b"this route's file")
    assert bk.wait_new_files(str(tmp_path), before, timeout=3, settle=1) == ["ours.csv"]
    assert bk.wait_new_files(str(tmp_path), bk.snapshot(tmp_path), timeout=2, settle=1,
                             start_timeout=1) == []
