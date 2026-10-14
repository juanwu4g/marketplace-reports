# -*- coding: utf-8 -*-
"""收取循环、补收、登记簿、批次结论。浏览器、紫鸟、时钟全是假的。"""
import argparse
import contextlib
import io
import json
import os

import pytest

import download_diag as dd
import meli_forms
import mercadopago
import pending_store
import run_batch
import run_downloads
import ziniao_client


# ------------------------------------------------------------------ 假的世界

class Clock(object):
    """time.time / time.sleep 的替身：只有 sleep 会让时间往前走。"""

    def __init__(self, now=1_800_000_000.0):
        self.now = now
        self.sleeps = []

    def time(self):
        return self.now

    def sleep(self, s):
        self.sleeps.append(s)
        self.now += s


class Driver(object):
    current_url = "https://www.mercadopago.com.mx/balance/reports/x"
    title = "Reportes"

    def __init__(self):
        self.shots = 0
        self.visited = []

    def get(self, url):
        self.visited.append(url)

    def execute_script(self, js, *a):
        return {"lang": "es-MX", "body": "", "alerts": []}

    def save_screenshot(self, path):
        self.shots += 1
        with open(path, "wb") as f:
            f.write(b"png")
        return True


def mp(report, rng="1/ago/2026 a 1/oct/2026", **kw):
    e = {"kind": "mercadopago", "report": report, "range": rng,
         "url": "https://mp/" + report, "label": report, "fmt": "csv",
         "requested_at": 1_800_000_000.0 - 3600}
    e.update(kw)
    return e


def new_result():
    return {"store": "S", "sales": [], "mp": [], "pending": [], "empty": [],
            "errors": []}


@pytest.fixture
def clock(monkeypatch):
    c = Clock()
    monkeypatch.setattr(run_downloads.time, "time", c.time)
    monkeypatch.setattr(run_downloads.time, "sleep", c.sleep)
    return c


def script(monkeypatch, plan, fetched=None):
    """mercadopago.probe/fetch 的替身。plan: {报表: [每次探测的回答...]}。

    回答可以是状态字符串、(状态, _probe 字典)，或一个要抛出的异常。
    最后一个回答会一直重复。
    """
    calls = {}

    def probe(d, p, settle=12):
        seq = plan[p["report"]]
        i = calls.get(p["report"], 0)
        calls[p["report"]] = i + 1
        ans = seq[min(i, len(seq) - 1)]
        if isinstance(ans, Exception):
            raise ans
        state, info = ans if isinstance(ans, tuple) else (ans, None)
        default = {"ready": "ready", "empty": "empty", "delayed": "delayed"}.get(
            state, "generating")
        p["_probe"] = info or {"row_state": default, "status_text": "En preparación"}
        return state

    def fetch(d, p, out_dir):
        files = (fetched or {}).get(p["report"], [])
        p["_fetch"] = "clicked"
        for f in files:                   # 真的落盘：诊断会去核对文件在不在
            with open(os.path.join(out_dir, f), "w") as fh:
                fh.write("data")
        return list(files)

    monkeypatch.setattr(mercadopago, "probe", probe)
    monkeypatch.setattr(mercadopago, "fetch", fetch)
    return calls


# ------------------------------------------------------------------ 收取循环

def test_probe_that_raises_stays_owed_with_the_exception_as_evidence(
        monkeypatch, clock, tmp_path):                                    # fix a
    script(monkeypatch, {"release": [KeyError("range")],
                         "collection": ["ready"]},
           fetched={"collection": ["collection-1.csv"]})
    a, b = mp("release"), mp("collection")
    result = new_result()
    left = run_downloads.collect_pending(Driver(), [a, b], str(tmp_path),
                                         result, budget=0)
    assert left == [a]                       # 以前是 []：补收会把它从登记簿删掉
    assert result["mp"] == ["collection-1.csv"]
    assert result["collected"] == {"mercadopago/collection": ["collection-1.csv"]}
    assert any("release/collect: KeyError" in x for x in result["errors"])
    assert a["last_state"] == "probe-error" and a["probes"] == 1
    assert a["evidence"]["exception_class"] == "KeyError"
    rec = dd.record_from_entry(a)
    assert rec["cause"] == "INTERNAL_ERROR" and "KeyError" in rec["explain"]
    assert not any(k.startswith("_") for k in a)       # 临时键不进登记簿
    json.dumps(a)


def test_last_sleep_is_clamped_and_one_final_probe_follows(
        monkeypatch, clock, tmp_path, capsys):                            # fix c
    calls = script(monkeypatch, {"release": ["waiting"]})
    d = Driver()
    e = mp("release")
    left = run_downloads.collect_pending(d, [e], str(tmp_path), new_result(),
                                         budget=30, poll=20)
    assert left == [e]
    assert clock.sleeps == [20, 10]          # 以前是 [20, 20]，然后印出"剩余 -10 秒"
    assert calls["release"] == 3             # 到点之后还有最后一次探测
    out = capsys.readouterr().out
    assert "剩余 30 秒" in out and "剩余 10 秒" in out and "剩余 0 秒" in out
    assert "剩余 -" not in out
    # 截图只在最后一次探测时拍
    assert d.shots == 1 and e["probes"] == 3
    assert e["evidence"]["screenshot"].startswith("_diag/")
    assert (e["first_probe_at"], e["last_probe_at"]) == (
        clock.now - 30, clock.now)


def test_budget_zero_scans_once_without_sleeping(monkeypatch, clock, tmp_path):
    calls = script(monkeypatch, {"release": ["waiting"], "withdraw": ["waiting"]})
    left = run_downloads.collect_pending(
        Driver(), [mp("release"), mp("withdraw")], str(tmp_path), new_result(),
        budget=0)
    assert len(left) == 2 and clock.sleeps == []
    assert calls == {"release": 1, "withdraw": 1}


def test_ready_row_whose_download_fails_is_not_called_not_ready(
        monkeypatch, clock, tmp_path, capsys):                            # fix d
    script(monkeypatch, {"release": ["ready"]}, fetched={"release": []})
    e = mp("release")
    left = run_downloads.collect_pending(Driver(), [e], str(tmp_path),
                                         new_result(), budget=0)
    out = capsys.readouterr().out
    assert left == [e]
    assert "尚未就绪" not in out and "文件没有落盘" in out
    assert e["last_state"] == "ready"
    assert dd.record_from_entry(e)["cause"] == "DOWNLOAD_NOT_LANDED"


def test_not_ready_line_says_what_was_seen(monkeypatch, clock, tmp_path, capsys):
    script(monkeypatch, {"release": [("waiting", {"row_state": "absent"})],
                         "account_statement": [
                             ("waiting", {"row_state": "accordion",
                                          "action": "generar"})]})
    a, b = mp("release"), mp("account_statement", flavor="accordion")
    run_downloads.collect_pending(Driver(), [a, b], str(tmp_path), new_result(),
                                  budget=0)
    out = capsys.readouterr().out
    assert "列表里还没有这一行" in out and "折叠面板：generar" in out
    assert dd.record_from_entry(b)["cause"] == "REQUEST_NOT_REGISTERED"


def test_terminal_answers_leave_the_cycle_with_their_state(
        monkeypatch, clock, tmp_path):
    script(monkeypatch, {"withdraw": ["empty"], "after_collection": ["delayed"]})
    a, b = mp("withdraw"), mp("after_collection")
    result = new_result()
    left = run_downloads.collect_pending(Driver(), [a, b], str(tmp_path), result,
                                         budget=600)
    assert left == [] and clock.sleeps == []
    assert result["no_data"] == [a] and result["retry_later"] == [b]
    assert dd.record_from_entry(a)["cause"] == "NO_DATA"
    assert dd.record_from_entry(b)["cause"] == "PLATFORM_DELAYED"
    assert run_downloads.DELAYED_MARK in result["pending"][0]


def test_ventas_probe_and_fetch_are_diagnosed_too(monkeypatch, clock, tmp_path):
    def probe(d, p, settle=10):
        p["_probe"] = {"row_state": "ventas-widget", "widget_rows": 4,
                       "rows_before": 3, "top_sub": "Generando", "top_has_dl": False}
        return False
    monkeypatch.setattr(meli_forms, "probe_sales_excel", probe)
    e = {"kind": "ventas", "rows_before": 3, "requested_at": clock.now - 100}
    left = run_downloads.collect_pending(Driver(), [e], str(tmp_path),
                                         new_result(), budget=0)
    assert left == [e] and e["last_status_text"] == "Generando"
    rec = dd.record_from_entry(e)
    assert (rec["form"], rec["cause"]) == ("ventas", "STILL_GENERATING")


# ------------------------------------------------------------------ 标签（fix e）

def test_late_ventas_is_labelled_by_its_real_form_and_cause():
    ventas = {"kind": "ventas", "requested_at": 1.0}
    dd.note_probe(ventas, {"row_state": "ventas-widget", "widget_rows": 2,
                           "rows_before": 1, "top_sub": "Generando"}, now=100.0)
    label = run_batch.owed_label(ventas)
    assert label.startswith("ventas - [STILL_GENERATING]")
    assert "None" not in label and "mercadopago" not in label
    gone = mp("release", requested_at=0.0)
    dd.note_probe(gone, {"row_state": "absent"}, now=1000.0)
    dd.note_probe(gone, {"row_state": "absent"}, now=2000.0)
    assert run_batch.owed_label(gone) == (
        "mercadopago/release (1/ago/2026 a 1/oct/2026) - [REQUEST_NOT_REGISTERED] "
        + dd.CAUSES["REQUEST_NOT_REGISTERED"][0])


# ------------------------------------------------------------------ 登记簿（fix g）

def test_prune_records_which_entries_expired(tmp_path):
    path = str(tmp_path / "pending.json")
    now = 2_000_000_000.0
    old = dict(mp("release"), store="A", requested_at=now - 8 * 86400,
               last_state="absent", probes=4, out_dir="/x")
    fresh = dict(mp("withdraw"), store="A", requested_at=now - 86400)
    pending_store.save([old, fresh], path)
    stale = pending_store.prune(path=path, now=now)
    assert [e["report"] for e in stale] == ["release"]
    assert [e["report"] for e in pending_store.load(path)] == ["withdraw"]
    kept = json.load(io.open(pending_store.expired_path(path), encoding="utf-8"))
    assert kept["entries"][0]["cause"] == "EXPIRED"
    assert kept["entries"][0]["expired_at"] == now
    assert kept["entries"][0]["last_state"] == "absent"
    # 追加，不是覆盖
    pending_store.save([dict(old, report="collection")], path)
    pending_store.prune(path=path, now=now)
    kept = json.load(io.open(pending_store.expired_path(path), encoding="utf-8"))
    assert [e["report"] for e in kept["entries"]] == ["release", "collection"]
    rec = dd.expired_record(stale[0])
    assert (rec["cause"], rec["store"], rec["form"]) == (
        "EXPIRED", "A", "mercadopago/release")
    assert pending_store.describe_entry(old) == (
        "A mercadopago/release（1/ago/2026 a 1/oct/2026）")


def test_prune_with_nothing_stale_writes_nothing(tmp_path):
    path = str(tmp_path / "pending.json")
    pending_store.save([dict(mp("release"), store="A", requested_at=100.0)], path)
    assert pending_store.prune(path=path, now=200.0) == []
    assert not os.path.exists(pending_store.expired_path(path))


# ------------------------------------------------------------------ 退出码

def row(store, status="ok", missing=(), pending=()):
    return {"store": store, "status": status, "files": 3, "errors": [],
            "pending": list(pending), "empty": [], "missing": list(missing)}


def rec(cause, form="mercadopago/release"):
    ev = {"NO_DATA": {"row_state": "empty"},
          "NOT_AVAILABLE": {"signal": "not-offered"},
          "STILL_GENERATING": {"row_state": "generating"},
          "UNKNOWN": {}}[cause]
    out = dd.classify(ev, {"form": form})
    assert out["cause"] == cause
    return out


OK_TR = {"ok": True, "partial": False, "skipped": None}
EXITS = [
    ("全部正常", [row("A"), row("B")], OK_TR, 0),
    ("没有清洗这一步", [row("A")], None, 0),
    ("清洗被跳过", [row("A")], {"ok": True, "partial": False, "skipped": "没有包"}, 0),
    ("缺的只是没数据 / 平台不提供",
     [row("A", missing=[rec("NO_DATA"), rec("NOT_AVAILABLE", "stock/costos/2")])],
     OK_TR, 0),
    ("有报表还欠着", [row("A", missing=[rec("STILL_GENERATING")]), row("B")], OK_TR, 1),
    ("原因不明也要人看", [row("A", missing=[rec("UNKNOWN", "stock/general")])], OK_TR, 1),
    ("一家 partial", [row("A", "partial"), row("B")], OK_TR, 1),
    ("一家 failed，另一家成功", [row("A", "failed"), row("B")], OK_TR, 1),
    ("一家 failed，另一家 partial", [row("A", "failed"), row("B", "partial")], OK_TR, 1),
    ("全部 failed", [row("A", "failed"), row("B", "failed")], OK_TR, 2),
    ("不在白名单也算没成", [row("A", "not-allowed")], OK_TR, 2),
    ("清洗部分失败", [row("A")], {"ok": False, "partial": True, "skipped": None}, 1),
    ("清洗整个失败", [row("A")], {"ok": False, "partial": False, "skipped": None}, 1),
    ("全部 failed 压过清洗", [row("A", "failed")],
     {"ok": False, "partial": True, "skipped": None}, 2),
]


@pytest.mark.parametrize("why,rows,tr,want", EXITS, ids=[e[0] for e in EXITS])
def test_batch_exit_code(why, rows, tr, want):
    assert run_batch.batch_verdict(rows, tr)["exit_code"] == want


def test_cleaning_failure_is_its_own_thing_not_a_store_failure():
    v = run_batch.batch_verdict([row("A"), row("B")],
                                {"ok": False, "partial": True, "skipped": None})
    assert v["exit_code"] == 1 and v["cleaning"] == "partial"
    assert (v["stores_total"], v["stores_ok"], v["stores_attention"],
            v["stores_failed"]) == (2, 2, 0, 0)


def test_verdict_counts_stores_and_causes():
    rows = [row("A", missing=[rec("STILL_GENERATING"), rec("NO_DATA", "x/y")]),
            row("B", "failed", missing=[rec("UNKNOWN", "ventas")]),
            row("C")]
    expired = [dd.classify({"signal": "expired"}, {"form": "mercadopago/release"})]
    v = run_batch.batch_verdict(rows, OK_TR, expired)
    assert (v["stores_total"], v["stores_ok"], v["stores_attention"],
            v["stores_failed"]) == (3, 1, 1, 1)
    assert v["missing_by_cause"] == {"STILL_GENERATING": 1, "NO_DATA": 1,
                                     "UNKNOWN": 1, "EXPIRED": 1}
    assert v["exit_code"] == 1


def test_summary_json_shape(tmp_path):
    rows = [dict(row("A", missing=[rec("STILL_GENERATING")]),
                 _result={"not": "serialisable", "x": object()}, _late={})]
    tr = {"ok": False, "partial": True, "skipped": None}
    path = str(tmp_path / "logs" / "batch_x.json")
    run_batch.write_summary(path, "x", rows, transform_result=tr,
                            extra=run_batch.batch_verdict(rows, tr))
    doc = json.load(io.open(path, encoding="utf-8"))
    for key in ("started", "stores", "transform", "missing_by_cause",
                "stores_total", "stores_ok", "stores_attention", "stores_failed"):
        assert key in doc
    assert doc["missing_by_cause"] == {"STILL_GENERATING": 1}
    assert doc["cleaning"] == "partial" and doc["exit_code"] == 1
    store = doc["stores"][0]
    assert not any(k.startswith("_") for k in store)
    m = store["missing"][0]
    for key in ("form", "period", "cause", "explain", "owner", "retry", "action",
                "evidence", "first_seen", "last_seen", "probes"):
        assert key in m


# ------------------------------------------------------------------ 补收

class Session(object):
    launcher_page = "about:launcher"

    def __init__(self):
        self.driver = Driver()


class Client(object):
    def __init__(self, fail=None):
        self.fail = fail
        self.opened = []

    def find_store(self, name):
        return {"browserName": name, "browserOauth": name}

    @contextlib.contextmanager
    def open_store(self, store, close_when_done=True):
        if self.fail:
            raise self.fail
        self.opened.append(store["browserName"])
        yield Session()


@pytest.fixture
def registry(monkeypatch, tmp_path):
    """把登记簿挪到临时目录。默认参数在定义时就绑死了，所以逐个函数包一层。"""
    path = str(tmp_path / "logs" / "pending.json")
    for name, nargs in (("load", 0), ("save", 1), ("add", 1), ("remove", 1),
                        ("for_store", 1)):
        real = getattr(pending_store, name)

        def wrapped(*a, _real=real, _n=nargs, **k):
            if len(a) <= _n:
                k.setdefault("path", path)
            return _real(*a, **k)
        monkeypatch.setattr(pending_store, name, wrapped)
    monkeypatch.setattr(meli_forms, "set_download_dir", lambda d, p: True)
    return path


def late_args(tmp_path):
    return argparse.Namespace(late_timeout=0, collect_poll=1, out=str(tmp_path),
                              _stamp="20261013_090000")


def test_late_pass_keeps_an_entry_whose_probe_raised(
        monkeypatch, clock, registry, tmp_path):                 # fix a，端到端
    folder = str(tmp_path / "A" / "run")
    os.makedirs(folder)
    pending_store.add([dict(mp("release"), store="A", out_dir=folder),
                       dict(mp("collection"), store="A", out_dir=folder),
                       {"kind": "ventas", "store": "A", "out_dir": folder,
                        "rows_before": 1, "requested_at": clock.now - 900}])
    script(monkeypatch, {"release": [RuntimeError("stale element")],
                         "collection": ["ready"]},
           fetched={"collection": ["collection-9.csv"]})

    def ventas_probe(d, p, settle=10):
        p["_probe"] = {"row_state": "ventas-widget", "widget_rows": 2,
                       "rows_before": 1, "top_sub": "Generando"}
        return False
    monkeypatch.setattr(meli_forms, "probe_sales_excel", ventas_probe)

    r = dict(row("A"), out_dir=folder)
    run_batch.collect_pass(Client(), ["A"], late_args(tmp_path), [r])

    left = pending_store.load()
    assert sorted(dd.form_of_entry(e) for e in left) == [
        "mercadopago/release", "ventas"]
    by = {dd.form_of_entry(e): e for e in left}
    # 登记簿条目上的诊断字段
    for key in ("probes", "first_probe_at", "last_probe_at", "last_state",
                "last_status_text", "cause", "evidence"):
        assert key in by["mercadopago/release"], key
    assert by["mercadopago/release"]["cause"] == "INTERNAL_ERROR"
    assert by["ventas"]["cause"] == "STILL_GENERATING"
    # 汇总里每条按真实的报表和原因写（fix e）
    assert sorted(x.split(" - ")[0] for x in r["pending"]) == [
        "mercadopago/release (1/ago/2026 a 1/oct/2026)", "ventas"]
    assert not any("None" in x or "still generating" in x for x in r["pending"])
    assert r["status"] == "partial" and r["late_collected"] == 1
    assert r["_late"]["by_dir"][folder]["mp"] == ["collection-9.csv"]


def test_late_pass_store_that_will_not_open_keeps_everything_with_a_cause(
        monkeypatch, clock, registry, tmp_path):
    pending_store.add([dict(mp("release"), store="A", out_dir=str(tmp_path))])
    err = ziniao_client.ZiniaoClient._error(
        "startBrowser failed: {...}", "startBrowser", '{"statusCode": -1}')
    r = dict(row("A"), out_dir=str(tmp_path))
    run_batch.collect_pass(Client(fail=err), ["A"], late_args(tmp_path), [r])
    (e,) = pending_store.load()
    assert e["cause"] == "STORE_OPEN_FAILED"
    assert e["evidence"]["step"] == "startBrowser"
    assert e["evidence"]["ziniao_response"] == '{"statusCode": -1}'
    assert "[STORE_OPEN_FAILED]" in r["pending"][0]


def test_final_verdict_rewrites_report_and_meta_from_late_evidence(
        monkeypatch, clock, registry, tmp_path):
    folder = str(tmp_path / "A" / "run")
    os.makedirs(folder)
    stock = dd.classify({"signal": "ui-missing"}, {"form": "stock/general"})
    was_unknown = dict(dd.classify({"row_state": "absent"},
                                   {"form": "mercadopago/release"}), owed=True)
    got_late = dict(dd.classify({"row_state": "generating"},
                                {"form": "mercadopago/collection"}), owed=True)
    result = {"store": "A", "out_dir": folder, "billing": {}, "sales": [],
              "stock": {}, "mp": [], "pending": [], "empty": [], "errors": [],
              "seconds": 1, "expected": [], "flags": {}, "cautions": [],
              "missing": [stock, was_unknown, got_late]}
    run_downloads.write_run_meta(folder, result, clock.now)
    still = dict(mp("release", requested_at=clock.now - 3600), store="A",
                 out_dir=folder)
    dd.note_probe(still, {"row_state": "absent"}, now=clock.now - 3000)
    dd.note_probe(still, {"row_state": "absent"}, now=clock.now)
    pending_store.add([still])
    r = dict(row("A"), out_dir=folder, _result=result,
             _late={"no_data": [], "by_dir": {folder: {"sales": [],
                                                      "mp": ["collection-9.csv"]}}})
    run_batch.finalize_diagnosis([r, dict(row("B"), out_dir=None)])

    assert {m["form"]: m["cause"] for m in r["missing"]} == {
        "stock/general": "UI_NOT_FOUND",
        "mercadopago/release": "REQUEST_NOT_REGISTERED"}
    assert r["status"] == "partial"
    meta = run_downloads.read_run_meta(folder)
    assert [m["form"] for m in meta["missing"]] == [
        "stock/general", "mercadopago/release"]
    assert "downloaded_at" in meta                 # 没被冲掉
    text = io.open(os.path.join(folder, "REPORT.md"), encoding="utf-8").read()
    assert "REQUEST_NOT_REGISTERED" in text and "collection-9.csv" in text
    assert "mercadopago/collection" not in text.split("## 未下载")[1].split(
        "### 补收")[0]


def test_collect_only_updates_the_old_folder_in_place(
        monkeypatch, clock, registry, tmp_path):
    folder = str(tmp_path / "A" / "old")
    os.makedirs(folder)
    owed = dict(dd.classify({"row_state": "generating"},
                            {"form": "mercadopago/release"}), owed=True)
    result = {"store": "A", "out_dir": folder, "billing": {}, "sales": [],
              "stock": {}, "mp": [], "pending": [], "empty": [], "errors": [],
              "seconds": 1, "missing": [owed], "flags": {}, "cautions": []}
    run_downloads.write_run_meta(folder, result, clock.now)
    run_batch.run_report.write(result)
    # --collect-only：行里没有 out_dir 也没有 result；这张已在补收时取回
    r = dict(row("A"), out_dir=None,
             _late={"no_data": [], "by_dir": {folder: {"sales": [],
                                                      "mp": ["release-1.csv"]}}})
    run_batch.finalize_diagnosis([r])
    assert r["missing"] == [] and r["status"] == "ok"
    assert run_downloads.read_run_meta(folder)["missing"] == []
    text = io.open(os.path.join(folder, "REPORT.md"), encoding="utf-8").read()
    assert "全部到齐" in text and "`release-1.csv`" in text


def test_expired_entries_are_written_back_to_their_own_folder(
        monkeypatch, clock, registry, tmp_path):
    folder = str(tmp_path / "Z" / "old")
    os.makedirs(folder)
    io.open(os.path.join(folder, "REPORT.md"), "w", encoding="utf-8").write(
        u"# old\n\n---\n\nfooter\n")
    exp = dd.expired_record(dict(mp("release"), store="Z", out_dir=folder))
    rows = [row("A")]
    run_batch.finalize_diagnosis(rows, expired=[exp])
    assert rows[0]["missing"] == []          # 不记到这一趟的店铺头上
    assert run_downloads.read_run_meta(folder)["missing"][0]["cause"] == "EXPIRED"
    assert "EXPIRED" in io.open(os.path.join(folder, "REPORT.md"),
                                encoding="utf-8").read()


# ------------------------------------------------------------------ 紫鸟（fix h）

def bare_client():
    c = object.__new__(ziniao_client.ZiniaoClient)
    c._opened, c.sent, c.user_info, c.port = [], [], {}, 0

    def send(data, timeout=120):
        c.sent.append(data["action"])
        if data["action"] == "startBrowser":
            return {"statusCode": 0, "browserOauth": "oauth-1",
                    "debuggingPort": 1, "core_version": "120.0"}
        return {"statusCode": 0}
    c.send = send
    return c


def test_open_store_closes_the_browser_when_the_driver_cannot_be_resolved():
    c = bare_client()

    def boom(info, wait=None):
        raise ziniao_client.ZiniaoError("webdriver 仍未出现")
    c._resolve_driver = boom
    with pytest.raises(ziniao_client.ZiniaoError) as ei:
        with c.open_store({"browserOauth": "oauth-1", "browserName": "A"}):
            pytest.fail("不该进到 with 里面")
    assert c.sent == ["startBrowser", "stopBrowser"]      # 以前没有 stopBrowser
    assert ei.value.ziniao_step == "resolve_driver"
    assert dd.failure_signal(ei.value) == "store-open"


def test_open_store_closes_the_browser_when_attaching_fails(monkeypatch):
    c = bare_client()
    c._resolve_driver = lambda info, wait=None: "/nonexistent/webdriver"

    def no_chrome(*a, **k):
        raise OSError("cannot start webdriver")
    monkeypatch.setattr(ziniao_client.webdriver, "Chrome", no_chrome)
    with pytest.raises(OSError) as ei:
        with c.open_store({"browserOauth": "oauth-1", "browserName": "A"}):
            pass
    assert c.sent == ["startBrowser", "stopBrowser"]
    assert ei.value.ziniao_step == "attach_webdriver"


def test_ziniao_error_carries_step_and_response():
    c = bare_client()
    with pytest.raises(ziniao_client.ZiniaoError) as ei:
        c._require_ok({"statusCode": -10003, "err": "未登录"}, "startBrowser")
    assert ei.value.ziniao_step == "startBrowser"
    assert "未登录" in ei.value.ziniao_response
    recs = dd.store_failure_records(ei.value, ["ventas"])
    assert recs[0]["cause"] == "STORE_OPEN_FAILED" and "未登录" in recs[0]["explain"]


# ------------------------------------------------------------------ 整家店（fix f）

class StoreSession(Session):
    info = {"debuggingPort": 1, "core_version": "120"}

    def check_ip(self):
        return False


class StoreClient(Client):
    @contextlib.contextmanager
    def open_store(self, store, close_when_done=True):
        yield StoreSession()


def store_args(*argv):
    return run_downloads.add_route_args(argparse.ArgumentParser()).parse_args(
        list(argv))


def fake_routes(monkeypatch):
    monkeypatch.setattr(run_downloads.store_config, "require_allowed", lambda n: n)
    monkeypatch.setattr(meli_forms, "set_download_dir", lambda d, p: True)
    monkeypatch.setattr(meli_forms, "ensure_language", lambda d: ("en_US", True))

    def ventas(d, period_months=6, diag=None):
        dd.flag(diag, "range_clamped", {"ventas": {"requested_months": 6,
                                                   "actual_months": 2}})
        dd.give_up(diag, d, "ventas", "ui-disabled", step="按钮被禁用")
        return None

    def report(d, report="settlement", days=60, fmt="csv", diag=None, cover_from=None):
        if report == "release":
            raise KeyError("boom")
        e = mp(report)
        if report == "after_collection":
            e["adopted_existing_row"] = {"requested": "a .. b", "adopted": "c"}
        return e

    def billing_month(d, out_dir, y, m, extra=2, diag=None):
        # 目标月本身失败：页面被重定向到别的账期
        dd.give_up(diag, d, "billing/%04d-%02d" % (y, m), "wrong-period-page",
                   required=True, shown="Septiembre", want="Julio")
        return {}

    def stock(d, out_dir, months_back=2, diag=None):
        with open(os.path.join(out_dir, "stock_general_full_1.xlsx"), "w") as f:
            f.write("x")
        dd.delivered(diag, "stock/general", ["stock_general_full_1.xlsx"])
        dd.give_up(diag, d, "stock/costos/2", "not-offered", shot=False)
        return {"general": ["stock_general_full_1.xlsx"]}

    monkeypatch.setattr(meli_forms, "request_sales_excel", ventas)
    monkeypatch.setattr(mercadopago, "request_report", report)
    monkeypatch.setattr(meli_forms, "download_billing_for_months", billing_month)
    monkeypatch.setattr(meli_forms, "download_stock_reports", stock)


def test_run_store_records_why_each_form_is_missing(monkeypatch, clock, tmp_path):
    fake_routes(monkeypatch)
    script(monkeypatch, {"settlement": ["ready"], "collection": ["waiting"],
                         "withdraw": ["empty"], "after_collection": ["delayed"],
                         "account_statement": [
                             ("waiting", {"row_state": "accordion",
                                          "action": "no-format-row"})]},
           fetched={"settlement": ["settlement_v2-1.csv"]})
    out = str(tmp_path / "S" / "run")
    result = run_downloads.run_store(
        StoreClient(), "S", store_args("--month", "2020-07", "--billing-after", "0"),
        out_dir=out, defer_slow=True)

    got = {m["form"]: m["cause"] for m in result["missing"]}
    assert got == {
        "ventas": "UI_NOT_FOUND",
        "mercadopago/release": "INTERNAL_ERROR",
        "mercadopago/collection": "STILL_GENERATING",
        "mercadopago/withdraw": "NO_DATA",
        "mercadopago/after_collection": "PLATFORM_DELAYED",
        "mercadopago/account_statement": "REQUEST_NOT_REGISTERED",
        "billing/2020-07": "PAGE_BLOCKED",
        "stock/costos/1": "UNKNOWN",
        "stock/costos/2": "NOT_AVAILABLE",
        "stock/consolidado": "UNKNOWN",
        "stock/devoluciones": "UNKNOWN"}
    assert "mercadopago/settlement" in result["expected"]
    # 账单、库存缺了要算错误，店铺不能还是 ok（fix f）
    assert any(e.startswith("billing: 1 张未下载") for e in result["errors"])
    assert any(e.startswith("stock: 3 张未下载") for e in result["errors"])
    assert result["flags"] == {
        "ip_check_failed": True, "locale": "en_US",
        "range_clamped": {"ventas": {"requested_months": 6, "actual_months": 2}}}
    assert result["cautions"] == [{"form": "mercadopago/after_collection",
                                   "kind": "adopted_existing_row",
                                   "requested": "a .. b", "adopted": "c"}]
    # 推迟到补收的条目带着诊断字段
    assert sorted(e["report"] for e in result["deferred"]) == [
        "account_statement", "collection"]
    assert all(e["probes"] == 1 and e["evidence"]["screenshot"]
               for e in result["deferred"])

    meta = run_downloads.read_run_meta(out)
    assert meta["expected"] == result["expected"]
    assert {m["form"] for m in meta["missing"]} == set(got)
    assert meta["flags"]["ip_check_failed"] is True
    text = io.open(os.path.join(out, "REPORT.md"), encoding="utf-8").read()
    assert "## 未下载 —— 原因" in text and "`billing/2020-07`" in text
    assert "### 留意：按列表里已有的行匹配" in text and "a .. b" in text
    assert os.path.isdir(os.path.join(out, "_diag"))


def test_guard_captures_evidence_when_a_route_raises(tmp_path):
    diag = dd.Collector(str(tmp_path), "S")
    result = {"errors": []}
    d = Driver()

    def route(driver, out_dir, diag=None):
        raise RuntimeError("Reportes panel never rendered")
    got = run_downloads._guard(result, "billing", route, d, "x", diag=diag)
    assert got is run_downloads.FAILED and d.shots == 1
    recs = dd.diagnose_store(["billing/1", "billing/2"], diag)
    assert [r["cause"] for r in recs] == ["UI_NOT_FOUND", "UI_NOT_FOUND"]
    # 老调用方（download_history）不传 diag：行为不变
    assert run_downloads._guard(result, "x", lambda: 1) == 1
    assert run_downloads._guard(result, "x", lambda: 1 / 0) is run_downloads.FAILED


def test_store_has_full_only_when_told(monkeypatch):
    monkeypatch.setattr(run_downloads.store_config, "_stores",
                        lambda: {"no_full": ["B"]})
    a = store_args("--no-full-stores", "A")
    assert run_downloads.store_has_full("A", a) is False
    assert run_downloads.store_has_full("B", a) is False
    assert run_downloads.store_has_full("C", a) is None


# ------------------------------------------------------------------ 整个批次

def batch_world(monkeypatch, tmp_path, argv):
    logs = str(tmp_path / "logs")
    monkeypatch.setattr(run_batch, "LOG_DIR", logs)
    monkeypatch.setattr(run_batch.time, "sleep", lambda s: None)
    monkeypatch.setattr(run_batch.store_config, "batch_stores", lambda: ["A", "B"])
    monkeypatch.setattr(run_batch.store_config, "describe", lambda: "test")
    monkeypatch.setattr(run_batch.pending_store, "prune", lambda: [])
    monkeypatch.setattr(run_batch.transform, "available", lambda: False)

    class Z(Client):
        def start(self, restart=True):
            pass
    monkeypatch.setattr(run_batch, "ZiniaoClient", Z)
    monkeypatch.setattr("sys.argv", ["run_batch.py", "--no-transform",
                                     "--pace", "0"] + argv)
    return logs


def test_batch_main_prints_the_why_section_and_the_marker_line(
        monkeypatch, clock, registry, tmp_path, capsys):
    logs = batch_world(monkeypatch, tmp_path,
                       ["--only", "mercadopago", "--out", str(tmp_path / "dl"),
                        "--late-timeout", "0"])
    script(monkeypatch, {"release": [("waiting", {"row_state": "absent"})]})

    def fake_run_store(client, name, args, out_dir=None, close_when_done=True,
                       defer_slow=False):
        if name == "A":
            raise ziniao_client.ZiniaoClient._error(
                "startBrowser failed: {...}", "startBrowser", '{"statusCode": -1}')
        os.makedirs(out_dir)
        owed = mp("release", requested_at=clock.now - 3600)
        dd.note_probe(owed, {"row_state": "absent"}, now=clock.now - 1800)
        nothing = dd.classify({"signal": "not-offered"}, {"form": "stock/costos/2"})
        result = {"store": name, "out_dir": out_dir, "billing": {}, "sales": [],
                  "stock": {}, "mp": [], "pending": [], "empty": [], "errors": [],
                  "seconds": 1, "deferred": [owed], "flags": {},
                  "expected": ["mercadopago/release", "stock/costos/2"],
                  "cautions": [{"form": "mercadopago/after_collection",
                                "requested": "a .. b", "adopted": "c"}],
                  "missing": [dict(dd.record_from_entry(owed), owed=True), nothing]}
        run_downloads.write_run_meta(out_dir, result, clock.now)
        result["report"] = run_batch.run_report.write(result)
        return result
    monkeypatch.setattr(run_batch.run_downloads, "run_store", fake_run_store)
    monkeypatch.setattr(run_batch.run_downloads, "print_summary", lambda r: [])

    code = run_batch.main()
    out = capsys.readouterr().out
    lines = out.rstrip().splitlines()

    assert code == 1                          # A 失败，B 有欠的：不是 2
    assert lines[-1].startswith("[BATCH_SUMMARY_JSON] ")
    path = lines[-1].split(" ", 1)[1]
    assert os.path.isabs(path) and path.startswith(logs)
    assert out.count("[BATCH_SUMMARY_JSON]") == 1

    section = out.split("未下载的报表与原因")[1]
    assert "A | mercadopago/settlement" in section
    assert "[STORE_OPEN_FAILED]" in section and "startBrowser" in section
    # 店内只探测到一次不在（分不清）；补收又探测到一次，这时分得清了
    assert "[REQUEST_NOT_REGISTERED]" in section and "[UNKNOWN]" not in section
    assert "[NOT_AVAILABLE]" in section
    assert "按原因统计" in section and "申请 a .. b → 采用 c" in section

    doc = json.load(io.open(path, encoding="utf-8"))
    assert doc["exit_code"] == 1
    assert (doc["stores_total"], doc["stores_ok"], doc["stores_attention"],
            doc["stores_failed"]) == (2, 0, 1, 1)
    assert doc["missing_by_cause"] == {"STORE_OPEN_FAILED": 6,
                                       "REQUEST_NOT_REGISTERED": 1,
                                       "NOT_AVAILABLE": 1}
    a, b = doc["stores"]
    assert a["status"] == "failed" and len(a["missing"]) == 6
    assert b["status"] == "partial"
    assert {m["form"]: m["cause"] for m in b["missing"]} == {
        "stock/costos/2": "NOT_AVAILABLE",
        "mercadopago/release": "REQUEST_NOT_REGISTERED"}
    assert b["pending"][0].startswith("mercadopago/release (1/ago/2026 a 1/oct/2026)")
    # B 的目录里，报告和元数据都换成了最终结论
    meta = run_downloads.read_run_meta(b["out_dir"])
    assert {m["cause"] for m in meta["missing"]} == {"NOT_AVAILABLE",
                                                    "REQUEST_NOT_REGISTERED"}
    text = io.open(os.path.join(b["out_dir"], "REPORT.md"), encoding="utf-8").read()
    assert "REQUEST_NOT_REGISTERED" in text and "UNKNOWN" not in text
    # 登记簿里那一条带着诊断字段
    (e,) = pending_store.load()
    assert e["probes"] == 2 and e["cause"] == "REQUEST_NOT_REGISTERED"


def test_batch_main_exits_2_when_no_store_got_anywhere(
        monkeypatch, clock, registry, tmp_path, capsys):
    batch_world(monkeypatch, tmp_path, ["--only", "ventas"])

    def dead(*a, **k):
        raise RuntimeError("invalid session id")
    monkeypatch.setattr(run_batch.run_downloads, "run_store", dead)
    assert run_batch.main() == 2
    out = capsys.readouterr().out
    assert "没有一家下载成功" in out and "[STORE_OPEN_FAILED]" in out
    assert out.rstrip().splitlines()[-1].startswith("[BATCH_SUMMARY_JSON] ")
