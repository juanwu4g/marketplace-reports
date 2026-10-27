# -*- coding: utf-8 -*-
"""美客多下载层的健壮性：店铺中途出错、登录页、语言关卡、点击判定、迟到的
下载、批次进度、紫鸟客户端退出、跨月欠账。浏览器、紫鸟、时钟全是假的。"""
import contextlib
import io
import json
import os
import runpy

import pytest

from core import download_diag as dd
from core import pending_store
from core import ziniao_client
from platforms.mercadolibre import meli_forms
from platforms.mercadolibre import mercadopago
from platforms.mercadolibre import run_batch
from platforms.mercadolibre import run_downloads

from tests.test_batch import (Client, Clock, Driver, Session, batch_world, mp,
                              registry, script, store_args, late_args)  # noqa: F401


@pytest.fixture
def clock(monkeypatch):
    c = Clock()
    monkeypatch.setattr(run_downloads.time, "time", c.time)
    monkeypatch.setattr(run_downloads.time, "sleep", c.sleep)
    return c


class LoginDriver(Driver):
    current_url = "https://www.mercadolibre.com.mx/jms/mlm/lgz/login?go=x"
    title = "Ingresa a tu cuenta"


class LauncherDies(Driver):
    """启动页打不开（代理慢 / 窗口关了）。"""

    def get(self, url):
        if url == "about:launcher":
            raise RuntimeError("timeout: Timed out receiving message from renderer")
        self.visited.append(url)


def store_client(driver):
    class S(Session):
        info = {"debuggingPort": 1, "core_version": "120"}

        def __init__(self):
            self.driver = driver

        def check_ip(self):
            return True

    class C(Client):
        @contextlib.contextmanager
        def open_store(self, store, close_when_done=True):
            self.opened.append(store["browserName"])
            yield S()
    return C()


def quiet_routes(monkeypatch, calls, language=lambda d: ("es_MX", False)):
    """所有路线的替身：记下谁被调用了。MercadoPago 申请成功、等着补收。"""
    monkeypatch.setattr(run_downloads.store_config, "require_allowed", lambda n: n)
    monkeypatch.setattr(meli_forms, "set_download_dir", lambda d, p: True)
    monkeypatch.setattr(meli_forms, "ensure_language", language)

    def ventas(d, period_months=6, diag=None):
        calls.append("ventas")
        return None

    def report(d, report="settlement", days=60, fmt="csv", diag=None, cover_from=None):
        calls.append("mp/" + report)
        return mp(report)

    def billing(d, out_dir, months=2, skip_current=False, diag=None):
        calls.append("billing")
        return {}

    def stock(d, out_dir, months_back=2, diag=None):
        calls.append("stock")
        return {}
    monkeypatch.setattr(meli_forms, "request_sales_excel", ventas)
    monkeypatch.setattr(mercadopago, "request_report", report)
    monkeypatch.setattr(meli_forms, "download_billing_reports", billing)
    monkeypatch.setattr(meli_forms, "download_stock_reports", stock)


# ------------------------------------------------------------------ F1

def test_launcher_page_failure_does_not_lose_the_store(monkeypatch, clock, tmp_path):
    calls = []
    quiet_routes(monkeypatch, calls)
    script(monkeypatch, {k: ["waiting"] for k in mercadopago.REPORTS})
    out = str(tmp_path / "S" / "run")
    result = run_downloads.run_store(
        store_client(LauncherDies()), "S",
        store_args("--only", "mercadopago", "--skip-ip-check"),
        out_dir=out, defer_slow=True)
    # 进店和收尾两次回启动页都失败了，店铺照常跑完、照常收尾
    assert len(result["deferred"]) == len(mercadopago.REPORTS)
    assert not any(e.startswith("store:") for e in result["errors"])
    assert os.path.isfile(os.path.join(out, "run_meta.json"))
    assert os.path.isfile(os.path.join(out, "REPORT.md"))


def test_error_after_the_requests_keeps_the_deferred_reports(monkeypatch, clock, tmp_path):
    calls = []
    quiet_routes(monkeypatch, calls)

    def probe(d, p, settle=12):
        if p["report"] == "settlement":
            p["_probe"] = {"row_state": "ready"}
            return mercadopago.READY
        raise mercadopago.NoReportAccess("No tienes permisos")   # PASSTHROUGH，_guard 不接
    monkeypatch.setattr(mercadopago, "probe", probe)

    def fetch(d, p, out_dir):
        with open(os.path.join(out_dir, "settlement-1.csv"), "w") as fh:
            fh.write("x")
        p["_fetch"] = "clicked"
        return ["settlement-1.csv"]
    monkeypatch.setattr(mercadopago, "fetch", fetch)

    out = str(tmp_path / "S" / "run")
    result = run_downloads.run_store(
        store_client(Driver()), "S", store_args("--only", "mercadopago"),
        out_dir=out, defer_slow=True)       # 以前：异常冲出 run_store，申请全丢
    assert any(e.startswith("store: NoReportAccess") for e in result["errors"])
    # 已取到的不再欠着，其余 5 张进了 deferred —— 批次据此写进登记簿
    assert result["mp"] == ["settlement-1.csv"]
    assert sorted(p["report"] for p in result["deferred"]) == sorted(
        k for k in mercadopago.REPORTS if k != "settlement")
    meta = run_downloads.read_run_meta(out)
    assert {m["form"] for m in meta["missing"]} == {
        "mercadopago/%s" % k for k in mercadopago.REPORTS if k != "settlement"}
    assert os.path.isfile(os.path.join(out, "REPORT.md"))


# ------------------------------------------------------------------ F3

def test_login_page_stops_every_route_as_page_blocked(monkeypatch, clock, tmp_path):
    calls = []
    quiet_routes(monkeypatch, calls, language=lambda d: (None, False))
    d = LoginDriver()
    out = str(tmp_path / "S" / "run")
    result = run_downloads.run_store(store_client(d), "S", store_args(),
                                     out_dir=out, defer_slow=True)
    assert calls == [], "被挡在登录页时一条路线都不该跑"
    assert result["blocked"] and "lgz/login" in result["blocked"]
    assert any(e.startswith("blocked:") for e in result["errors"])
    assert result["missing"] and {m["cause"] for m in result["missing"]} == {
        "PAGE_BLOCKED"}
    assert {m["form"] for m in result["missing"]} == set(result["expected"])
    assert d.shots == 1, "整店一张截图，不是每条路线一张"
    assert result["deferred"] == []


def test_late_pass_does_not_reopen_a_store_blocked_in_this_run(
        monkeypatch, clock, registry, tmp_path, capsys):
    old = str(tmp_path / "A" / "20261001_000000")
    os.makedirs(old)
    pending_store.add([dict(mp("release"), store="A", out_dir=old)])
    c = Client()
    row = {"store": "A", "status": "partial", "files": 0, "errors": [],
           "pending": [], "empty": [], "out_dir": str(tmp_path / "A" / "now"),
           "blocked": "地址是 https://…/login"}
    run_batch.collect_pass(c, ["A"], late_args(tmp_path), [row])
    assert c.opened == []
    assert len(pending_store.load()) == 1, "旧欠账原样留给下一趟"
    assert "跳过 A" in capsys.readouterr().out


# ------------------------------------------------------------------ F12

def test_language_that_cannot_be_restored_stops_the_routes(monkeypatch, clock, tmp_path):
    calls = []

    def wrong(d):
        raise RuntimeError("界面语言切换后仍是 en_US（期望 es_MX）。")
    quiet_routes(monkeypatch, calls, language=wrong)
    out = str(tmp_path / "S" / "run")
    result = run_downloads.run_store(store_client(Driver()), "S", store_args(),
                                     out_dir=out, defer_slow=True)
    assert calls == [], "语言不对还跑，下到的就是错语言的导出"
    assert {m["cause"] for m in result["missing"]} == {"LANGUAGE_WRONG"}
    assert {m["form"] for m in result["missing"]} == set(result["expected"])
    assert any(e.startswith("language:") for e in result["errors"])


# ------------------------------------------------------------------ F16

class AccordionDriver(Driver):
    def __init__(self, answer):
        Driver.__init__(self)
        self.answer = answer

    def execute_script(self, js, *a):
        return self.answer


@pytest.mark.parametrize("answer", ["not-found", "no-panel", "no-link", None])
def test_accordion_click_that_found_nothing_is_not_a_click(monkeypatch, tmp_path, answer):
    waited = []
    monkeypatch.setattr(meli_forms, "wait_for_downloads",
                        lambda *a, **k: waited.append(1) or [])
    p = mp("account_statement", rng="Agosto de 2026", flavor="accordion")
    assert mercadopago.fetch(AccordionDriver(answer), p, str(tmp_path)) == []
    assert waited == [], "没点到就不该白等 150 秒"
    assert p["_fetch"].startswith("no-click-target")
    if answer:
        assert answer in p["_fetch"]


def test_accordion_click_that_clicked_waits_for_the_file(monkeypatch, tmp_path):
    monkeypatch.setattr(meli_forms, "wait_for_downloads", lambda *a, **k: ["a.csv"])
    p = mp("account_statement", rng="Agosto de 2026", flavor="accordion")
    got = mercadopago.fetch(AccordionDriver("clicked:Descargar"), p, str(tmp_path))
    assert got == ["a.csv"] and p["_fetch"] == "clicked"


def test_unclicked_ready_report_is_ui_not_found_not_download_not_landed(
        monkeypatch, clock, tmp_path):
    monkeypatch.setattr(meli_forms, "wait_for_downloads", lambda *a, **k: [])

    def probe(d, p, settle=12):
        p["_probe"] = {"row_state": "ready"}
        return mercadopago.READY
    monkeypatch.setattr(mercadopago, "probe", probe)
    p = mp("account_statement", rng="Agosto de 2026", flavor="accordion")
    result = {"store": "S", "sales": [], "mp": [], "pending": [], "empty": [],
              "errors": []}
    left = run_downloads.collect_pending(AccordionDriver("not-found"), [p],
                                         str(tmp_path), result, budget=0)
    assert left == [p]
    assert dd.record_from_entry(p)["cause"] == "UI_NOT_FOUND"


# ------------------------------------------------------------------ F2

def test_late_rename_of_an_earlier_partial_is_not_taken_as_new(monkeypatch, tmp_path):
    folder = str(tmp_path)
    monkeypatch.setattr(meli_forms, "LATE_PARTIALS", {})
    open(os.path.join(folder, "export.xlsx.crdownload"), "w").close()
    before = meli_forms.snapshot_dir(folder)
    steps = iter([
        # 上一条路线的文件写完改名；这一次点击的文件随后到
        lambda: os.rename(os.path.join(folder, "export.xlsx.crdownload"),
                          os.path.join(folder, "export.xlsx")),
        lambda: open(os.path.join(folder, "mine.xlsx"), "w").close(),
    ])
    clock = Clock()

    def sleep(s):
        clock.sleep(s)
        next(steps, lambda: None)()
    monkeypatch.setattr(meli_forms.time, "time", clock.time)
    monkeypatch.setattr(meli_forms.time, "sleep", sleep)
    got = meli_forms.wait_for_downloads(folder, before, min_files=1, timeout=60,
                                        settle=4)
    assert got == ["mine.xlsx"]


def test_partial_left_by_a_timed_out_wait_is_excluded_later(monkeypatch, tmp_path):
    folder = str(tmp_path)
    monkeypatch.setattr(meli_forms, "LATE_PARTIALS", {})
    clock = Clock()
    monkeypatch.setattr(meli_forms.time, "time", clock.time)
    monkeypatch.setattr(meli_forms.time, "sleep", clock.sleep)
    before = meli_forms.snapshot_dir(folder)
    open(os.path.join(folder, "late.csv.crdownload"), "w").close()
    assert meli_forms.wait_for_downloads(folder, before, timeout=10) == []
    assert meli_forms.LATE_PARTIALS[os.path.abspath(folder)] == {"late.csv"}

    # 下一次等待用的是一份更早的快照（重试时沿用同一个 before）
    os.rename(os.path.join(folder, "late.csv.crdownload"),
              os.path.join(folder, "late.csv"))
    open(os.path.join(folder, "next.csv"), "w").close()
    got = meli_forms.wait_for_downloads(folder, before, timeout=30, settle=4)
    assert got == ["next.csv"]


# ------------------------------------------------------------------ F11 / A / B

def fake_store_result(name, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    return {"store": name, "out_dir": out_dir, "billing": {}, "sales": [],
            "stock": {}, "mp": [], "pending": [], "empty": [], "errors": [],
            "seconds": 1, "deferred": [], "flags": {}, "expected": [],
            "cautions": [], "missing": []}


def test_batch_summary_is_written_after_every_store(
        monkeypatch, clock, registry, tmp_path):
    logs = batch_world(monkeypatch, tmp_path, ["--only", "ventas",
                                               "--out", str(tmp_path / "dl")])
    seen = []

    def run_store(client, name, args, out_dir=None, close_when_done=True,
                  defer_slow=False):
        if name == "B":
            (path,) = [os.path.join(logs, f) for f in os.listdir(logs)
                       if f.startswith("batch_")]
            seen.append(json.load(io.open(path, encoding="utf-8")))
        return fake_store_result(name, out_dir)
    monkeypatch.setattr(run_batch.run_downloads, "run_store", run_store)
    monkeypatch.setattr(run_batch.run_downloads, "print_summary", lambda r: [])
    assert run_batch.main() == 0
    (doc,) = seen
    assert [s["store"] for s in doc["stores"]] == ["A"]
    assert doc["in_progress"] is True and doc["stores_planned"] == ["A", "B"]
    final = [f for f in os.listdir(logs) if f.startswith("batch_")]
    doc = json.load(io.open(os.path.join(logs, final[0]), encoding="utf-8"))
    assert "in_progress" not in doc and len(doc["stores"]) == 2


def test_client_that_cannot_start_is_a_hard_failure_with_a_summary(
        monkeypatch, clock, registry, tmp_path, capsys):
    logs = batch_world(monkeypatch, tmp_path, ["--only", "ventas"])

    def broken(*a, **k):
        raise KeyError("user_info")
    monkeypatch.setattr(run_batch, "ZiniaoClient", broken)
    assert run_batch.main() == 2
    out = capsys.readouterr().out
    assert "紫鸟客户端起不来" in out
    path = out.rstrip().splitlines()[-1].split(" ", 1)[1]
    doc = json.load(io.open(path, encoding="utf-8"))
    assert doc["exit_code"] == 2 and path.startswith(logs)
    assert [s["status"] for s in doc["stores"]] == ["failed", "failed"]


def test_unexpected_exception_at_top_level_exits_2(monkeypatch):
    def boom():
        raise ValueError("unexpected")
    monkeypatch.setattr(run_batch.store_config, "batch_stores", boom)
    monkeypatch.setattr("sys.argv", ["run_batch.py"])
    with pytest.raises(SystemExit) as ei:
        runpy.run_module("platforms.mercadolibre.run_batch", run_name="__main__")
    assert ei.value.code == 2


def gone():
    return ziniao_client.ZiniaoClient._error(
        "getBrowserList: no response from client", "getBrowserList", None)


def test_client_death_restarts_once_and_retries_the_store(
        monkeypatch, clock, registry, tmp_path):
    batch_world(monkeypatch, tmp_path, ["--only", "ventas", "--out", str(tmp_path / "dl")])
    starts, tries = [], []

    class Z(Client):
        def start(self, restart=True):
            starts.append(restart)
    monkeypatch.setattr(run_batch, "ZiniaoClient", Z)

    def run_store(client, name, args, out_dir=None, close_when_done=True,
                  defer_slow=False):
        tries.append(name)
        if name == "A" and tries.count("A") == 1:
            raise gone()
        return fake_store_result(name, out_dir)
    monkeypatch.setattr(run_batch.run_downloads, "run_store", run_store)
    monkeypatch.setattr(run_batch.run_downloads, "print_summary", lambda r: [])
    assert run_batch.main() == 0
    assert tries == ["A", "A", "B"]
    assert len(starts) == 2 and starts[1] is True


def test_client_that_will_not_restart_fails_the_rest_and_stops(
        monkeypatch, clock, registry, tmp_path, capsys):
    logs = batch_world(monkeypatch, tmp_path, ["--only", "ventas"])
    tries = []

    class Z(Client):
        n = 0

        def start(self, restart=True):
            Z.n += 1
            if Z.n > 1:
                raise ziniao_client.ZiniaoError("client did not open port 16851 in time")
    monkeypatch.setattr(run_batch, "ZiniaoClient", Z)

    def run_store(client, name, args, **k):
        tries.append(name)
        raise gone()
    monkeypatch.setattr(run_batch.run_downloads, "run_store", run_store)
    assert run_batch.main() == 2
    assert tries == ["A"], "重启失败后不再挨家去撞"
    out = capsys.readouterr().out
    assert "重启失败" in out
    path = out.rstrip().splitlines()[-1].split(" ", 1)[1]
    doc = json.load(io.open(path, encoding="utf-8"))
    assert [s["store"] for s in doc["stores"]] == ["A", "B"]
    assert all("重启失败" in s["errors"][0] for s in doc["stores"])
    assert doc["missing_by_cause"].keys() == {"STORE_OPEN_FAILED"}


# ------------------------------------------------------------------ C

def _meta(folder, expected, missing=()):
    os.makedirs(folder)
    with io.open(os.path.join(folder, "run_meta.json"), "w", encoding="utf-8") as fh:
        json.dump({"expected": expected, "missing": list(missing)}, fh)


def test_owed_report_of_another_month_is_not_superseded(tmp_path, monkeypatch):
    store = tmp_path / "S"
    aug, sep = str(store / "20261001_000000_m202608"), str(store / "20261002_000000_m202609")
    exp = ["mercadopago/release"]
    _meta(aug, exp, [{"form": "mercadopago/release", "cause": "STILL_GENERATING"}])
    _meta(sep, exp)
    removed = []
    monkeypatch.setattr(pending_store, "remove", lambda es: removed.extend(es))
    entry = {"kind": "mercadopago", "report": "release", "store": "S", "out_dir": aug}
    assert run_batch.drop_superseded([entry]) == [entry]
    assert removed == []

    row = {"store": "S", "out_dir": sep, "expected": exp}
    carried = [{"form": "mercadopago/release", "carried_from": os.path.basename(aug),
                "cause": "STILL_GENERATING"}]
    run_batch.settle_carried(row, [], carried, [entry], sep, exp)
    assert removed == [] and carried[0]["superseded"] is False


def test_newest_run_sees_every_month_folder_in_the_new_layout(tmp_path):
    exp = ["ventas"]
    store = tmp_path / "raw" / "ML" / "MX" / "S"
    regular_old = str(store / "其他" / "20261001_000000")
    regular_new = str(store / "其他" / "20261020_000000")
    month_run = str(store / "2026-09" / "20261010_000000_m202609")
    for f in (regular_old, regular_new, month_run):
        _meta(f, exp)
    # 只按 m 后缀区分月份：两种布局的回答一致
    assert run_batch.newest_run([month_run]) == (regular_new, exp)
    assert run_batch.newest_run([regular_old], month=None) == (regular_new, exp)
    assert run_batch.newest_run([regular_old], month="2026-09") == (month_run, exp)
    legacy = tmp_path / "downloads" / "S"
    for n in ("20261001_000000", "20261020_000000", "20261010_000000_m202609"):
        _meta(str(legacy / n), exp)
    assert run_batch.newest_run([str(legacy / "20261001_000000")], month="2026-09") == (
        str(legacy / "20261010_000000_m202609"), exp)
