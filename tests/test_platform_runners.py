# -*- coding: utf-8 -*-
"""core/platform_batch and core/platform_reports — the shared runners behind every
non-MercadoLibre platform's run_batch / run_reports — with fake platform modules,
a fake Ziniao client and a fake report_build. Nothing opens a browser."""
import glob
import io
import json
import os
import sys
import types

import pytest

from core import download_diag
from core import paths
from core import pending_store
from core import platform_batch as PB
from core import platform_reports as PR
from core import report_kit
from core import store_config

CODE, MARKET, COUNTRY = "FA", "FA-CL", "CL"     # a registered code; the modules are fakes
MONTH = "2026-09"
S1, S2 = "RUN_S1", "RUN_S2"


def rec(form, cause):
    explain, owner, retry, action = download_diag.CAUSES[cause]
    return {"form": form, "period": "", "cause": cause, "explain": explain, "owner": owner,
            "retry": retry, "action": action, "evidence": {}, "first_seen": "",
            "last_seen": "", "probes": 0, "in_report": True}


@pytest.fixture
def sandbox(tmp_path, monkeypatch):
    cfg = {"default": S1, "allowed": [S1, S2, "ML_1"],
           "catalog": {S1: {"platform": CODE, "country": COUNTRY},
                       S2: {"platform": CODE, "country": COUNTRY},
                       "ML_1": {"platform": "ML", "country": "MX"}},
           "batches": {MARKET: [S1, S2]}}
    monkeypatch.setattr(store_config, "_stores", lambda: cfg)
    monkeypatch.setattr(paths, "ROOT", str(tmp_path))
    return tmp_path


# ------------------------------------------------------------------ fake platform

def out_dir_of(root, name):
    return os.path.join(str(root), "downloads", name, "20261001_000000_m202609")


def write_meta(d, **meta):
    os.makedirs(d, exist_ok=True)
    with io.open(os.path.join(d, "run_meta.json"), "w", encoding="utf-8") as fh:
        json.dump(meta, fh)


def read_meta(d):
    try:
        with io.open(os.path.join(d, "run_meta.json"), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def update_meta(d, **fields):
    m = read_meta(d)
    m.update(fields)
    write_meta(d, **m)
    return m


def fake_platform(root, behaviour=None, collect=None, implemented=True):
    """forms, periods, run_downloads, site modules of a pretend platform."""
    behaviour = behaviour or {}
    collect = collect if collect is not None else {}

    def run_store(client, name, month, only=None, collect_timeout=240):
        how = behaviour.get(name, "ok")
        if how == "raise":
            raise RuntimeError("boom")
        if how == "not_allowed":
            raise store_config.StoreNotAllowed(name)
        d = out_dir_of(root, name)
        missing, pending = [], []
        if how == "attention":
            missing = [rec("stock", "UI_NOT_FOUND")]
        if how == "owed":
            missing = [rec("late", "STILL_GENERATING")]
            pending = [{"store": name, "report": "r_late", "form": "late", "range": month,
                        "out_dir": d}]
        write_meta(d, files={"orders": ["o.xlsx"]}, missing=missing)
        return {"store": name, "market": MARKET, "month": month, "out_dir": d,
                "files": {"orders": ["o.xlsx"]}, "missing": missing, "errors": [],
                "pending": pending, "expected": ["orders"], "seconds": 1}

    rd = types.SimpleNamespace(
        ROUTE_NAMES=("orders", "late"),
        expected=lambda month, only=None: ["orders", "late"],
        run_store=run_store,
        print_summary=lambda result: print("summary %s" % result["store"]),
        count_files=lambda r: sum(len(v) for v in (r.get("files") or {}).values()),
        read_run_meta=read_meta, update_run_meta=update_meta)
    if not implemented:
        def require():
            raise NotImplementedError("还没实现：测试")
        rd.require_implemented = require

    class Ctx(object):
        def __init__(self, driver, out_dir, _x, rng, facts, log=print):
            self.out_dir, self.rng, self.facts = out_dir, rng, facts

    def collect_owed(ctx, entry):
        how = collect.get(entry["store"], "files")
        if how == "raise":
            raise RuntimeError("page changed")
        return [] if how == "none" else ["late.xlsx"]

    forms = types.SimpleNamespace(Ctx=Ctx, collect_owed=collect_owed)
    periods = types.SimpleNamespace(previous_month=lambda today: MONTH,
                                    parse_month=lambda m: (int(m[:4]), int(m[5:7])))
    site = types.SimpleNamespace(get=lambda country: {"country": country})
    return forms, periods, rd, site


class Driver(object):
    def execute_cdp_cmd(self, *a):
        pass


class Session(object):
    driver = Driver()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


class FakeClient(object):
    platform_name = "Falabella-智利"
    broken = False
    unopenable = ()

    def __init__(self):
        if FakeClient.broken:
            raise RuntimeError("no ziniao")

    def is_up(self):
        return True

    def start(self, restart=False):
        pass

    def find_store(self, name):
        if name in FakeClient.unopenable:
            raise RuntimeError("gone")
        return {"name": name, "platform_name": FakeClient.platform_name}

    def open_store(self, store):
        return Session()


@pytest.fixture
def client(monkeypatch):
    import core.ziniao_client as zc
    FakeClient.broken, FakeClient.unopenable = False, ()
    monkeypatch.setattr(zc, "ZiniaoClient", FakeClient)
    return FakeClient


def runner(root, log_collected=False, **kw):
    forms, periods, rd, site = fake_platform(root, **kw)
    return PB.BatchRunner(CODE, MARKET, forms=forms, periods=periods, run_downloads=rd,
                          site=site, log_collected=log_collected)


def last_line(capsys):
    out = capsys.readouterr().out
    lines = [l for l in out.splitlines() if l.strip()]
    return out, lines[-1]


def summary_of(line):
    assert line.startswith("[BATCH_SUMMARY_JSON] ")
    path = line.split(" ", 1)[1]
    with io.open(path, encoding="utf-8") as fh:
        return path, json.load(fh)


# ------------------------------------------------------------------ batch: verdict

def test_batch_verdict_exit_codes(sandbox):
    r = runner(sandbox)
    ok = {"status": "ok", "missing": [rec("x", "NO_DATA")], "errors": []}
    att = {"status": "partial", "missing": [rec("x", "UI_NOT_FOUND")], "errors": []}
    bad = {"status": "failed", "missing": [], "errors": ["x"]}
    assert r.batch_verdict([])["exit_code"] == 0
    assert r.batch_verdict([ok])["exit_code"] == 0
    assert r.batch_verdict([ok, att])["exit_code"] == 1
    assert r.batch_verdict([ok, bad])["exit_code"] == 1
    assert r.batch_verdict([bad, bad])["exit_code"] == 2
    v = r.batch_verdict([ok, att, bad])
    assert (v["stores_ok"], v["stores_attention"], v["stores_failed"]) == (1, 1, 1)
    assert v["platform"] == CODE
    assert v["transform"]["note"] == "Falabella 没有清洗这一步"     # name from core/markets


def test_row_status(sandbox):
    r = runner(sandbox)
    base = {"store": S1, "market": MARKET, "out_dir": "d", "pending": [], "seconds": 1,
            "files": {"a": ["f"]}, "missing": [], "errors": []}
    assert r.row_of(base)["status"] == "ok"
    assert r.row_of(dict(base, missing=[rec("x", "UI_NOT_FOUND")]))["status"] == "partial"
    assert r.row_of(dict(base, missing=[rec("x", "NO_DATA")]))["status"] == "ok"
    assert r.row_of(dict(base, files={}, errors=["e"]))["status"] == "failed"


# ------------------------------------------------------------------ batch: stores

class Args(object):
    def __init__(self, stores=None, market=MARKET):
        self.stores, self.market = stores, market


def test_resolve_stores(sandbox, monkeypatch):
    r = runner(sandbox)
    assert r.resolve_stores(Args()) == [S1, S2]
    assert r.resolve_stores(Args([S2])) == [S2]
    with pytest.raises(store_config.StoreNotAllowed):
        r.resolve_stores(Args(["NOPE"]))
    with pytest.raises(store_config.StoreConfigError):
        r.resolve_stores(Args(["ML_1"]))
    monkeypatch.setattr(store_config, "stores_for_market", lambda market: [])
    with pytest.raises(store_config.StoreConfigError) as e:
        r.resolve_stores(Args())
    assert "没有要跑的 Falabella 店铺" in str(e.value)


def test_refusals_before_ziniao(sandbox, client, capsys):
    FakeClient.broken = True                           # would give 2 if it were reached
    r = runner(sandbox)
    assert r.main(["--stores", "NOPE"]) == 3
    assert r.main(["--stores", "ML_1"]) == 4
    assert r.main(["--market", "ML-MX"]) == 4
    assert r.main(["--market", "QQ-MX"]) == 4
    assert "不是 Falabella 的市场" in capsys.readouterr().out
    assert not os.path.exists(os.path.join(str(sandbox), "logs"))


def test_dry_run_and_not_implemented(sandbox, client, capsys):
    FakeClient.broken = True
    assert runner(sandbox).main(["--dry-run"]) == 0
    out = capsys.readouterr().out
    assert "%s -> orders, late" % S1 in out and "BATCH_SUMMARY_JSON" not in out
    assert runner(sandbox, implemented=False).main([]) == 4
    assert "[未实现] 还没实现" in capsys.readouterr().out


# ------------------------------------------------------------------ batch: a run

def test_run_writes_summary_last(sandbox, client, capsys):
    r = runner(sandbox, behaviour={S2: "attention"})
    assert r.main(["--month", MONTH]) == 1
    out, line = last_line(capsys)
    path, doc = summary_of(line)
    assert os.path.basename(path).startswith("batch_FA_")
    assert os.path.dirname(path) == os.path.join(str(sandbox), "logs")
    assert [s["store"] for s in doc["stores"]] == [S1, S2]
    assert all("_result" not in s for s in doc["stores"])
    assert doc["exit_code"] == 1 and "late" not in doc
    assert "# [1/2] %s" % S1 in out and "批次结论：2 家店，正常 1，要看 1，失败 0" in out


def test_one_store_failing_does_not_stop_the_batch(sandbox, client, capsys):
    assert runner(sandbox, behaviour={S1: "raise"}).main([]) == 1
    _out, line = last_line(capsys)
    _p, doc = summary_of(line)
    assert [s["status"] for s in doc["stores"]] == ["failed", "ok"]
    assert doc["stores"][0]["missing"][0]["cause"] == "STORE_OPEN_FAILED"
    assert runner(sandbox, behaviour={S1: "raise", S2: "raise"}).main([]) == 2


def test_store_refused_mid_batch_stops_with_3(sandbox, client, capsys):
    assert runner(sandbox, behaviour={S2: "not_allowed"}).main([]) == 3
    assert "BATCH_SUMMARY_JSON" not in capsys.readouterr().out


def test_ziniao_that_does_not_start(sandbox, client, capsys):
    FakeClient.broken = True
    assert runner(sandbox).main([]) == 2
    _out, line = last_line(capsys)
    _p, doc = summary_of(line)
    assert [s["status"] for s in doc["stores"]] == ["failed", "failed"]


# ------------------------------------------------------------------ batch: owed reports

def pending_entries(root):
    return pending_store.load(os.path.join(str(root), "logs", "pending_FA.json"))


def test_owed_collected_in_the_same_run(sandbox, client, capsys):
    r = runner(sandbox, behaviour={S1: "owed"}, log_collected=True)
    r.main([])
    out, line = last_line(capsys)
    _p, doc = summary_of(line)
    assert doc["late"] == {"done": 1, "still_owed": 0}
    assert pending_entries(sandbox) == []
    meta = read_meta(out_dir_of(sandbox, S1))
    assert meta["files"]["late"] == ["late.xlsx"]
    assert not [m for m in meta["missing"] if m["form"] == "late"]
    assert "补取欠账：1 条" in out and "  %s late -> ['late.xlsx']" % S1 in out


def test_log_collected_is_a_hook(sandbox, client, capsys):
    runner(sandbox, behaviour={S1: "owed"}, log_collected=False).main([])
    assert "late -> " not in capsys.readouterr().out


def test_still_owed_is_kept_and_counted(sandbox, client, capsys):
    r = runner(sandbox, behaviour={S1: "owed", S2: "owed"}, collect={S1: "none", S2: "raise"})
    r.main([])
    out, line = last_line(capsys)
    _p, doc = summary_of(line)
    assert doc["late"] == {"done": 0, "still_owed": 2}
    left = pending_entries(sandbox)
    assert sorted(e["store"] for e in left) == [S1, S2]
    assert all(e["probes"] == 1 and e["last_probe_at"] for e in left)
    assert "%s late 补取失败：page changed" % S2 in out


def test_store_unopenable_at_catch_up_keeps_its_owed(sandbox, client, capsys):
    FakeClient.unopenable = (S1,)
    runner(sandbox, behaviour={S1: "owed"}).main([])
    assert [e["store"] for e in pending_entries(sandbox)] == [S1]
    assert "打不开，欠账留到下次" in capsys.readouterr().out


def test_collect_only(sandbox, client, capsys):
    path = os.path.join(str(sandbox), "logs", "pending_FA.json")
    d = out_dir_of(sandbox, S1)
    write_meta(d, files={}, missing=[rec("late", "STILL_GENERATING")])
    entry = {"store": S1, "report": "r_late", "form": "late", "range": MONTH, "out_dir": d}
    other = dict(entry, range="2026-08")
    pending_store.add([entry, other], path=path)
    r = runner(sandbox, collect={S1: "none"})
    assert r.main(["--collect-only", "--month", MONTH]) == 1      # still owed
    _out, line = last_line(capsys)
    _p, doc = summary_of(line)
    assert doc["collect_only"] and doc["stores"] == [] and doc["late"]["still_owed"] == 1
    r = runner(sandbox)
    assert r.main(["--collect-only", "--month", MONTH]) == 0
    assert [e["range"] for e in pending_entries(sandbox)] == ["2026-08"]   # other month kept
    assert read_meta(d)["files"] == {"late": ["late.xlsx"]}
    assert r.main(["--collect-only", "--month", MONTH]) == 0               # nothing owed
    _out, line = last_line(capsys)
    assert summary_of(line)[1]["late"] == {}


def test_platform_check_before_opening_at_catch_up(sandbox, client, capsys):
    FakeClient.platform_name = "Amazon-墨西哥"            # Ziniao disagrees with config
    try:
        runner(sandbox, behaviour={S1: "owed"}).main([])
    finally:
        FakeClient.platform_name = "Falabella-智利"
    assert [e["store"] for e in pending_entries(sandbox)] == [S1]


# ------------------------------------------------------------------ reports

def month_dir(root, store, stamp, month="202609"):
    d = os.path.join(str(root), "downloads", store, "%s_m%s" % (stamp, month))
    os.makedirs(d)
    return d


def touch(d, name):
    with open(os.path.join(d, name), "w") as fh:
        fh.write("x")


def report_runner(markers=None, **kw):
    periods = types.SimpleNamespace(previous_month=lambda today: MONTH,
                                    parse_month=lambda m: (int(m[:4]), int(m[5:7])))
    return PR.ReportRunner(CODE, "fake_platform_pkg", MARKET, periods,
                           folder_markers=markers, **kw)


def test_folder_markers_tried_in_order(sandbox):
    old = month_dir(sandbox, S1, "20261001_000000")
    touch(old, "Main_1.csv")
    touch(old, "run_meta.json")
    new = month_dir(sandbox, S1, "20261005_000000")
    touch(new, "run_meta.json")
    month_dir(sandbox, S1, "20261101_000000", month="202610")
    r = report_runner(lambda month: ("Main_*.csv", "run_meta.json"))
    assert r.latest_folder(S1, MONTH) == old                   # the one with the file
    assert report_runner().latest_folder(S1, MONTH) == new     # default: newest run_meta
    assert r.latest_folder(S1, "2026-08") is None
    os.remove(os.path.join(old, "Main_1.csv"))
    assert r.latest_folder(S1, MONTH) == new                   # falls back
    base = os.path.join(str(sandbox), "downloads")
    assert r.latest_folder(S1, MONTH, base) == new


def test_resolve_reason_is_a_hook(sandbox, capsys):
    usable, unusable = report_runner().resolve([S1], MONTH)
    assert usable == [] and unusable == [(S1, "没有 2026-09 的下载目录")]
    r = report_runner(no_folder_reason="没有 %s 的下载目录（或目录里没有交易明细）")
    assert r.resolve([S1], MONTH)[1][0][1].endswith("（或目录里没有交易明细）")
    assert "—— 没有可用的下载目录" in capsys.readouterr().out


def test_report_exit_codes():
    ok, look, bad = {"outcome": "ok"}, {"outcome": "checks_failed"}, {"outcome": "crashed"}
    u, nu = [("a", "d")], [("b", "why")]
    assert PR.exit_code([ok], [], []) == 2
    assert PR.exit_code([ok], u, []) == 0
    assert PR.exit_code([look], u, []) == 1
    assert PR.exit_code([{"outcome": "formula_errors"}], u, []) == 1
    assert PR.exit_code([ok, bad], u, []) == 2
    assert PR.exit_code([{"outcome": "strict_aborted"}], u, []) == 2
    assert PR.exit_code([ok], u, nu, per_store=True) == 2
    assert PR.exit_code([ok], u, nu, per_store=False) == 1


@pytest.mark.parametrize("exc,code,text", [
    (store_config.StoreNotAllowed("x"), 3, "店铺不在白名单"),
    (store_config.StoreConfigError("x"), 4, "配置错误"),
])
def test_run_as_script_maps_exceptions(exc, code, text, capsys):
    def main():
        raise exc
    with pytest.raises(SystemExit) as e:
        PR.run_as_script(main)
    assert e.value.code == code and text in capsys.readouterr().out
    with pytest.raises(SystemExit) as e:
        PR.run_as_script(lambda: 1)
    assert e.value.code == 1


@pytest.fixture
def fake_build(monkeypatch):
    calls = []
    mod = types.ModuleType("fake_platform_pkg.report_build")

    def build(pairs, month, out, strict=False):
        calls.append((list(pairs), out))
        outcome = "checks_failed" if any(n == S2 for n, _ in pairs) and len(pairs) == 1 else "ok"
        return {"outcome": outcome, "out": out,
                "checks": [{"ok": outcome == "ok", "skipped": False, "severity": "warn",
                            "name": "检查", "detail": "差一点"}]}
    mod.build = build
    monkeypatch.setitem(sys.modules, "fake_platform_pkg", types.ModuleType("fake_platform_pkg"))
    monkeypatch.setitem(sys.modules, "fake_platform_pkg.report_build", mod)
    monkeypatch.setattr(report_kit, "recalc", lambda out: (False, "跳过重算"))
    return calls


def test_reports_main(sandbox, fake_build, capsys):
    touch(month_dir(sandbox, S1, "20261001_000000"), "run_meta.json")
    touch(month_dir(sandbox, S2, "20261001_000000"), "run_meta.json")
    r = report_runner()
    assert r.main(["--month", MONTH]) == 1                   # S2's checks failed
    out = capsys.readouterr().out
    outs = [os.path.basename(o) for _p, o in fake_build]
    assert outs == ["%s_2026-09.xlsx" % S1, "%s_2026-09.xlsx" % S2, "CL_FA_全店汇总_2026-09.xlsx"]
    assert os.path.dirname(fake_build[0][1]) == os.path.join(
        str(sandbox), "downloads", "data", "reports", "financial")
    assert "单店报表：%s" % S1 in out and "汇总报表：2 家店铺" in out and "全店汇总" in out
    assert "[warn] 检查：差一点" in out and "[提示] 跳过重算" in out
    del fake_build[:]
    assert r.main(["--month", MONTH, "--stores", S1, "--no-combined", "--out-dir",
                   str(sandbox / "o")]) == 0
    assert [os.path.basename(o) for _p, o in fake_build] == ["%s_2026-09.xlsx" % S1]
    assert r.main(["--month", MONTH, "--dry-run"]) == 0
    assert r.main(["--month", "2026-07"]) == 2


def test_reports_crash_and_refusals(sandbox, fake_build, monkeypatch, capsys):
    touch(month_dir(sandbox, S1, "20261001_000000"), "run_meta.json")

    def boom(*a, **k):
        raise RuntimeError("exploded")
    monkeypatch.setattr(sys.modules["fake_platform_pkg.report_build"], "build", boom)
    r = report_runner()
    assert r.main(["--month", MONTH, "--stores", S1]) == 2
    assert "生成出错" in capsys.readouterr().out
    with pytest.raises(store_config.StoreNotAllowed):
        r.main(["--stores", "NOPE"])
    with pytest.raises(store_config.StoreConfigError):
        r.main(["--stores", "ML_1"])
    with pytest.raises(store_config.StoreConfigError):
        r.main(["--market", "ML-MX"])
    assert not glob.glob(os.path.join(str(sandbox), "downloads", "data", "*"))


# ------------------------------------------------------------------ the thin modules

@pytest.mark.parametrize("pkg", ["platforms.falabella", "platforms.amazon"])
def test_platform_modules_are_thin(pkg):
    import importlib
    for name, cls in (("run_batch", PB.BatchRunner), ("run_reports", PR.ReportRunner)):
        m = importlib.import_module("%s.%s" % (pkg, name))
        assert isinstance(m.RUNNER, cls) and m.main == m.RUNNER.main
        assert m.RUNNER.code == m.PLATFORM
    rr = importlib.import_module(pkg + ".run_reports").RUNNER
    assert rr.package == pkg
