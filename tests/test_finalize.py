# -*- coding: utf-8 -*-
"""core/finalize: final versions, the month folder and clean-up
(docs/RESULTS_AND_PERIODS.md §4), on synthetic results files and folders."""
import io
import json
import os
import shutil

import pytest

from core import finalize as FZ
from core import layout
from core import paths
from core import pending_store
from core import results as R
from core import store_config

MONTH = "2026-09"
MARKET = "ML-MX"
BASES = ("bill", "order")


@pytest.fixture
def v2(tmp_path, monkeypatch):
    stores = {"S1": ("ML", "MX"), "S2": ("ML", "MX"), "S3": ("ML", "MX")}
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({
        "paths": {"data_root": str(tmp_path / "data")},
        "stores": {"allowed": sorted(stores), "catalog": dict(
            (s, {"platform": p, "country": c}) for s, (p, c) in stores.items())}}), encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    return tmp_path


def name(store, month, basis):
    return "%s_%s%s.xlsx" % (store, month, "_账单基准" if basis == "bill" else "")


def cname(month, basis):
    return "MX_ML_全店汇总_%s%s.xlsx" % (month, "_账单基准" if basis == "bill" else "")


def unit(out_dir=None):
    return FZ.Unit("ML", MARKET, MONTH, out_dir or layout.monthly_dir(MONTH, MARKET), BASES, name, cname)


def make(store, run, built, level, basis="bill", workbook=True, month=MONTH):
    """A version: its raw run folder, its workbook in _versions/<version>/, its results file."""
    v = R.version_id(run, built)
    raw = layout.run_folder(store, run)
    os.makedirs(raw, exist_ok=True)
    wb = None
    if workbook:
        wb = os.path.join(FZ.version_dir(layout.monthly_dir(month, MARKET), v), name(store, month, basis))
        os.makedirs(os.path.dirname(wb), exist_ok=True)
        with io.open(wb, "w", encoding="utf-8") as fh:
            fh.write(v)
    doc = R.document("ML", MARKET, "MX", "MXN", 2, store, month, basis, v,
                     "%s-%s-%sT%s:%s:%s" % (built[:4], built[4:6], built[6:8], built[9:11], built[11:13], built[13:]),
                     [raw], wb, {"level": level, "reasons": [] if level == R.COMPLETE else ["原因 %s" % level]},
                     figures={"pl": [{"key": "x", "label": "X", "original": "", "kind": "d", "value": 1.0}]})
    R.write(doc)
    return v, raw, wb


def read(path):
    with io.open(path, encoding="utf-8") as fh:
        return fh.read()


class Builder(object):
    def __init__(self):
        self.calls = []

    def __call__(self, basis, pairs, dst):
        self.calls.append((basis, [(s, d["version"]) for s, d in pairs], os.path.basename(dst)))
        with io.open(dst, "w", encoding="utf-8") as fh:
            fh.write("summary")
        return {"outcome": "ok", "out": dst, "checks": []}


R1, R2, R3 = "20261003_010000_m202609", "20261012_010000_m202609", "20261026_010000_m202609"


def test_finals_copies_drafts_summaries_and_notes(v2):
    md = layout.monthly_dir(MONTH, MARKET)
    v1b, _r, _w = make("S1", R1, "20261003_050000", R.COMPLETE, "bill")
    v2b, _r, _w = make("S1", R2, "20261012_050000", R.GAPS, "bill")            # newer, not countable
    v1o, _r, _w = make("S1", R1, "20261003_050000", R.GAPS, "order")
    v2o, _r, _w = make("S1", R2, "20261012_050000", R.PROVISIONAL, "order")    # no final for order
    make("S2", R2, "20261012_050000", R.COMPLETE_WARN, "bill")
    make("S2", R2, "20261012_050000", R.COMPLETE, "order")
    os.makedirs(md, exist_ok=True)
    stale = os.path.join(md, cname(MONTH, "bill").replace(".xlsx", "_未定稿.xlsx"))
    open(stale, "w").close()
    b = Builder()
    out = FZ.run(unit(), ["S1", "S2"], ["S1", "S2", "S3"], b)
    assert out["finals"][("S1", "bill")]["version"] == v1b and out["finals"][("S1", "order")] is None
    # final copied under the usual name; no final -> newest version as 未定稿
    assert read(os.path.join(md, name("S1", MONTH, "bill"))) == v1b
    assert read(os.path.join(md, "S1_2026-09_未定稿.xlsx")) == v2o
    assert not os.path.exists(os.path.join(md, name("S1", MONTH, "order")))
    # summaries from the chosen versions; 未定稿 when a store is not final; S3 (no version) left out
    assert b.calls == [("bill", [("S1", v1b), ("S2", R.version_id(R2, "20261012_050000"))], cname(MONTH, "bill")),
                       ("order", [("S1", v2o), ("S2", R.version_id(R2, "20261012_050000"))],
                        "MX_ML_全店汇总_2026-09_未定稿.xlsx")]
    assert not os.path.exists(stale)                          # the old 未定稿 bill summary is gone
    notes = read(os.path.join(md, FZ.NOTES))
    assert "S1 · 账单基准：最终版本 %s（完整）" % v1b in notes
    assert "%s 不是最终版本：有缺口：原因 gaps" % v2b in notes
    assert "S1 · 订单基准：还没有最终版本（月份文件夹里是 %s 的未定稿）" % v2o in notes
    assert "%s 不是最终版本：暂定：原因 provisional" % v2o in notes          # every version says why
    assert "%s 不是最终版本：有缺口：原因 gaps" % v1o in notes
    # S1's month has no final for order: nothing of S1 is cleaned
    assert len(R.versions("S1", MONTH)) == 4
    assert os.path.isdir(layout.run_folder("S1", R2))


def test_a_final_appears_later_replaces_the_draft(v2):
    md = layout.monthly_dir(MONTH, MARKET)
    make("S1", R1, "20261003_050000", R.GAPS, "bill")
    FZ.run(FZ.Unit("ML", MARKET, MONTH, md, ("bill",), name, cname), ["S1"])
    assert os.path.isfile(os.path.join(md, "S1_2026-09_账单基准_未定稿.xlsx"))
    v, _r, _w = make("S1", R2, "20261012_050000", R.COMPLETE, "bill")
    FZ.run(FZ.Unit("ML", MARKET, MONTH, md, ("bill",), name, cname), ["S1"])
    assert read(os.path.join(md, "S1_2026-09_账单基准.xlsx")) == v
    assert not os.path.exists(os.path.join(md, "S1_2026-09_账单基准_未定稿.xlsx"))


def test_clean_up_once_every_basis_is_final(v2):
    md = layout.monthly_dir(MONTH, MARKET)
    old_b, raw1, wb_old_b = make("S2", R1, "20261003_050000", R.COMPLETE, "bill")      # superseded
    old_o, _r, wb_old_o = make("S2", R1, "20261003_050000", R.COMPLETE, "order")
    fin_b, raw2, wb_fin_b = make("S2", R2, "20261012_050000", R.COMPLETE, "bill")
    fin_o, _r, wb_fin_o = make("S2", R2, "20261012_050000", R.COMPLETE_WARN, "order")
    gap_b, raw3, wb_gap_b = make("S2", R3, "20261026_050000", R.GAPS, "bill")         # newer but gaps
    gap_o, _r, _w = make("S2", R3, "20261026_050000", R.FAILED, "order", workbook=False)
    # S1 shares the version folder of R1 and must keep its workbook there
    s1_v, _r, s1_wb = make("S1", R1, "20261003_050000", R.GAPS, "bill")
    # a raw folder never reported, older than the final (gone) and one newer (kept)
    unused_old = layout.run_folder("S2", "20261005_000000_m202609")
    unused_new = layout.run_folder("S2", "20261028_000000_m202609")
    odd = os.path.join(os.path.dirname(raw1), "manual_copy")          # not a run name: never touched
    for d in (unused_old, unused_new, odd):
        os.makedirs(d)
    pend = os.path.join(layout.logs_dir(), "pending.json")
    pending_store.save([{"store": "S2", "report": "r", "range": "a", "out_dir": raw3},
                        {"store": "S2", "report": "r", "range": "b", "out_dir": raw2}], pend)
    out = FZ.run(unit(), ["S2"], [], None)
    # finals kept, everything else of S2's month gone
    assert sorted((v["basis"], v["version"]) for v in R.versions("S2", MONTH)) == [("bill", fin_b),
                                                                                   ("order", fin_o)]
    assert os.path.isfile(wb_fin_b) and os.path.isfile(wb_fin_o)
    assert not os.path.exists(wb_old_b) and not os.path.exists(wb_old_o) and not os.path.exists(wb_gap_b)
    assert not os.path.isdir(os.path.dirname(wb_gap_b))               # R3's version folder emptied -> removed
    assert os.path.isfile(s1_wb) and os.path.isdir(os.path.dirname(s1_wb))   # S1's file in the shared folder
    assert os.path.isdir(raw2) and not os.path.exists(raw1) and not os.path.exists(raw3)
    assert not os.path.exists(unused_old) and os.path.isdir(unused_new) and os.path.isdir(odd)
    assert sorted(out["cleaned"]) == sorted([raw1, raw3, unused_old])
    assert [e["range"] for e in pending_store.load(pend)] == ["b"]
    assert read(os.path.join(md, name("S2", MONTH, "bill"))) == fin_b
    notes = read(os.path.join(md, FZ.NOTES))
    assert "删除没有被最终版本使用的下载目录：%s" % raw1 in notes and "删除旧版本结果文件" in notes
    # S1 untouched
    assert [v["version"] for v in R.versions("S1", MONTH)] == [s1_v]


def test_a_folder_another_month_reads_is_kept(v2):
    """September's report read an October-named run folder (it held September's
    bill): cleaning October must not delete it."""
    oct_run = "20261103_010000_m202610"
    v = R.version_id(oct_run, "20261003_050000")
    raw = layout.run_folder("S2", oct_run)
    os.makedirs(raw)
    R.write(R.document("ML", MARKET, "MX", "MXN", 2, "S2", MONTH, "bill", v, "2026-10-03T05:00:00", [raw], None,
                       {"level": R.GAPS, "reasons": []}))
    make("S2", "20261110_010000_m202610", "20261110_050000", R.COMPLETE, "bill", month="2026-10")
    oct_md = layout.monthly_dir("2026-10", MARKET)
    out = FZ.run(FZ.Unit("ML", MARKET, "2026-10", oct_md, ("bill",), name, cname), ["S2"])
    assert os.path.isdir(raw) and out["cleaned"] == []


def test_month_without_final_is_never_cleaned(v2):
    make("S1", R1, "20261003_050000", R.GAPS, "bill")
    make("S1", R2, "20261012_050000", R.FAILED, "bill", workbook=False)
    files = lambda: sorted(os.path.join(d, f) for d, _s, fs in os.walk(str(v2 / "data")) for f in fs)
    before = files()
    FZ.run(FZ.Unit("ML", MARKET, MONTH, layout.monthly_dir(MONTH, MARKET), ("bill",), name, cname), ["S1"])
    assert len(R.versions("S1", MONTH)) == 2
    assert os.path.isdir(layout.run_folder("S1", R1)) and os.path.isdir(layout.run_folder("S1", R2))
    md = layout.monthly_dir(MONTH, MARKET)
    added = [f for f in files() if f not in before]
    assert set(before) <= set(files())                        # nothing deleted
    assert sorted(os.path.basename(f) for f in added) == ["S1_2026-09_账单基准_未定稿.xlsx", FZ.NOTES]
    assert all(os.path.dirname(f) == md for f in added)


def test_nothing_outside_the_roots_is_deleted(v2, tmp_path):
    """A results file whose workbook lives elsewhere, a raw folder named in the
    results that is outside the store's tree: neither is touched."""
    outside_wb = tmp_path / "elsewhere" / "S2_2026-09_账单基准.xlsx"
    outside_wb.parent.mkdir()
    outside_wb.write_text("keep")
    outside_raw = tmp_path / "elsewhere" / "20261001_000000_m202609"
    outside_raw.mkdir()
    v = R.version_id(str(outside_raw), "20261001_050000")
    R.write(R.document("ML", MARKET, "MX", "MXN", 2, "S2", MONTH, "bill", v, "2026-10-01T05:00:00",
                       [str(outside_raw)], str(outside_wb), {"level": R.GAPS, "reasons": []}))
    make("S2", R2, "20261012_050000", R.COMPLETE, "bill")
    FZ.run(FZ.Unit("ML", MARKET, MONTH, layout.monthly_dir(MONTH, MARKET), ("bill",), name, cname), ["S2"])
    assert outside_wb.read_text() == "keep" and outside_raw.is_dir()
    assert len(R.versions("S2", MONTH, "bill")) == 1                     # its results file did go


def test_a_folder_that_is_not_a_month_folder_is_left_alone(v2, tmp_path):
    make("S2", R1, "20261003_050000", R.COMPLETE, "bill")
    make("S2", R2, "20261012_050000", R.COMPLETE, "bill")
    flat = tmp_path / "flat"
    flat.mkdir()
    out = FZ.run(FZ.Unit("ML", MARKET, MONTH, str(flat), ("bill",), name, cname), ["S2"], [], Builder())
    assert os.listdir(str(flat)) == [] and out["combined"] == [] and out["cleaned"] == []
    assert len(R.versions("S2", MONTH)) == 2


def test_legacy_layout_touches_no_file(tmp_path, monkeypatch):
    monkeypatch.setattr(paths, "ROOT", str(tmp_path))
    monkeypatch.setattr(store_config, "_stores", lambda: {"allowed": ["S1"]})
    flat = layout.reports_root()
    os.makedirs(flat)
    run = os.path.join(str(tmp_path), "downloads", "S1", R1)
    os.makedirs(run)
    for built, level in (("20261003_050000", R.COMPLETE), ("20261012_050000", R.GAPS)):
        R.write(R.document("ML", MARKET, "MX", "MXN", 2, "S1", MONTH, "bill", R.version_id(run, built),
                           "%s-%s-%sT05:00:00" % (built[:4], built[4:6], built[6:8]), [run],
                           os.path.join(flat, "S1_2026-09_账单基准.xlsx"), {"level": level, "reasons": []}))
    b = Builder()
    before = sorted(os.walk(str(tmp_path)))
    out = FZ.run(FZ.Unit("ML", MARKET, MONTH, flat, ("bill",), name, cname), ["S1"], [], b)
    assert out["finals"][("S1", "bill")]["grade"]["level"] == R.COMPLETE
    assert b.calls == [] and sorted(os.walk(str(tmp_path))) == before


# ------------------------------------------------------------------ a real runner in the new layout

def test_falabella_runner_in_the_new_layout(tmp_path, monkeypatch):
    """Falabella's run_reports (core/platform_reports) with data_root: the store
    workbook goes to _versions/<version>/, the final is copied to the month
    folder, the summary is built from it, 版本说明 is written; a second run of the
    same download makes a newer final and cleans the first version."""
    import test_fa_report as FAT
    from core import report_kit as K
    from platforms.falabella import run_reports
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"paths": {"data_root": str(tmp_path / "data")},
                               "stores": {"allowed": [FAT.STORE], "batches": {"FA-CL": [FAT.STORE]},
                                          "catalog": {FAT.STORE: {"platform": "FA", "country": "CL"}}}}),
                   encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    monkeypatch.setattr(K, "recalc", lambda out: (False, "跳过重算"))
    src = FAT.make_folder(tmp_path / "src")
    run = layout.run_folder(FAT.STORE, os.path.basename(src))
    shutil.copytree(src, run)
    md = layout.monthly_dir(FAT.MONTH, "FA-CL")
    stamps = iter(["20261022_100000", "20261022_110000"])
    monkeypatch.setattr(R, "built_stamp", lambda when=None: next(stamps))
    assert run_reports.main(["--month", FAT.MONTH, "--out-dir", md]) in (0, 1)
    [doc] = R.versions(FAT.STORE, FAT.MONTH, "main")
    first = doc["version"]
    assert doc["workbook"] == os.path.join(md, "_versions", first, "%s_2026-09.xlsx" % FAT.STORE)
    assert doc["grade"]["level"] in R.COUNTABLE                 # NOT_AVAILABLE invoice is not a gap
    assert sorted(x for x in os.listdir(md) if not x.startswith("_")) == [
        "CL_FA_全店汇总_2026-09.xlsx", "%s_2026-09.xlsx" % FAT.STORE, FZ.NOTES]
    assert run_reports.main(["--month", FAT.MONTH, "--out-dir", md]) in (0, 1)
    [doc] = R.versions(FAT.STORE, FAT.MONTH, "main")
    assert doc["version"] != first and os.listdir(os.path.join(md, "_versions")) == [doc["version"]]
    assert os.path.isdir(run)                                   # the final still uses it


def test_mercadolibre_bill_basis_runner_in_the_new_layout(v2, tmp_path, monkeypatch):
    """run_reports --basis bill with data_root: version folder, results file, the
    final (or 未定稿) copied, the bill summary built by finalize, 版本说明."""
    import sys
    import test_ml_bill_report as BT
    from core import report_kit as K
    from platforms.mercadolibre import run_reports
    cfg = json.loads((tmp_path / "config.json").read_text(encoding="utf-8"))
    cfg["stores"]["allowed"].append("STORE_MX")
    cfg["stores"]["catalog"]["STORE_MX"] = {"platform": "ML", "country": "MX"}
    (tmp_path / "config.json").write_text(json.dumps(cfg), encoding="utf-8")
    monkeypatch.setattr(K, "recalc", lambda out, timeout=180: (False, "跳过重算"))
    src = BT.make_mx_folder(tmp_path / "src")
    run = layout.run_folder("STORE_MX", "20261013_000000_m202609")
    shutil.copytree(src, run)
    md = layout.monthly_dir(MONTH, MARKET)
    monkeypatch.setattr(sys, "argv", ["run_reports", "--month", MONTH, "--stores", "STORE_MX", "--basis", "bill",
                                      "--out-dir", md])
    assert run_reports.main() in (0, 1)
    [doc] = R.versions("STORE_MX", MONTH, "bill")
    assert os.path.dirname(doc["workbook"]) == FZ.version_dir(md, doc["version"])
    ok = doc["grade"]["level"] in R.COUNTABLE
    final = "STORE_MX_2026-09_账单基准%s.xlsx" % ("" if ok else "_未定稿")
    summary = "MX_ML_全店汇总_2026-09_账单基准%s.xlsx" % ("" if ok else "_未定稿")
    assert sorted(os.listdir(md)) == sorted(["_versions", final, summary, FZ.NOTES])


def test_drop_pending_covers_every_platform_list(tmp_path, monkeypatch):
    from core import finalize as FZ
    monkeypatch.setattr(layout, "logs_dir", lambda config_path=None: str(tmp_path))
    gone = str(tmp_path / "raw" / "run_gone")
    for name in ("pending.json", "pending_FA.json", "pending_AZ.json"):
        pending_store.save([{"store": "S", "report": "r", "range": "a", "out_dir": gone},
                            {"store": "S", "report": "r", "range": "b", "out_dir": str(tmp_path / "kept")}],
                           str(tmp_path / name))
    lines = []
    assert FZ.drop_pending([gone], lambda text, keep=False: lines.append(text)) == 3
    for name in ("pending.json", "pending_FA.json", "pending_AZ.json"):
        assert [e["range"] for e in pending_store.load(str(tmp_path / name))] == ["b"]
