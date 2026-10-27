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
    out = FZ.run(unit(), ["S1", "S2"], ["S1", "S2"], b)
    assert out["finals"][("S1", "bill")]["version"] == v1b and out["finals"][("S1", "order")] is None
    # final copied under the usual name; no final -> newest version as 未定稿
    assert read(os.path.join(md, name("S1", MONTH, "bill"))) == v1b
    assert read(os.path.join(md, "S1_2026-09_未定稿.xlsx")) == v2o
    assert not os.path.exists(os.path.join(md, name("S1", MONTH, "order")))
    # summaries from the chosen versions; 未定稿 when a store is not final
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
    gap_b, raw3, wb_gap_b = make("S2", R3, "20261026_050000", R.GAPS, "bill")         # newer but gaps: kept
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
    # R3 is newer than the final and has no countable version (its report has gaps,
    # e.g. MercadoPago still generating): the re-download is kept, so is its owed entry
    assert os.path.isdir(raw2) and not os.path.exists(raw1) and os.path.isdir(raw3)
    assert not os.path.exists(unused_old) and os.path.isdir(unused_new) and os.path.isdir(odd)
    assert sorted(out["cleaned"]) == sorted([raw1, unused_old])
    assert [e["range"] for e in pending_store.load(pend)] == ["a", "b"]
    assert read(os.path.join(md, name("S2", MONTH, "bill"))) == fin_b
    notes = read(os.path.join(md, FZ.NOTES))
    assert "删除没有被最终版本使用的下载目录：%s" % raw1 in notes and "删除旧版本结果文件" in notes
    assert "保留 %s" % raw3 in notes
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


def test_month_folder_keeps_only_finals_and_drafts(v2):
    """A plain-named file in the month folder for a store-basis without a final
    (e.g. an old report the layout migration brought in) is moved to
    _versions/_旧文件, never deleted; so is a plain-named summary when the new
    summary is a draft."""
    u = unit()
    os.makedirs(u.out_dir, exist_ok=True)
    old_store = os.path.join(u.out_dir, name("S1", MONTH, "bill"))
    old_summary = os.path.join(u.out_dir, cname(MONTH, "bill"))
    for p in (old_store, old_summary):
        with io.open(p, "w", encoding="utf-8") as fh:
            fh.write("sent in September")
    make("S1", "20261003_000000_m202609", "20261003_040000", R.GAPS)
    FZ.run(u, ["S1"], build_combined=Builder(), say=lambda *a: None)
    assert not os.path.exists(old_store) and not os.path.exists(old_summary)
    assert os.path.isfile(os.path.join(u.out_dir, FZ.draft_name(name("S1", MONTH, "bill"))))
    retired = os.listdir(os.path.join(u.out_dir, FZ.VERSIONS, FZ.RETIRED))
    assert sorted(f.split("_", 2)[2] for f in retired) == sorted([name("S1", MONTH, "bill"), cname(MONTH, "bill")])
    assert all(read(os.path.join(u.out_dir, FZ.VERSIONS, FZ.RETIRED, f)) == "sent in September" for f in retired)


# ------------------------------------------------------------------ clean-up never takes evidence it should keep

def _bill_unit():
    return FZ.Unit("ML", MARKET, MONTH, layout.monthly_dir(MONTH, MARKET), ("bill",), name, cname)


def test_migrated_final_keeps_its_raw_folder(v2):
    """F1(b): a final made before tools/migrate_layout holds the old path
    <ROOT>/downloads/<store>/<run>; its copy in the new tree has the same run
    name and must survive a rebuild whose new version is not countable."""
    legacy = os.path.join(paths.ROOT, "downloads", "S1", R1)
    os.makedirs(legacy)
    v = R.version_id(R1, "20261003_050000")
    R.write(R.document("ML", MARKET, "MX", "MXN", 2, "S1", MONTH, "bill", v, "2026-10-03T05:00:00",
                       [legacy], os.path.join(paths.ROOT, "downloads", "data", "x.xlsx"),
                       {"level": R.COMPLETE, "reasons": []},
                       figures={"pl": [{"key": "x", "label": "X", "original": "", "kind": "d", "value": 1.0}]}))
    migrated = layout.run_folder("S1", R1)
    os.makedirs(migrated)
    make("S1", R2, "20261012_050000", R.GAPS, "bill")
    lines = []
    out = FZ.run(_bill_unit(), ["S1"], say=lines.append)       # F1(c): the final's workbook is gone: no crash
    assert R.final("S1", MONTH, "bill")["version"] == v
    assert os.path.isdir(migrated) and out["cleaned"] == []
    notes = read(os.path.join(layout.monthly_dir(MONTH, MARKET), FZ.NOTES))
    assert "的工作簿不在了" in notes and "保留 %s" % migrated in notes


def test_bill_only_step_keeps_the_order_final_folder(v2):
    """F2: a step that builds only the bill basis must not delete the run folder
    the order basis' final reads, nor clean while the order basis has no final."""
    make("S1", R1, "20261003_050000", R.COMPLETE, "order")
    make("S1", R2, "20261012_050000", R.FAILED, "order", workbook=False)
    make("S1", R1, "20261003_050000", R.GAPS, "bill")
    make("S1", R2, "20261012_050000", R.COMPLETE, "bill")
    out = FZ.run(_bill_unit(), ["S1"])
    assert os.path.isdir(layout.run_folder("S1", R1)) and out["cleaned"] == []
    assert len(R.versions("S1", MONTH, "order")) == 2          # the order basis' files are its own step's


def test_bill_only_step_waits_for_an_order_final(v2):
    make("S1", R1, "20261003_050000", R.GAPS, "order")
    make("S1", R1, "20261003_050000", R.GAPS, "bill")
    make("S1", R2, "20261012_050000", R.COMPLETE, "bill")
    out = FZ.run(_bill_unit(), ["S1"])
    assert out["cleaned"] == [] and len(R.versions("S1", MONTH, "bill")) == 2
    assert os.path.isdir(layout.run_folder("S1", R1))


def test_newer_folder_with_only_failed_versions_is_kept(v2):
    """F3: a re-download newer than the final whose report crashed is kept."""
    _v, raw1, _w = make("S1", R1, "20261003_050000", R.COMPLETE, "bill")
    _v, raw3, _w = make("S1", R3, "20261026_050000", R.FAILED, "bill", workbook=False)
    out = FZ.run(_bill_unit(), ["S1"])
    assert os.path.isdir(raw1) and os.path.isdir(raw3) and out["cleaned"] == []


def test_folder_with_owed_reports_is_kept_and_listed(v2):
    """F3: an older folder that a pending*.json entry still points at is kept
    and its entry stays (another platform's list counts too)."""
    _v, raw1, _w = make("S1", R1, "20261003_050000", R.GAPS, "bill")
    _v, raw2, _w = make("S1", R2, "20261012_050000", R.COMPLETE, "bill")
    pend = os.path.join(layout.logs_dir(), "pending_FA.json")
    pending_store.save([{"store": "S1", "report": "r", "range": "a", "out_dir": raw1}], pend)
    out = FZ.run(_bill_unit(), ["S1"])
    assert os.path.isdir(raw1) and out["cleaned"] == []
    assert [e["out_dir"] for e in pending_store.load(pend)] == [raw1]


def test_unreadable_pending_list_keeps_every_folder(v2):
    _v, raw1, _w = make("S1", R1, "20261003_050000", R.GAPS, "bill")
    make("S1", R2, "20261012_050000", R.COMPLETE, "bill")
    os.makedirs(layout.logs_dir(), exist_ok=True)
    bad = os.path.join(layout.logs_dir(), "pending.json")
    with io.open(bad, "w", encoding="utf-8") as fh:
        fh.write("{broken")
    out = FZ.run(_bill_unit(), ["S1"])
    assert os.path.isdir(raw1) and out["cleaned"] == [] and os.path.isfile(bad)   # read only, not moved


def test_newer_is_judged_by_the_time_in_the_run_name(v2):
    """F3: the final reads a folder named for the next month (it held this
    month's bill); this month's folders are older or newer by their time."""
    oct_run = "20261103_010000_m202610"
    oct_raw = layout.run_folder("S1", oct_run)
    os.makedirs(oct_raw)
    R.write(R.document("ML", MARKET, "MX", "MXN", 2, "S1", MONTH, "bill", R.version_id(oct_run, "20261103_050000"),
                       "2026-11-03T05:00:00", [oct_raw], None, {"level": R.COMPLETE, "reasons": []}))
    older = layout.run_folder("S1", "20261020_000000_m202609")
    newer = layout.run_folder("S1", "20261105_000000_m202609")
    for d in (older, newer):
        os.makedirs(d)
    out = FZ.run(_bill_unit(), ["S1"])
    assert out["cleaned"] == [older] and os.path.isdir(newer) and os.path.isdir(oct_raw)


def test_final_with_an_unreadable_run_time_keeps_every_newer_candidate(v2):
    odd = os.path.join(os.path.dirname(layout.run_folder("S1", R1)), "manual_copy")
    os.makedirs(odd)
    R.write(R.document("ML", MARKET, "MX", "MXN", 2, "S1", MONTH, "bill", R.version_id(odd, "20261003_050000"),
                       "2026-10-03T05:00:00", [odd], None, {"level": R.COMPLETE, "reasons": []}))
    other = layout.run_folder("S1", R1)
    os.makedirs(other)
    out = FZ.run(_bill_unit(), ["S1"])
    assert out["cleaned"] == [] and os.path.isdir(other)


def test_failed_copy_is_reported_and_blocks_the_clean_up(v2, monkeypatch):
    """F9: the final could not be copied into the month folder (open in Excel):
    the failure is in 版本说明.txt and in the result, and the store is not
    cleaned (the old version is the only other copy)."""
    md = layout.monthly_dir(MONTH, MARKET)
    _v, raw1, wb1 = make("S1", R1, "20261003_050000", R.COMPLETE, "bill")
    _v, raw2, _w = make("S1", R2, "20261012_050000", R.COMPLETE, "bill")
    real = shutil.copy2

    def locked(src, dst, *a, **k):
        raise PermissionError("locked")
    monkeypatch.setattr(shutil, "copy2", locked)
    out = FZ.run(_bill_unit(), ["S1"])
    monkeypatch.setattr(shutil, "copy2", real)
    assert out["publish_failed"] == [os.path.join(md, name("S1", MONTH, "bill"))]
    assert out["cleaned"] == [] and os.path.isdir(raw1) and os.path.isfile(wb1)
    assert len(R.versions("S1", MONTH, "bill")) == 2
    notes = read(os.path.join(md, FZ.NOTES))
    assert "[未能复制]" in notes and "[跳过清理] S1" in notes
    assert not [f for f in os.listdir(md) if f.startswith(".copy_")]


def test_runner_exit_code_is_1_when_a_copy_failed(tmp_path, monkeypatch, capsys):
    """F9 in a real runner (Falabella's): the copy failure makes the exit code 1
    and is in the summary line run_monthly reads."""
    import test_fa_report as FAT
    from core import report_kit as K
    from platforms.falabella import run_reports
    md = _fa_machine(tmp_path, monkeypatch, FAT, K)
    assert run_reports.main(["--month", FAT.MONTH, "--out-dir", md, "--no-combined"]) == 0
    monkeypatch.setattr(FZ, "_copy", lambda src, dst, log: log("[未能复制] %s" % dst, keep=True) and False)
    capsys.readouterr()
    assert run_reports.main(["--month", FAT.MONTH, "--out-dir", md, "--no-combined"]) == 1
    summary = [ln for ln in capsys.readouterr().out.splitlines() if ln.startswith("[REPORT_SUMMARY_JSON]")][-1]
    assert "未能更新：%s_2026-09.xlsx" % FAT.STORE in json.loads(summary.split(" ", 1)[1])["attention"]


def _fa_machine(tmp_path, monkeypatch, FAT, K):
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"paths": {"data_root": str(tmp_path / "data")},
                               "stores": {"allowed": [FAT.STORE], "batches": {"FA-CL": [FAT.STORE]},
                                          "catalog": {FAT.STORE: {"platform": "FA", "country": "CL"}}}}),
                   encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    monkeypatch.setattr(K, "recalc", lambda out: (False, "跳过重算"))
    src = FAT.make_folder(tmp_path / "src")
    shutil.copytree(src, layout.run_folder(FAT.STORE, os.path.basename(src)))
    return layout.monthly_dir(FAT.MONTH, "FA-CL")


def test_hand_run_without_out_dir_lands_in_the_month_folder(tmp_path, monkeypatch):
    """F17: in the new layout a runner started without --out-dir writes into the
    month folder (so finalize publishes and cleans), not under <reports>/_versions."""
    import test_fa_report as FAT
    from core import platform_reports as PR
    from core import report_kit as K
    from platforms.falabella import run_reports
    md = _fa_machine(tmp_path, monkeypatch, FAT, K)
    assert PR.default_out_dir(FAT.MONTH, "FA-CL") == md
    assert run_reports.main(["--month", FAT.MONTH]) in (0, 1)
    [doc] = R.versions(FAT.STORE, FAT.MONTH, "main")
    assert os.path.dirname(doc["workbook"]) == FZ.version_dir(md, doc["version"])
    assert os.path.isfile(os.path.join(md, "%s_2026-09.xlsx" % FAT.STORE))
    assert not os.path.isdir(os.path.join(layout.reports_root(), FZ.VERSIONS))


def test_default_out_dir_in_the_legacy_layout_is_the_flat_folder(tmp_path, monkeypatch):
    from core import platform_reports as PR
    monkeypatch.setattr(paths, "ROOT", str(tmp_path))
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(tmp_path / "config.json"))
    assert PR.default_out_dir("2026-09", "FA-CL") == layout.reports_root() == PR.default_out_dir()


def test_mercadolibre_hand_run_without_out_dir(v2, tmp_path, monkeypatch):
    """F17 for MercadoLibre's run_reports."""
    import sys
    import test_ml_bill_report as BT
    from core import report_kit as K
    from platforms.mercadolibre import run_reports
    cfg = json.loads((tmp_path / "config.json").read_text(encoding="utf-8"))
    cfg["stores"]["allowed"].append("STORE_MX")
    cfg["stores"]["catalog"]["STORE_MX"] = {"platform": "ML", "country": "MX"}
    (tmp_path / "config.json").write_text(json.dumps(cfg), encoding="utf-8")
    monkeypatch.setattr(K, "recalc", lambda out, timeout=180: (False, "跳过重算"))
    shutil.copytree(BT.make_mx_folder(tmp_path / "src"), layout.run_folder("STORE_MX", "20261013_000000_m202609"))
    monkeypatch.setattr(sys, "argv", ["run_reports", "--month", MONTH, "--stores", "STORE_MX", "--basis", "bill"])
    assert run_reports.main() in (0, 1)
    [doc] = R.versions("STORE_MX", MONTH, "bill")
    assert os.path.dirname(doc["workbook"]) == FZ.version_dir(layout.monthly_dir(MONTH, MARKET), doc["version"])


def test_store_without_any_version_makes_the_summary_a_draft(v2):
    """F15: S3 is in the market's store list but has no version this month: the
    summary lacks it, so it is 未定稿, and 版本说明 says why."""
    md = layout.monthly_dir(MONTH, MARKET)
    make("S1", R1, "20261003_050000", R.COMPLETE, "bill")
    b = Builder()
    out = FZ.run(_bill_unit(), ["S1"], ["S1", "S3"], b)
    assert b.calls == [("bill", [("S1", R.version_id(R1, "20261003_050000"))],
                        FZ.draft_name(cname(MONTH, "bill")))]
    assert out["combined"][0]["stores_absent"] == ["S3"]
    notes = read(os.path.join(md, FZ.NOTES))
    assert "S3：本月没有任何版本" in notes and "没有任何版本：S3" in notes
    b = Builder()
    FZ.run(_bill_unit(), ["S1"], ["S1"], b)                     # every listed store final: plain name
    assert b.calls[0][2] == cname(MONTH, "bill")
    assert not os.path.exists(os.path.join(md, FZ.draft_name(cname(MONTH, "bill"))))
