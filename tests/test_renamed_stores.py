# -*- coding: utf-8 -*-
"""Stores renamed in Ziniao (config.json stores.catalog "formerly").

The old name's results and download folders stay on disk; every reader folds
them into the current store: one store in the period report and the overview,
the final chosen across both folders, the newest download found under either
name. Synthetic data only; nothing opens a store.
"""
import io
import json
import os

import pytest

import period_data as P
from core import finalize as FZ
from core import layout
from core import month_overview as MO
from core import results as R
from core import store_config
from core.period_report import cli
from core.period_report import model as MD
from core.period_report import output as G

OLD, NEW = "清货2_EE02", "QH_EE02"
MONTH = "2026-09"
Q3 = MD.quarter("2026-Q3")
BILL = [("revenue", 120.0), ("fees_total", -20.0), ("contrib", 100.0, "f"), ("cogs", 0.0, "in"),
        ("profit", 100.0, "f")]


def _config(tmp_path, monkeypatch, catalog, v2=True, allowed=None):
    cfg = {"stores": {"allowed": list(catalog) if allowed is None else allowed, "catalog": catalog}}
    if v2:
        cfg["paths"] = {"data_root": str(tmp_path / "data")}
    else:
        from core import paths
        monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    p = tmp_path / "config.json"
    p.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(p))
    from core import run_monthly
    monkeypatch.setattr(run_monthly, "GUI_SETTINGS", str(tmp_path / "gui_settings.json"))


def _before(tmp_path, monkeypatch, v2=True):
    """config.json as it was before the rename: the store under its old name."""
    _config(tmp_path, monkeypatch, {OLD: {"platform": "ML", "country": "MX"},
                                    "S2": {"platform": "ML", "country": "MX"}}, v2)


def _after(tmp_path, monkeypatch, v2=True, formerly=None):
    """config.json after the rename: the new name, formerly the old one."""
    _config(tmp_path, monkeypatch, {NEW: {"platform": "ML", "country": "MX",
                                          "formerly": [OLD] if formerly is None else formerly},
                                    "S2": {"platform": "ML", "country": "MX"}}, v2)


def _put(store, month, built, level="complete", basis="bill", pl=None):
    d = P.doc("S1", month, built, level=level, basis=basis, pl=pl or BILL)
    d["store"] = store
    R.write(d)
    return d


# ------------------------------------------------------------------ config

def test_former_names_and_current_name(tmp_path, monkeypatch):
    _after(tmp_path, monkeypatch)
    assert store_config.former_names(NEW) == [OLD]
    assert store_config.former_names("S2") == []
    assert store_config.current_name(OLD) == NEW
    assert store_config.current_name(NEW) == NEW and store_config.current_name("S2") == "S2"
    assert store_config.former_map() == {OLD: NEW}
    assert store_config.config_problems() == []
    assert store_config.market_of(NEW) == "ML-MX"


def test_a_plain_string_and_several_names_are_accepted(tmp_path, monkeypatch):
    _after(tmp_path, monkeypatch, formerly=OLD)
    assert store_config.former_names(NEW) == [OLD]
    _after(tmp_path, monkeypatch, formerly=[OLD, " 更早的名字 ", OLD, ""])
    assert store_config.former_names(NEW) == [OLD, "更早的名字"]
    assert store_config.current_name("更早的名字") == NEW


def test_a_former_name_that_is_still_a_current_store_is_a_problem(tmp_path, monkeypatch):
    _config(tmp_path, monkeypatch, {NEW: {"platform": "ML", "country": "MX", "formerly": ["S2"]},
                                    "S2": {"platform": "ML", "country": "MX"}})
    probs = dict(store_config.config_problems())
    assert NEW in probs and "S2" in probs[NEW] and "S2" not in probs
    with pytest.raises(store_config.StoreConfigError):
        store_config.market_of(NEW)                         # the claiming store is held back …
    assert store_config.market_of("S2") == "ML-MX"          # … the other one runs as usual
    assert store_config.former_names(NEW) == [] and store_config.current_name("S2") == "S2"
    # an allowlisted store that the catalog does not know is still a current store
    _config(tmp_path, monkeypatch, {NEW: {"platform": "ML", "country": "MX", "formerly": [OLD]}},
            allowed=[NEW, OLD])
    assert NEW in dict(store_config.config_problems())


def test_a_former_name_claimed_by_two_stores_is_both_stores_problem(tmp_path, monkeypatch):
    _config(tmp_path, monkeypatch, {NEW: {"platform": "ML", "country": "MX", "formerly": [OLD]},
                                    "S2": {"platform": "ML", "country": "MX", "formerly": OLD},
                                    "S3": {"platform": "ML", "country": "MX"}})
    probs = dict(store_config.config_problems())
    assert set(probs) == {NEW, "S2"} and OLD in probs[NEW] and "S2" in probs[NEW]
    assert store_config.current_name(OLD) == OLD
    assert store_config.stores_for_market("ML-MX") == ["S3"]


def test_a_wrong_formerly_shape_or_self_reference_is_a_problem(tmp_path, monkeypatch):
    _config(tmp_path, monkeypatch, {NEW: {"platform": "ML", "country": "MX", "formerly": {"x": 1}},
                                    "S2": {"platform": "ML", "country": "MX", "formerly": ["S2"]}})
    probs = dict(store_config.config_problems())
    assert "写法不对" in probs[NEW] and "它自己" in probs["S2"]
    _config(tmp_path, monkeypatch, {NEW: {"platform": "ML", "country": "MX", "formerly": [OLD, 3]}})
    assert NEW in dict(store_config.config_problems())


def test_a_former_name_without_any_folder_is_harmless(tmp_path, monkeypatch):
    _after(tmp_path, monkeypatch)
    _put(NEW, MONTH, "20261003_000000")
    assert R.stores() == [NEW] and R.months(NEW) == [MONTH]
    assert R.final(NEW, MONTH, "bill")["version"].startswith("20261003")
    assert R.merged_former("ML", "MX", NEW) == []
    assert R.problems(NEW, MONTH) == []
    assert layout.run_folders(NEW) == []
    D = MD.gather("MX", Q3)
    assert D.stores == [NEW] and D.merged == []


# ------------------------------------------------------------------ results

def test_versions_of_both_folders_are_one_store_and_the_final_is_the_newest_countable(tmp_path, monkeypatch):
    _before(tmp_path, monkeypatch)
    _put(OLD, "2026-08", "20260903_000000")                                     # August, old name only
    _put(OLD, MONTH, "20261003_000000")                                         # complete, old name
    _put(OLD, MONTH, "20261003_000000", basis="order")
    _after(tmp_path, monkeypatch)
    _put(NEW, MONTH, "20261012_000000", level=R.GAPS)                           # newer, not countable
    assert R.stores() == [NEW]                                                  # the old name is not a store
    vs = R.versions(NEW, MONTH, "bill")
    assert [v["version"][:8] for v in vs] == ["20261012", "20261003"]           # newest first, both folders
    assert R.final(NEW, MONTH, "bill")["version"].startswith("20261003")        # the old folder's complete one
    assert R.final_at("ML", "MX", NEW, MONTH, "bill")["version"].startswith("20261003")
    assert R.has_files_at("ML", "MX", NEW, MONTH, "order")                     # only the old folder has it
    assert R.months(NEW) == ["2026-08", MONTH] and R.months_at("ML", "MX", NEW) == ["2026-08", MONTH]
    assert R.first_month_at("ML", "MX", NEW) == "2026-08"
    assert R.merged_former("ML", "MX", NEW) == [OLD]
    # a newer complete version under the new name becomes the final
    _put(NEW, MONTH, "20261026_000000", level=R.COMPLETE_WARN)
    assert R.final(NEW, MONTH, "bill")["version"].startswith("20261026")
    # new files go under the current name only
    assert os.path.isfile(R.path_for(NEW, MONTH, R.final(NEW, MONTH, "bill")["version"], "bill"))
    assert NEW in R.path_for(NEW, MONTH, "x__20261026_000000", "bill")


def test_a_store_with_results_only_under_the_old_name_is_listed_under_the_new_one(tmp_path, monkeypatch):
    _before(tmp_path, monkeypatch)
    _put(OLD, MONTH, "20261003_000000")
    _after(tmp_path, monkeypatch)
    assert not os.path.isdir(os.path.join(layout.results_root(), "ML", "MX", NEW))
    assert R.stores("ML", "MX") == [NEW]
    assert R.final(NEW, MONTH, "bill") is not None
    assert R.first_month_at("ML", "MX", NEW) == MONTH


def test_problems_say_which_folder(tmp_path, monkeypatch):
    _before(tmp_path, monkeypatch)
    d = os.path.dirname(R.path_for(OLD, MONTH, "x__20261003_040000", "bill"))
    os.makedirs(d)
    io.open(os.path.join(d, "x__20261003_040000.bill.json"), "w").write("{")
    _after(tmp_path, monkeypatch)
    probs = R.problems(NEW, MONTH)
    assert len(probs) == 1 and probs[0][0] == "x__20261003_040000.bill.json" and OLD in probs[0][1]


def test_without_the_formerly_link_the_old_folder_is_a_store_of_its_own(tmp_path, monkeypatch):
    _before(tmp_path, monkeypatch)
    _put(OLD, MONTH, "20261003_000000")
    _after(tmp_path, monkeypatch, formerly=[])
    _put(NEW, MONTH, "20261012_000000")
    assert R.stores("ML", "MX") == [NEW, OLD]
    assert [v["version"][:8] for v in R.versions(NEW, MONTH, "bill")] == ["20261012"]


def test_the_same_version_copied_under_both_names_is_read_once(tmp_path, monkeypatch):
    _before(tmp_path, monkeypatch)
    d = _put(OLD, MONTH, "20261003_000000")
    _after(tmp_path, monkeypatch)
    d["store"] = NEW
    R.write(d)
    assert len(R.versions(NEW, MONTH, "bill")) == 1


# ------------------------------------------------------------------ period report

def test_period_report_counts_a_renamed_store_once(tmp_path, monkeypatch):
    _before(tmp_path, monkeypatch)
    for m, b in (("2026-07", "20260803_000000"), ("2026-08", "20260903_000000")):
        _put(OLD, m, b)
        _put("S2", m, b)
    _put(OLD, MONTH, "20261003_000000")                             # September first under the old name
    _after(tmp_path, monkeypatch)
    _put(NEW, MONTH, "20261012_000000")                             # … then again under the new one
    _put("S2", MONTH, "20261012_000000")
    D = MD.gather("MX", Q3)
    assert D.stores == [NEW, "S2"]                                  # 2 stores, not 3
    assert D.missing() == [] and D.not_started() == []
    assert D.final(NEW, MONTH)["version"].startswith("20261012")
    assert D.final(NEW, "2026-07")["version"].startswith("20260803")
    assert D.store_total(NEW, "revenue") == 360.0 and D.total("revenue") == 720.0
    assert D.merged == [(OLD, NEW)]
    res = G.generate(Q3, "MX")
    assert res["status"] == G.COMPLETE and len(res["stores"]) == 2
    lines = G.summary_lines(res)
    assert "  店铺 2 家，3 个月；用了 6 个店铺月份的完整版本" in lines
    assert "  旧名 %s 的结果并入 %s" % (OLD, NEW) in lines
    # asked for by the old name, the report still reads the store
    assert MD.gather("MX", Q3, stores=[OLD]).stores == [NEW]
    assert cli.stray_stores([OLD], ["ML-MX"]) == []


def test_cli_preview_names_the_merged_former_name(tmp_path, monkeypatch, capsys):
    _before(tmp_path, monkeypatch)
    _put(OLD, MONTH, "20261003_000000")
    _after(tmp_path, monkeypatch)
    code = cli.main(["--quarter", "2026-Q3", "--region", "MX", "--preview"])
    out = capsys.readouterr().out
    assert "旧名 %s 的结果并入 %s" % (OLD, NEW) in out and OLD + "  " not in out
    assert code in (cli.OK, cli.INCOMPLETE)


# ------------------------------------------------------------------ month overview

def test_overview_shows_one_row_under_the_new_name(tmp_path, monkeypatch):
    _before(tmp_path, monkeypatch)
    _put(OLD, MONTH, "20261003_000000")
    _put(OLD, MONTH, "20261003_000000", basis="order",
         pl=[("net_sales", 110.0), ("order_fees_total", -8.0), ("bill_fees_total", -2.0),
             ("contrib", 100.0, "f"), ("profit", 100.0, "f")])
    _after(tmp_path, monkeypatch)
    _put(NEW, MONTH, "20261012_000000", level=R.GAPS)
    folder = layout.alert_dir(MONTH)
    os.makedirs(folder, exist_ok=True)
    said = []
    got = MO.build(MONTH, folder, say=said.append, log_dir=str(tmp_path / "logs"))
    assert got["error"] is None
    from openpyxl import load_workbook
    ws = load_workbook(got["path"])[MO.SHEET_OVERVIEW]
    rows = [r for r in ws.iter_rows(min_row=MO.HEADER_ROW + 1, values_only=True) if any(v is not None for v in r)]
    names = [r[2] for r in rows if not str(r[2]).startswith("小计")]
    assert names == [NEW, NEW, "S2", "S2"]                          # bill + order; no old-name row
    assert not any(OLD in str(r[2]) for r in rows)
    by = dict(((r[2], r[3]), r) for r in rows)
    assert by[(NEW, "账单基准 正式")][4] == "完整"                   # the old folder's complete final shown
    assert by[(NEW, "订单基准 参考")][4] == "完整"
    assert "  旧名 %s 的结果并入 %s" % (OLD, NEW) in said
    assert MO.stores_with_results(MONTH) == [("ML", "MX", NEW)]
    hist = load_workbook(got["path"])[MO.SHEET_HISTORY]
    hrows = [r for r in hist.iter_rows(min_row=MO.HEADER_ROW + 1, values_only=True) if r[0] is not None]
    assert len([r for r in hrows if r[0] == NEW]) == 3 and not any(r[0] == OLD for r in hrows)


# ------------------------------------------------------------------ download folders

@pytest.mark.parametrize("v2", [True, False])
def test_run_folders_include_the_old_name_newest_first(tmp_path, monkeypatch, v2):
    _before(tmp_path, monkeypatch, v2)
    old_run = layout.run_folder(OLD, "20261003_000000_m202609")
    old_aug = layout.run_folder(OLD, "20260903_000000_m202608")
    _after(tmp_path, monkeypatch, v2)
    new_run = layout.run_folder(NEW, "20261001_000000_m202609")        # older than the old name's
    for d in (old_run, old_aug, new_run):
        os.makedirs(d)
    assert [p for _n, p in layout.run_folders(NEW)] == [old_run, new_run, old_aug]
    assert [p for _n, p in layout.run_folders_for_month(NEW, MONTH)] == [old_run, new_run]
    # a platform runner picks the newest usable folder whichever name it sits under
    from core import platform_reports as PR
    from platforms.falabella import periods
    runner = PR.ReportRunner("FA", "platforms.falabella", "FA-CL", periods)
    io.open(os.path.join(old_run, "run_meta.json"), "w").write("{}")
    io.open(os.path.join(new_run, "run_meta.json"), "w").write("{}")
    assert runner.latest_folder(NEW, MONTH) == old_run
    os.remove(os.path.join(old_run, "run_meta.json"))
    assert runner.latest_folder(NEW, MONTH) == new_run


def test_mercadolibre_finder_sees_the_old_name(tmp_path, monkeypatch):
    from platforms.mercadolibre import financial_report as FR
    _before(tmp_path, monkeypatch)
    old_run = layout.run_folder(OLD, "20261003_000000_m202609")
    os.makedirs(old_run)
    _after(tmp_path, monkeypatch)
    assert FR.latest_run(NEW)[0] is None                       # empty folder: not usable, but found
    assert layout.run_folders(NEW)[0][1] == old_run


def test_clean_up_never_deletes_the_old_names_folders(tmp_path, monkeypatch):
    """A final under the new name: the superseded old-name version's results
    file goes, its raw folder (under the old name) stays."""
    _before(tmp_path, monkeypatch)
    old_raw = layout.run_folder(OLD, "20261003_000000_m202609")
    os.makedirs(old_raw)
    d = P.doc("S1", MONTH, "20261003_050000", pl=BILL)
    d["store"], d["source_folders"] = OLD, [old_raw]
    R.write(d)
    _after(tmp_path, monkeypatch)
    new_raw = layout.run_folder(NEW, "20261012_000000_m202609")
    os.makedirs(new_raw)
    md = layout.monthly_dir(MONTH, "ML-MX")
    v = R.version_id("20261012_000000_m202609", "20261012_050000")
    wb = os.path.join(FZ.version_dir(md, v), "%s_%s_账单基准.xlsx" % (NEW, MONTH))
    os.makedirs(os.path.dirname(wb))
    io.open(wb, "w").write("x")
    d2 = P.doc("S1", MONTH, "20261012_050000", pl=BILL)
    d2.update(store=NEW, source_folders=[new_raw], workbook=wb)
    R.write(d2)
    unit = FZ.Unit("ML", "ML-MX", MONTH, md, ("bill",),
                   lambda s, m, b: "%s_%s_账单基准.xlsx" % (s, m), lambda m, b: "MX_ML_全店汇总_%s.xlsx" % m)
    out = FZ.run(unit, [NEW], [], None)
    assert [x["version"] for x in R.versions(NEW, MONTH, "bill")] == [v]     # the old version's file cleaned
    assert os.path.isdir(old_raw) and os.path.isdir(new_raw)                 # the old-name folder untouched
    assert old_raw not in out["cleaned"]


def test_migration_places_an_old_names_folder_by_the_current_store(tmp_path, monkeypatch):
    from tools import migrate_layout as ML
    _after(tmp_path, monkeypatch, v2=False)
    assert ML.store_market(NEW, []) == ("ML", "MX", "")
    pf, cc, how = ML.store_market(OLD, [])
    assert (pf, cc) == ("ML", "MX") and NEW in how
