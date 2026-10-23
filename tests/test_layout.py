# -*- coding: utf-8 -*-
"""core/layout.py: the legacy layout is exactly where things always were; the v2
layout (config.json → paths.data_root) puts every kind of data in its own tree,
and the code that writes or finds folders follows it."""
import io
import json
import os

import pytest

from core import layout
from core import paths
from core import store_config


@pytest.fixture
def legacy(tmp_path, monkeypatch):
    monkeypatch.setattr(paths, "ROOT", str(tmp_path))
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"stores": {"allowed": ["BOCINA_TA02"]}}), encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    return tmp_path


@pytest.fixture
def v2(tmp_path, monkeypatch):
    monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    data = tmp_path / "data"
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({
        "paths": {"data_root": str(data)},
        "stores": {"allowed": ["BOCINA_TA02", "DAB_CL-Fala (Coco)"],
                   "catalog": {"BOCINA_TA02": {"platform": "ML", "country": "MX"},
                               "DAB_CL-Fala (Coco)": {"platform": "FA", "country": "CL"}}}}),
        encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    return data


def test_legacy_is_where_things_always_were(legacy):
    root = str(legacy)
    assert layout.mode() == layout.LEGACY
    assert layout.raw_root() == os.path.join(root, "downloads")
    assert layout.run_folder("BOCINA_TA02", "20261021_233002_m202609") == \
        os.path.join(root, "downloads", "BOCINA_TA02", "20261021_233002_m202609")
    assert layout.reports_root() == os.path.join(root, "downloads", "data", "reports", "financial")
    assert layout.monthly_dir("2026-09", "MX") == layout.reports_root()       # flat, as always
    assert layout.alert_dir("2026-09") == layout.reports_root()
    assert layout.logs_dir() == os.path.join(root, "logs")
    assert layout.mx_sales_paths()["processed"] == os.path.join(root, "downloads", "data", "processed")


def test_v2_tree(v2):
    data = str(v2)
    assert layout.mode() == layout.V2
    assert layout.run_folder("BOCINA_TA02", "20261021_233002_m202609") == os.path.join(
        data, "raw", "ML", "MX", "BOCINA_TA02", "2026-09", "20261021_233002_m202609")
    assert layout.run_folder("DAB_CL-Fala (Coco)", "20261022_045756_m202609") == os.path.join(
        data, "raw", "FA", "CL", "DAB_CL-Fala (Coco)", "2026-09", "20261022_045756_m202609")
    assert layout.run_folder("BOCINA_TA02", "20261021_233002_history").endswith(
        os.path.join("BOCINA_TA02", layout.NO_MONTH, "20261021_233002_history"))
    assert layout.monthly_dir("2026-09", "MX") == os.path.join(data, "reports", "monthly", "2026-09", "ML-MX")
    assert layout.monthly_dir("2026-09", "FA-CL", override="/share/r") == os.path.join(
        os.path.abspath("/share/r"), "monthly", "2026-09", "FA-CL")
    assert layout.alert_dir("2026-09") == os.path.join(data, "reports", "monthly", "2026-09")
    assert layout.period_dir("quarterly", "2026-Q3", "MX") == os.path.join(
        data, "reports", "quarterly", "2026-Q3", "ML-MX")
    assert layout.results_dir("BOCINA_TA02", "2026-09") == os.path.join(
        data, "results", "ML", "MX", "BOCINA_TA02", "2026-09")
    assert layout.logs_dir() == os.path.join(data, "logs")
    assert layout.mx_sales_paths()["raw"] == os.path.join(data, "raw", "ML", "MX")


def test_run_folders_across_months_newest_first(v2):
    for name in ("20261001_000000_m202608", "20261021_233002_m202609", "20261016_141741_m202609"):
        os.makedirs(layout.run_folder("BOCINA_TA02", name))
    names = [n for n, _ in layout.run_folders("BOCINA_TA02")]
    assert names == ["20261021_233002_m202609", "20261016_141741_m202609", "20261001_000000_m202608"]
    sept = [n for n, _ in layout.run_folders_for_month("BOCINA_TA02", "2026-09")]
    assert sept == ["20261021_233002_m202609", "20261016_141741_m202609"]


def test_month_of_run():
    assert layout.month_of_run("20261021_233002_m202609") == "2026-09"
    assert layout.month_of_run("/x/y/20261021_233002_m202609/") == "2026-09"
    assert layout.month_of_run("20261021_233002") is None
    assert layout.month_of_run("20261021_233002_history") is None


def test_unreadable_config_never_guesses_a_data_folder(tmp_path, monkeypatch):
    cfg = tmp_path / "config.json"
    cfg.write_text("{not json", encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    with pytest.raises(store_config.StoreConfigError):
        layout.raw_root()
    assert layout.logs_dir()                       # logs still have somewhere to go


def test_ml_finder_follows_v2(v2):
    from platforms.mercadolibre import financial_report as FR
    d = layout.run_folder("BOCINA_TA02", "20261021_233002_m202609")
    os.makedirs(d)
    assert FR.latest_run("BOCINA_TA02")[0] is None            # empty folder: not usable, but found
    assert layout.run_folders("BOCINA_TA02")[0][1] == d


def test_platform_finder_follows_v2(v2):
    from core import platform_reports as PR
    from platforms.falabella import periods

    store = "DAB_CL-Fala (Coco)"
    d = layout.run_folder(store, "20261022_045756_m202609")
    os.makedirs(d)
    io.open(os.path.join(d, "run_meta.json"), "w").write("{}")
    runner = PR.ReportRunner("FA", "platforms.falabella", "FA-CL", periods)
    assert runner.latest_folder(store, "2026-09") == d
    assert runner.latest_folder(store, "2026-08") is None


def test_runner_puts_reports_and_alerts_in_month_folders(v2):
    from core import run_monthly as RM
    plan = RM.build_plan("report", "2026-09", None, str(v2 / "reports"), py="py", all_stores=["X"],
                         region="CL")
    cmd = plan[0]["cmd"]
    assert cmd[cmd.index("--out-dir") + 1] == os.path.join(str(v2), "reports", "monthly", "2026-09", "ML-CL")
    assert RM.alert_path(str(v2 / "reports"), "2026-09", "CL") == os.path.join(
        str(v2), "reports", "monthly", "2026-09", RM.alert_name("2026-09", "CL"))


def test_runner_unchanged_in_legacy(legacy):
    from core import run_monthly as RM
    plan = RM.build_plan("report", "2026-09", None, "/out", py="py")
    cmd = plan[0]["cmd"]
    assert cmd[cmd.index("--out-dir") + 1] == "/out"
    assert RM.alert_path("/out", "2026-09") == os.path.join(os.path.abspath("/out"), RM.alert_name("2026-09"))


def test_cleaning_pipeline_is_told_where_the_data_is(v2):
    from platforms.mercadolibre import transform
    env = transform.data_env()
    assert env["MX_SALES_RAW_DIR"] == os.path.join(str(v2), "raw", "ML", "MX")
    assert env["MX_SALES_WAREHOUSE"] == os.path.join(str(v2), "work", "mx_sales", "mx_sales.duckdb")
    assert transform.PIPELINE_ROOT.endswith("pipelines")


def test_subset_run_keeps_the_summary_only_in_v2(v2):
    from core import run_monthly as RM
    cmd = RM.build_plan("report", "2026-09", ["BOCINA_TA02"], str(v2 / "reports"), py="py")[0]["cmd"]
    assert "--no-combined" not in cmd and "--stores" in cmd


def test_subset_run_skips_the_summary_in_legacy(legacy):
    from core import run_monthly as RM
    cmd = RM.build_plan("report", "2026-09", ["BOCINA_TA02"], "/out", py="py")[0]["cmd"]
    assert "--no-combined" in cmd
