# -*- coding: utf-8 -*-
"""tools/backfill_results: dry-run lists and writes nothing; --apply runs the
platforms' run_reports for the store-months that have no results file yet."""
import json
import os
import shutil

import pytest

from core import layout
from core import paths
from core import report_kit as K
from core import results as R
from core import store_config
from tools import backfill_results as B

import test_fa_report as FAT
import test_ml_bill_report as BT

MONTH = "2026-09"
ML_STORE = "STORE_MX"


def configure(tmp_path, monkeypatch, v2=True):
    cfg = tmp_path / "config.json"
    doc = {"stores": {"allowed": [ML_STORE, FAT.STORE, "IDLE_MX"], "batch": [ML_STORE, "IDLE_MX"],
                      "batches": {"FA-CL": [FAT.STORE]},
                      "catalog": {ML_STORE: {"platform": "ML", "country": "MX"},
                                  "IDLE_MX": {"platform": "ML", "country": "MX"},
                                  FAT.STORE: {"platform": "FA", "country": "CL"}}}}
    if v2:
        doc["paths"] = {"data_root": str(tmp_path / "data")}
    cfg.write_text(json.dumps(doc), encoding="utf-8")
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    fa = layout.run_folder(FAT.STORE, "20261016_120000_m202609")
    shutil.copytree(FAT.make_folder(tmp_path / "src_fa"), fa)
    ml = layout.run_folder(ML_STORE, "20261013_000000_m202609")
    shutil.copytree(BT.make_mx_folder(tmp_path / "src_ml"), ml)
    return fa, ml


def files(root):
    return sorted(os.path.join(d, f) for d, _s, fs in os.walk(str(root)) for f in fs)


class Recorder(object):
    def __init__(self):
        self.cmds = []

    def __call__(self, cmd, **kw):
        self.cmds.append(cmd)
        return 0


def test_dry_run_lists_and_writes_nothing(tmp_path, monkeypatch):
    configure(tmp_path, monkeypatch)
    before = files(tmp_path)
    run, said = Recorder(), []
    assert B.main(["--month", MONTH, "--region", "ALL"], run=run, say=said.append) == 0
    assert run.cmds == [] and files(tmp_path) == before
    text = "\n".join(said)
    assert "预演（dry-run）" in text and "IDLE_MX" in text and "没有可用的下载目录" in text
    assert "MX 2026-09 要重出：%s" % ML_STORE in text and "FA-CL 2026-09 要重出：%s" % FAT.STORE in text


def test_apply_runs_one_command_per_unit_and_month(tmp_path, monkeypatch):
    configure(tmp_path, monkeypatch)
    run = Recorder()
    assert B.main(["--from", "2026-08", "--to", MONTH, "--region", "MX", "FA-CL", "--apply"], run=run,
                  say=lambda x: None) == 0
    by = dict((c[c.index("-m") + 1], c) for c in run.cmds)
    assert len(run.cmds) == 2                                  # August: no folder anywhere
    ml = by["platforms.mercadolibre.run_reports"]
    assert ml[ml.index("--region") + 1] == "MX" and ml[ml.index("--stores") + 1:ml.index("--out-dir")] == [ML_STORE]
    assert ml[-1] == layout.monthly_dir(MONTH, "ML-MX") and "--no-combined" not in ml
    fa = by["platforms.falabella.run_reports"]
    assert fa[fa.index("--market") + 1] == "FA-CL" and fa[-1] == layout.monthly_dir(MONTH, "FA-CL")


def test_store_months_with_results_are_skipped_unless_forced(tmp_path, monkeypatch):
    fa, _ml = configure(tmp_path, monkeypatch)
    R.write(R.document("FA", "FA-CL", "CL", "CLP", 0, FAT.STORE, MONTH, "main", R.version_id(fa, "20261020_000000"),
                       "2026-10-20T00:00:00", [fa], None, {"level": R.FAILED, "reasons": []}))
    run = Recorder()
    assert B.main(["--month", MONTH, "--region", "FA-CL", "--apply"], run=run, say=lambda x: None) == 0
    assert run.cmds == []
    assert B.main(["--month", MONTH, "--region", "FA-CL", "--apply", "--force"], run=run, say=lambda x: None) == 0
    assert len(run.cmds) == 1


def test_legacy_writes_elsewhere_and_no_summary(tmp_path, monkeypatch):
    monkeypatch.setattr(paths, "ROOT", str(tmp_path / "code"))
    configure(tmp_path, monkeypatch, v2=False)
    run = Recorder()
    assert B.main(["--month", MONTH, "--region", "FA-CL", "--apply"], run=run, say=lambda x: None) == 0
    [cmd] = run.cmds
    assert "--no-combined" in cmd and cmd[-1] == os.path.join(layout.reports_root(), "_backfill", MONTH)


def test_apply_end_to_end_with_falabella(tmp_path, monkeypatch):
    """--apply running Falabella's run_reports in-process: a results file appears,
    the month folder gets the final and the summary."""
    configure(tmp_path, monkeypatch)
    monkeypatch.setattr(K, "recalc", lambda out, timeout=180: (False, "跳过重算"))
    from platforms.falabella import run_reports

    def run(cmd, **kw):
        return run_reports.main(cmd[cmd.index("--month"):])
    assert B.main(["--month", MONTH, "--region", "FA-CL", "--apply"], run=run, say=lambda x: None) in (0, 1)
    [doc] = R.versions(FAT.STORE, MONTH, "main")
    md = layout.monthly_dir(MONTH, "FA-CL")
    assert os.path.isfile(os.path.join(md, "%s_2026-09.xlsx" % FAT.STORE))
    assert os.path.isfile(os.path.join(md, "CL_FA_全店汇总_2026-09.xlsx"))
    # a second backfill finds the results file and does nothing
    said = []
    assert B.main(["--month", MONTH, "--region", "FA-CL", "--apply"], run=run, say=said.append) == 0
    assert "\n没有要补录的店-月份。" in said


def test_bad_arguments(tmp_path, monkeypatch):
    configure(tmp_path, monkeypatch)
    with pytest.raises(store_config.StoreConfigError):
        B.main(["--region", "MX"], run=Recorder(), say=lambda x: None)
    with pytest.raises(store_config.StoreNotAllowed):
        B.main(["--month", MONTH, "--region", "MX", "--stores", "NOPE"], run=Recorder(), say=lambda x: None)
