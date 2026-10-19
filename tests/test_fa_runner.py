# -*- coding: utf-8 -*-
"""Step 3: Falabella in the one runner and the one window, next to MercadoLibre.

MercadoLibre's units keep their country codes and their commands; Falabella runs
as unit "FA-CL" with its own package, always with explicit stores. Child
processes are fake; config.json is a stub; no store is ever opened.
"""
import io
import json
import os
import sys

import pytest

from core import markets
from core import region_choice as RC
from core import run_monthly
from core import runlog
from core import store_config

PY = sys.executable
MONTH = "2026-09"
MX = ["BOCINA_SM", "EWTTO_SM"]
CL = ["Tecnohub_CL", "Andes_CL"]
FA = ["DAB_CL-Fala (Coco)", "AIS_CL-Fala"]

LEGACY = {"default": "BOCINA_SM", "allowed": MX + CL, "batch": MX,
          "regions": {"Tecnohub_CL": "CL", "Andes_CL": "CL"}, "batches": {"CL": CL}}

CATALOG = {"default": "BOCINA_SM", "allowed": MX + CL + FA,
           "catalog": dict([(n, {"platform": "ML", "country": "MX"}) for n in MX]
                           + [(n, {"platform": "ML", "country": "CL"}) for n in CL]
                           + [(n, {"platform": "FA", "country": "CL"}) for n in FA]),
           "batches": {"ML-MX": MX, "ML-CL": CL, "FA-CL": FA}}


def stub(monkeypatch, cfg):
    d = json.loads(json.dumps(cfg))
    monkeypatch.setattr(store_config, "_stores", lambda: d)
    return d


@pytest.fixture
def env(tmp_path, monkeypatch):
    logs, out = tmp_path / "logs", tmp_path / "out"
    monkeypatch.setattr(run_monthly, "LOG_DIR", str(logs))
    monkeypatch.setattr(run_monthly, "LOCK", str(logs / "monthly.lock"))
    monkeypatch.setattr(run_monthly, "config", lambda: {})
    monkeypatch.setattr(run_monthly, "known_stores", lambda: [])
    monkeypatch.setattr(run_monthly, "POLL", 0.1)

    class Env(object):
        log_dir, out_dir = str(logs), str(out)

        def run(self, argv, plans=None):
            return run_monthly.execute(run_monthly.parse_args(
                ["--month", MONTH, "--out-dir", self.out_dir] + argv), plan=plans)

        def last(self):
            rid = runlog.list_runs(self.log_dir)[0]["run_id"]
            p = runlog.run_paths(rid, self.log_dir)
            with io.open(p["log"], encoding="utf-8") as f:
                recs = [runlog.parse_line(x) for x in f.read().splitlines()]
            assert all(recs)
            return runlog.read_json(p["json"]), recs
    return Env()


def child(code, text=""):
    return [PY, "-c", "import sys\nsys.stdout.reconfigure(encoding='utf-8')\n"
                      "print(%r, flush=True)\nsys.exit(%d)\n" % (text, code)]


def plan(*steps):
    return [{"key": k, "label": k, "cmd": c} for k, c in steps]


def dry(env, capsys, argv):
    assert env.run(argv + ["--dry-run"]) == 0
    return [l.strip() for l in capsys.readouterr().out.splitlines()
            if "run_batch" in l or "run_reports" in l]


# ------------------------------------------------- MercadoLibre does not move

@pytest.mark.parametrize("step", ["both", "download", "report", "collect"])
def test_mercadolibre_commands_identical_with_new_config(env, capsys, monkeypatch, step):
    extra = [] if step == "both" else ["--step", step]
    stub(monkeypatch, LEGACY)
    old = {r: dry(env, capsys, extra + (["--region", r] if r != "MX" else [])) for r in ("MX", "CL")}
    stub(monkeypatch, CATALOG)
    new = {r: dry(env, capsys, extra + (["--region", r] if r != "MX" else [])) for r in ("MX", "CL")}
    assert new == old
    assert all("falabella" not in l for v in new.values() for l in v)


def test_alert_names(monkeypatch):
    assert run_monthly.alert_name(MONTH, "MX") == "！报表异常_202609.txt"
    assert run_monthly.alert_name(MONTH, "CL") == "！报表异常_CL_202609.txt"
    assert run_monthly.alert_name(MONTH, "FA-CL") == "！报表异常_CL_FA_202609.txt"


# ------------------------------------------------- Falabella unit

def test_falabella_unit_commands(env, capsys, monkeypatch):
    stub(monkeypatch, CATALOG)
    assert run_monthly.parse_args(["--region", "fa-cl"]).region == "FA-CL"
    cmds = dry(env, capsys, ["--region", "FA-CL"])
    dl, rp = cmds
    assert "-m platforms.falabella.run_batch --month 2026-09 --market FA-CL --stores" in dl
    assert "DAB_CL-Fala (Coco)" in dl and "AIS_CL-Fala" in dl
    assert "-m platforms.falabella.run_reports --month 2026-09 --market FA-CL --out-dir" in rp
    assert "--no-combined" not in rp
    sub = dry(env, capsys, ["--region", "FA-CL", "--stores", "AIS_CL-Fala"])
    assert sub[0].endswith("--stores AIS_CL-Fala")
    assert "--stores AIS_CL-Fala --no-combined" in sub[1]
    col = dry(env, capsys, ["--region", "FA-CL", "--step", "collect"])
    assert "--collect-only --month 2026-09 --market FA-CL --stores" in col[0]
    rerun = dry(env, capsys, ["--region", "FA-CL", "--stores", "AIS_CL-Fala",
                              "--reports-for", "all"])
    assert "--stores" not in rerun[1] and "--no-combined" not in rerun[1]


def test_stores_must_belong_to_the_unit(env, capsys, monkeypatch):
    stub(monkeypatch, CATALOG)
    # a Falabella store in a MercadoLibre Chile run, and the other way round: exit 3
    assert env.run(["--region", "CL", "--stores", "AIS_CL-Fala", "--dry-run"]) == 3
    assert env.run(["--region", "FA-CL", "--stores", "Tecnohub_CL", "--dry-run"]) == 3
    assert "不属于" in capsys.readouterr().out


def test_unit_without_stores_is_exit_4(env, capsys, monkeypatch):
    stub(monkeypatch, LEGACY)
    assert env.run(["--region", "FA-CL", "--dry-run"]) == 4


def test_all_runs_mercadolibre_first_then_falabella(env, capsys, monkeypatch):
    stub(monkeypatch, CATALOG)
    lineup, skipped = run_monthly.region_lineup()
    assert [u["region"] for u in lineup] == ["MX", "CL", "FA-CL"]
    assert skipped == ["PE"]
    stub(monkeypatch, LEGACY)
    lineup, skipped = run_monthly.region_lineup()
    assert [u["region"] for u in lineup] == ["MX", "CL"]          # no Falabella line at all
    assert skipped == ["PE"]


def test_all_regions_run_with_falabella(env, monkeypatch):
    stub(monkeypatch, CATALOG)
    code = env.run(["--region", "ALL", "--entrance", "scheduled"], {
        "MX": plan(("download", child(0)), ("report", child(0))),
        "CL": plan(("download", child(0)), ("report", child(0))),
        "FA-CL": plan(("download", child(1, "# [1/2] DAB_CL-Fala (Coco)\n  [放弃] stock")),
                      ("report", child(0, "单店报表：DAB_CL-Fala (Coco)"))),
    })
    assert code == 1
    doc, recs = env.last()
    assert [(r["region"], r["exit"]) for r in doc["regions"]] == [("MX", 0), ("CL", 0), ("FA-CL", 1)]
    fa = [r for r in recs if runlog.split_region(r["message"])[0] == "FA-CL"]
    assert fa and any("Falabella 智利（FA-CL）" in r["message"] for r in fa)
    assert any(r["store"] == "DAB_CL-Fala (Coco)" for r in fa)
    assert os.path.exists(os.path.join(env.out_dir, "！报表异常_CL_FA_202609.txt"))
    assert not os.path.exists(os.path.join(env.out_dir, "！报表异常_CL_202609.txt"))
    assert not os.path.exists(run_monthly.LOCK)


# ------------------------------------------------- window logic

def test_window_choices(monkeypatch):
    stub(monkeypatch, LEGACY)
    assert RC.labels() == ["墨西哥", "智利", "秘鲁"]                 # unchanged without Falabella
    assert RC.legacy_tasks()[:3] == [("MX", "MX月度报表"), ("CL", "CL月度报表"),
                                     ("PE", "PE月度报表")]
    assert all("FA" not in n for _, n in RC.legacy_tasks())
    stub(monkeypatch, CATALOG)
    assert RC.labels() == ["墨西哥", "智利", "秘鲁", "Falabella 智利"]
    assert RC.from_label("Falabella 智利") == "FA-CL"
    assert RC.from_label("智利") == "CL"
    assert RC.stores("FA-CL") == FA and RC.stores("CL") == CL
    assert [r for r, _ in RC.lineup(RC.ALL)] == ["MX", "CL", "FA-CL"]
    assert RC.remembered({"region": "fa-cl"}) == "FA-CL"
    assert RC.args_region("--entrance scheduled --region FA-CL") == "FA-CL"
    assert RC.args_region("--entrance scheduled --region ALL") == "ALL"
    assert RC.action_region({"store": "AIS_CL-Fala"}) == "FA-CL"
    assert RC.action_region({"store": "Tecnohub_CL"}) == "CL"
    plan_ = RC.rerun_plan([{"store": "AIS_CL-Fala", "title": "x", "rerun": True}])
    assert plan_["region"] == "FA-CL"
    mixed = RC.rerun_plan([{"store": "AIS_CL-Fala", "title": "x", "rerun": True},
                           {"store": "Tecnohub_CL", "title": "x", "rerun": True}])
    assert mixed["region"] is None                                  # one unit per run


def test_preview_command_uses_the_platform_package(monkeypatch):
    stub(monkeypatch, CATALOG)
    ml = ["-m", "platforms.mercadolibre.run_reports"]
    assert RC.preview_cmd(PY, ml, MONTH, None, "CL") == [PY] + ml + ["--month", MONTH, "--dry-run",
                                                                    "--region", "CL"]
    assert RC.preview_cmd(PY, ml, MONTH, ["AIS_CL-Fala"], "FA-CL") == [
        PY, "-m", "platforms.falabella.run_reports", "--month", MONTH, "--dry-run",
        "--market", "FA-CL", "--stores", "AIS_CL-Fala"]
    assert [r for r, _ in RC.preview_cmds(PY, ml, MONTH, None, RC.ALL)] == ["MX", "CL", "FA-CL"]


def test_log_tags_and_names():
    assert runlog.split_region("[FA-CL] hola") == ("FA-CL", "hola")
    assert runlog.split_region("[CL] hola") == ("CL", "hola")
    assert runlog.split_region("[XX-YY] hola") == (None, "[XX-YY] hola")
    assert runlog.region_name("FA-CL") == "Falabella 智利"
    assert runlog.region_name("CL") == "智利"
    assert markets.all_units() == ["MX", "CL", "PE", "FA-CL"]
