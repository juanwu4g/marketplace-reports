# -*- coding: utf-8 -*-
"""tools/migrate_layout.py on a fake machine: plan writes nothing; copy puts every
file where core/layout expects it, verifies it, rewrites pending paths and only
then switches config.json; remove-old deletes only verified old files; undo
switches back."""
import io
import json
import os

import pytest

from core import entrance
from core import layout
from core import paths
from core import store_config
from tools import migrate_layout as ML

RUN = "20261021_233002_m202609"


def write(p, text="x"):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with io.open(p, "w", encoding="utf-8") as fh:
        fh.write(text)


@pytest.fixture
def machine(tmp_path, monkeypatch):
    root = tmp_path / "code"
    monkeypatch.setattr(paths, "ROOT", str(root))
    cfg = root / "config.json"
    write(str(cfg), json.dumps({"stores": {
        "allowed": ["BOCINA_TA02", "DAB_CL-Fala (Coco)"],
        "catalog": {"BOCINA_TA02": {"platform": "ML", "country": "MX"},
                    "DAB_CL-Fala (Coco)": {"platform": "FA", "country": "CL"}}}}))
    monkeypatch.setattr(store_config, "CONFIG_PATH", str(cfg))
    monkeypatch.setattr(entrance, "SETTINGS", str(root / "gui_settings.json"))
    dl = root / "downloads"
    write(str(dl / "BOCINA_TA02" / RUN / "Reporte_Facturacion_MercadoLibre_Sept2026.xlsx"), "bill")
    write(str(dl / "BOCINA_TA02" / "20260909_021624" / "ventas.xlsx"), "old run without month")
    write(str(dl / "DAB_CL-Fala (Coco)" / "20261022_045756_m202609" / "run_meta.json"), "{}")
    write(str(dl / "DAB_CL (Coco) 子账号" / RUN / "run_meta.json"), json.dumps({"site": "cl", "month": "2026-09"}))
    write(str(dl / "NOBODY" / RUN / "a.txt"))                                   # cannot be placed
    write(str(dl / "mx_sales" / "__pycache__" / "x.pyc"))                     # leftover of the moved code
    fin = dl / "data" / "reports" / "financial"
    write(str(fin / "BOCINA_TA02_2026-09.xlsx"))
    write(str(fin / "BOCINA_TA02_2026-09_账单基准.xlsx"))
    write(str(fin / "MX_ML_全店汇总_2026-09.xlsx"))
    write(str(fin / "CL_FA_全店汇总_2026-09.xlsx"))
    write(str(fin / "！报表异常_CL_FA_202609.txt"))
    write(str(fin / "random.txt"))
    write(str(dl / "data" / "reports" / "BOCINA_TA02_20261021_233002.xlsx"))  # cleaning workbook
    write(str(dl / "data" / "processed" / "ventas_mx" / "part.parquet"))
    write(str(dl / "data" / "mx_sales.duckdb"))
    write(str(root / "logs" / "runs" / "r.log"))
    write(str(root / "logs" / "pending.json"), json.dumps({"entries": [
        {"store": "BOCINA_TA02", "report": "x", "range": "y", "out_dir": str(dl / "BOCINA_TA02" / RUN)}]}))
    write(str(root / "gui_settings.json"), json.dumps({"out_dir": str(fin), "region": "MX"}))
    return tmp_path


def test_plan_writes_nothing(machine, capsys):
    data = machine / "data"
    assert ML.main(["plan", "--data-root", str(data)]) == 0
    assert not data.exists()
    out = capsys.readouterr().out
    assert "NOBODY" in out and "放不进新布局" in out


def test_copy_switch_and_remove_old(machine, capsys):
    root, data = machine / "code", machine / "data"
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    assert layout.mode() == layout.V2 and layout.configured_data_root() == str(data)
    # raw folders where layout looks for them
    assert os.path.isfile(os.path.join(layout.run_folder("BOCINA_TA02", RUN),
                                       "Reporte_Facturacion_MercadoLibre_Sept2026.xlsx"))
    assert layout.run_folders("BOCINA_TA02")[-1][0] == "20260909_021624"            # no month -> 其他, still found
    assert (data / "raw" / "FA" / "CL" / "DAB_CL-Fala (Coco)" / "2026-09" / "20261022_045756_m202609").is_dir()
    assert (data / "raw" / "ML" / "CL" / "DAB_CL (Coco) 子账号" / "2026-09" / RUN).is_dir()   # by run_meta site
    assert not (data / "raw" / "ML" / "MX" / "NOBODY").exists()
    # reports sorted by month and market
    m = data / "reports" / "monthly" / "2026-09"
    for p in (m / "ML-MX" / "BOCINA_TA02_2026-09.xlsx", m / "ML-MX" / "BOCINA_TA02_2026-09_账单基准.xlsx",
              m / "ML-MX" / "MX_ML_全店汇总_2026-09.xlsx", m / "FA-CL" / "CL_FA_全店汇总_2026-09.xlsx",
              m / "！报表异常_CL_FA_202609.txt", data / "reports" / "_旧文件" / "random.txt",
              data / "work" / "mx_sales" / "processed" / "ventas_mx" / "part.parquet",
              data / "work" / "mx_sales" / "mx_sales.duckdb",
              data / "work" / "mx_sales" / "reports" / "BOCINA_TA02_20261021_233002.xlsx",
              data / "logs" / "runs" / "r.log"):
        assert p.is_file(), p
    # pending entries follow the moved folder; the old default report folder setting is dropped
    pend = json.load(open(str(data / "logs" / "pending.json"), encoding="utf-8"))
    assert pend["entries"][0]["out_dir"] == layout.run_folder("BOCINA_TA02", RUN)
    assert "out_dir" not in json.load(open(str(root / "gui_settings.json"), encoding="utf-8"))
    # nothing old was touched
    assert (root / "downloads" / "BOCINA_TA02" / RUN).is_dir()
    manifest = [os.path.join(str(data / "logs" / "migration"), f)
                for f in os.listdir(str(data / "logs" / "migration"))][0]
    # remove-old: an old file changed after the copy is kept
    write(str(root / "downloads" / "data" / "reports" / "financial" / "random.txt"), "changed")
    removed, kept = ML.remove_old(manifest)
    assert removed > 10
    assert [k for k, _ in kept] == [str(root / "downloads" / "data" / "reports" / "financial" / "random.txt")]
    assert not (root / "downloads" / "BOCINA_TA02").exists()
    assert (root / "downloads" / "NOBODY" / RUN / "a.txt").is_file()                # never copied, never deleted
    assert not (root / "logs").exists()


def test_copy_again_is_harmless_and_conflicts_stop_the_switch(machine):
    data = machine / "data"
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    assert ML.main(["copy", "--data-root", str(data)]) == 0                         # identical: skipped
    ML.undo()
    assert layout.mode() == layout.LEGACY
    write(os.path.join(str(data), "raw", "ML", "MX", "BOCINA_TA02", "2026-09", RUN,
                       "Reporte_Facturacion_MercadoLibre_Sept2026.xlsx"), "different")
    assert ML.main(["copy", "--data-root", str(data)]) == 1
    assert layout.mode() == layout.LEGACY                                            # not switched


def test_refusals(machine):
    root = machine / "code"
    assert ML.main(["copy", "--data-root", str(root / "data")]) == 2                # inside the code folder
    write(str(root / "logs" / "monthly.lock"), "running")
    assert ML.main(["plan", "--data-root", str(machine / "data")]) == 2             # a run is active
    os.remove(str(root / "logs" / "monthly.lock"))
    assert ML.main(["remove-old", "--manifest", "x.json"]) == 2                     # still legacy


def test_undo_drops_the_new_default_report_folder_setting(machine):
    root, data = machine / "code", machine / "data"
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    # the window saved the folder it showed (the new default)
    write(str(root / "gui_settings.json"), json.dumps({"out_dir": str(data / "reports"), "region": "MX"}))
    ML.undo()
    assert "out_dir" not in json.load(open(str(root / "gui_settings.json"), encoding="utf-8"))
    assert layout.reports_root() == str(root / "downloads" / "data" / "reports" / "financial")
