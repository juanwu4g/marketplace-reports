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


def test_results_files_follow_the_moved_folders_and_workbooks(machine):
    """F1(a): a results file written in the old layout names the old run folder
    and the old flat workbook; its copy names their new places (core/finalize
    decides from source_folders which download folder a final uses)."""
    root, data = machine / "code", machine / "data"
    dl = root / "downloads"
    old_raw = str(dl / "BOCINA_TA02" / RUN)
    old_wb = str(dl / "data" / "reports" / "financial" / "BOCINA_TA02_2026-09_账单基准.xlsx")
    rel = os.path.join("ML", "MX", "BOCINA_TA02", "2026-09", RUN + "__20261022_010000.bill.json")
    doc = {"schema": 1, "store": "BOCINA_TA02", "month": "2026-09", "basis": "bill", "run": RUN,
           "source_folders": [old_raw, "/somewhere/else/20260101_000000_m202601"], "workbook": old_wb,
           "grade": {"level": "complete", "reasons": []}, "pl": []}
    write(str(dl / "data" / "results" / rel), json.dumps(doc, ensure_ascii=False))
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    new = json.load(open(str(data / "results" / rel), encoding="utf-8"))
    assert new["source_folders"] == [layout.run_folder("BOCINA_TA02", RUN), "/somewhere/else/20260101_000000_m202601"]
    assert new["workbook"] == str(data / "reports" / "monthly" / "2026-09" / "ML-MX" /
                                  "BOCINA_TA02_2026-09_账单基准.xlsx")
    assert new["grade"] == doc["grade"] and new["run"] == RUN
    # the old file is untouched; copying again finds the identical rewritten file
    assert json.load(open(str(dl / "data" / "results" / rel), encoding="utf-8")) == doc
    ML.undo()
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    # remove-old accepts the rewritten copy (its checksum is the one recorded)
    manifest = sorted(os.listdir(str(data / "logs" / "migration")))[-1]
    removed, kept = ML.remove_old(os.path.join(str(data / "logs" / "migration"), manifest))
    assert not (dl / "data" / "results").exists() and not [k for k, why in kept if "results" in k]


# ------------------------------------------------------------ the window's own report folder

@pytest.fixture
def shared(machine):
    """「报表输出到」 names a folder on a shared drive, not the old default: the
    old layout's flat report folder, the new layout's reports root. Its files
    are sorted in place, under that same folder."""
    root = machine / "code"
    out = machine / "S" / "Finance" / "电商数据"
    write(str(out / "BOCINA_TA02_2026-09.xlsx"), "store")
    write(str(out / "BOCINA_TA02_2026-09_指标.json"), "{}")
    write(str(out / "MX_ML_全店汇总_2026-09.xlsx"), "combined")
    write(str(out / "！需要处理_202609.txt"), "alert")
    write(str(out / "！报表异常_CL_FA_202608.txt"), "old alert")
    write(str(out / "本月总览_2026-09.xlsx"), "overview")
    write(str(out / "期间报表" / "quarterly" / "2026-Q3" / "ML-MX" / "MX_ML_2026-Q3_账单基准.xlsx"), "q3")
    write(str(out / "stray.txt"), "stray")
    write(str(out / "_versions" / "v1" / "BOCINA_TA02_2026-09.xlsx"), "version")
    write(str(root / "gui_settings.json"), json.dumps({"out_dir": str(out), "region": "MX"}))
    return out


def _out_items(plan, out):
    return dict((src, dst) for _k, src, dst in plan["items"] if ML._under(src, str(out)))


def test_plan_sorts_the_window_folder_in_place(shared, machine, capsys):
    out, data = shared, machine / "data"
    plan = ML.build_plan(str(data))
    got = _out_items(plan, out)
    m = out / "monthly" / "2026-09"
    want = {
        out / "BOCINA_TA02_2026-09.xlsx": m / "ML-MX" / "BOCINA_TA02_2026-09.xlsx",
        out / "BOCINA_TA02_2026-09_指标.json": m / "ML-MX" / "BOCINA_TA02_2026-09_指标.json",
        out / "MX_ML_全店汇总_2026-09.xlsx": m / "ML-MX" / "MX_ML_全店汇总_2026-09.xlsx",
        out / "！需要处理_202609.txt": m / "！需要处理_202609.txt",
        out / "！报表异常_CL_FA_202608.txt": out / "monthly" / "2026-08" / "！报表异常_CL_FA_202608.txt",
        out / "本月总览_2026-09.xlsx": m / "本月总览_2026-09.xlsx",
        out / "期间报表" / "quarterly" / "2026-Q3" / "ML-MX" / "MX_ML_2026-Q3_账单基准.xlsx":
            out / "quarterly" / "2026-Q3" / "ML-MX" / "MX_ML_2026-Q3_账单基准.xlsx",
        out / "stray.txt": out / "_旧文件" / "stray.txt",
        out / "_versions" / "v1" / "BOCINA_TA02_2026-09.xlsx":
            out / "_旧文件" / "_versions" / "v1" / "BOCINA_TA02_2026-09.xlsx",
    }
    assert got == dict((str(k), str(v)) for k, v in want.items())
    # never under data_root/reports; the old default folder's files still go there
    assert not [d for d in got.values() if ML._under(d, str(data))]
    assert str(data / "reports" / "monthly" / "2026-09" / "ML-MX" / "BOCINA_TA02_2026-09.xlsx") in \
        [d for _k, _s, d in plan["items"]]
    assert ML.main(["plan", "--data-root", str(data)]) == 0
    assert not (out / "monthly").exists() and not data.exists()
    text = capsys.readouterr().out
    assert "就地整理" in text and "9 个旧文件" in text and "2 个放进 _旧文件" in text


def test_copy_sorts_the_window_folder_in_place(shared, machine, capsys):
    root, out, data = machine / "code", shared, machine / "data"
    dl = root / "downloads"
    rel = os.path.join("ML", "MX", "BOCINA_TA02", "2026-09", RUN + "__20261022_010000.bill.json")
    old_wb = str(out / "BOCINA_TA02_2026-09.xlsx")
    write(str(dl / "data" / "results" / rel),
          json.dumps({"schema": 1, "store": "BOCINA_TA02", "month": "2026-09", "run": RUN,
                      "source_folders": [], "workbook": old_wb, "pl": []}))
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    text = capsys.readouterr().out
    assert "那里原有的 9 个旧文件已经按新布局整理" in text
    # where core/layout will look for them in the new layout, under the same root
    assert layout.mode() == layout.V2
    assert (os.path.join(layout.monthly_dir("2026-09", "ML-MX", override=str(out)), "BOCINA_TA02_2026-09.xlsx")
            == str(out / "monthly" / "2026-09" / "ML-MX" / "BOCINA_TA02_2026-09.xlsx"))
    for p in (out / "monthly" / "2026-09" / "ML-MX" / "BOCINA_TA02_2026-09.xlsx",
              out / "monthly" / "2026-09" / "ML-MX" / "MX_ML_全店汇总_2026-09.xlsx",
              out / "monthly" / "2026-09" / "！需要处理_202609.txt",
              out / "monthly" / "2026-09" / "本月总览_2026-09.xlsx",
              out / "quarterly" / "2026-Q3" / "ML-MX" / "MX_ML_2026-Q3_账单基准.xlsx",
              out / "_旧文件" / "stray.txt"):
        assert p.is_file(), p
    assert str(out / "quarterly" / "2026-Q3" / "ML-MX") == layout.period_dir("quarterly", "2026-Q3", "ML-MX",
                                                                               override=str(out))
    # originals untouched; the setting is kept (it is where the user wants reports)
    assert (out / "BOCINA_TA02_2026-09.xlsx").read_text(encoding="utf-8") == "store"
    assert (out / "期间报表" / "quarterly" / "2026-Q3" / "ML-MX" / "MX_ML_2026-Q3_账单基准.xlsx").is_file()
    assert json.load(open(str(root / "gui_settings.json"), encoding="utf-8"))["out_dir"] == str(out)
    # the results file follows the workbook
    new = json.load(open(str(data / "results" / rel), encoding="utf-8"))
    assert new["workbook"] == str(out / "monthly" / "2026-09" / "ML-MX" / "BOCINA_TA02_2026-09.xlsx")
    # the manifest knows the folder, so remove-old deletes the flat originals (and only them)
    mdir = str(data / "logs" / "migration")
    manifest = os.path.join(mdir, sorted(os.listdir(mdir))[-1])
    doc = json.load(open(manifest, encoding="utf-8"))
    assert doc["out_dir"] == str(out)
    assert any(ML._under(it["src"], str(out)) for it in doc["items"])
    # a second copy skips the identical copies; the sorted files are not sorted again
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    assert len(_out_items(ML.build_plan(str(data)), out)) == 9
    assert not (out / "monthly" / "monthly").exists() and not (out / "_旧文件" / "monthly").exists()
    # undo leaves everything where it is (the old layout still finds the flat originals)
    ML.undo()
    assert layout.mode() == layout.LEGACY
    assert (out / "monthly" / "2026-09" / "ML-MX" / "BOCINA_TA02_2026-09.xlsx").is_file()
    assert (out / "BOCINA_TA02_2026-09.xlsx").is_file()
    assert json.load(open(str(root / "gui_settings.json"), encoding="utf-8"))["out_dir"] == str(out)
    assert ML.main(["copy", "--data-root", str(data)]) == 0
    removed, kept = ML.remove_old(manifest)
    assert not [k for k, _ in kept if ML._under(k, str(out))]
    for gone in (out / "BOCINA_TA02_2026-09.xlsx", out / "stray.txt", out / "期间报表", out / "_versions",
                 out / "本月总览_2026-09.xlsx"):
        assert not gone.exists(), gone
    assert sorted(os.listdir(str(out))) == ["_旧文件", "monthly", "quarterly"]
    assert (out / "monthly" / "2026-09" / "ML-MX" / "BOCINA_TA02_2026-09.xlsx").is_file()
    # nothing left to sort: the next plan lists no file of the folder
    assert _out_items(ML.build_plan(str(data)), out) == {}


def test_window_folder_that_cannot_be_opened_refuses(shared, machine, capsys):
    root, data = machine / "code", machine / "data"
    write(str(root / "gui_settings.json"), json.dumps({"out_dir": str(machine / "S" / "gone"), "region": "MX"}))
    assert ML.main(["plan", "--data-root", str(data)]) == 2
    assert "打不开" in capsys.readouterr().out


def test_window_folder_inside_old_downloads_is_moved_the_ordinary_way(machine, capsys):
    root, data = machine / "code", machine / "data"
    inside = root / "downloads" / "data" / "reports" / "mine"
    write(str(inside / "BOCINA_TA02_2026-09.xlsx"))
    write(str(root / "gui_settings.json"), json.dumps({"out_dir": str(inside), "region": "MX"}))
    plan = ML.build_plan(str(data))
    assert [d for _k, s, d in plan["items"] if ML._under(s, str(inside))] == \
        [str(data / "work" / "_旧文件" / "reports" / "mine" / "BOCINA_TA02_2026-09.xlsx")]
    assert ML.main(["plan", "--data-root", str(data)]) == 0
    assert "没有就地整理" in capsys.readouterr().out
