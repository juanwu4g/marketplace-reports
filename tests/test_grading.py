# -*- coding: utf-8 -*-
"""core/grading: the grade of a monthly report version (docs/RESULTS_AND_PERIODS.md §3)
and core/results' helpers for the writers (P&L evaluation, documents)."""
import pytest

from core import grading as G
from core import report_family as RF
from core import report_kit as K
from core import results as R


def ck(key, ok=False, severity="error", skipped=False, name=None):
    return K.check(key, "S1", name or key, ok, "detail", severity=severity, skipped=skipped)


def level(*a, **k):
    return G.grade(*a, **k)["level"]


def test_failed_when_no_workbook():
    g = G.grade("ML", "bill", False, [ck("residuals")], why_not_written="生成出错")
    assert g["level"] == R.FAILED and "生成出错" in g["reasons"][0]


def test_complete_when_nothing_failed():
    assert G.grade("FA", "main", True, [ck("pl_tie", ok=True), ck("iva", severity="warn")]) == \
        {"level": R.COMPLETE, "reasons": []}


def test_warn_level_failures_never_change_the_grade():
    assert level("ML", "order", True, [ck("mp_settlement", severity="warn")]) == R.COMPLETE


def test_skipped_error_check_is_not_a_failure():
    assert level("ML", "order", True, [ck("bridge_b", skipped=True)]) == R.COMPLETE


def test_blocking_error_is_a_gap():
    g = G.grade("ML", "bill", True, [ck("residuals", name="全部残差为 0")])
    assert g["level"] == R.GAPS and g["reasons"] == ["校验未过：全部残差为 0"]


@pytest.mark.parametrize("platform,basis,key", sorted(G.NON_BLOCKING))
def test_non_blocking_errors_make_complete_warn(platform, basis, key):
    g = G.grade(platform, basis, True, [ck(key, name="名")])
    assert g["level"] == R.COMPLETE_WARN and g["reasons"] == ["提醒（不影响金额）：名"]
    assert G.NON_BLOCKING[(platform, basis, key)]           # every entry has its written reason


def test_non_blocking_is_per_platform_and_basis():
    assert level("ML", "order", True, [ck("map_cover")]) == R.GAPS      # map_cover is bill-only
    assert level("ML", "bill", True, [ck("cn_orphan")]) == R.GAPS
    assert level("FA", "main", True, [ck("cn_orphan")]) == R.GAPS


def test_missing_forms():
    def m(form, cause, in_report=None):
        return {"form": form, "cause": cause, "in_report": in_report}
    # the report reads it and it was not downloaded
    g = G.grade("FA", "main", True, [], [m("stock", "DOWNLOAD_NOT_LANDED", True)])
    assert g["level"] == R.GAPS and g["reasons"] == ["缺下载：stock（文件没有落盘）"]
    # nothing to download: no data, or the platform does not offer it (yet)
    assert level("FA", "main", True, [], [m("order/x", "NO_DATA", True)]) == R.COMPLETE
    assert level("FA", "main", True, [], [m("billing1/2026-09-21_2026-10-20", "NOT_AVAILABLE", True)]) == R.COMPLETE
    # the report does not read it
    assert level("AZ", "main", True, [], [m("invoice", "UNKNOWN", False)]) == R.COMPLETE
    # MercadoLibre writes no in_report: core/download_diag decides (MP collection is not read,
    # BILL_MP_V3 is not read; a bill is)
    assert level("ML", "bill", True, [], [m("mercadopago/collection", "NO_PERMISSION"),
                                          m("billing/2026-09/BILL_MP_V3", "UI_NOT_FOUND")]) == R.COMPLETE
    assert level("ML", "bill", True, [], [m("billing/2026-09/unidentified", "DOWNLOAD_NOT_LANDED")]) == R.GAPS


def test_data_incomplete_and_formula_errors_are_gaps():
    g = G.grade("ML", "order", True, [], data_incomplete=True)
    assert g["level"] == R.GAPS and "No pudimos mostrar algunos datos" in g["reasons"][0]
    assert level("FA", "main", True, [], formula_errors=True) == R.GAPS


def test_provisional_after_gaps_before_warn():
    assert level("ML", "bill", True, [ck("map_cover")], provisional=["Oct2026.xlsx"]) == R.PROVISIONAL
    assert level("ML", "bill", True, [ck("residuals")], provisional=["Oct2026.xlsx"]) == R.GAPS
    g = G.grade("ML", "order", True, [], provisional=["Octubre 2026"])
    assert g["reasons"] == ["账单尚未关账：Octubre 2026（关账后重出）"]


# ------------------------------------------------------------------ core/results helpers

def test_pl_rows_evaluates_like_the_workbook():
    spec = [{"key": "a", "label": "A", "kind": "d", "value": 100.0},
            {"key": "b", "label": "B", "kind": "d", "value": -30.0},
            {"key": "c", "label": "C", "kind": "d", "value": -20.0},
            {"key": "fees", "label": "F", "kind": "sum", "formula": RF.Span("b", "c")},
            {"key": "contrib", "label": "G", "kind": "f", "formula": ("a", "fees")},
            {"key": "cogs", "label": "H", "kind": "in", "value": 10.0},
            {"key": "profit", "label": "P", "kind": "f", "formula": ("contrib", "-cogs")}]
    got = dict((r["key"], r["value"]) for r in R.pl_rows(spec))
    assert got == {"a": 100.0, "b": -30.0, "c": -20.0, "fees": -50.0, "contrib": 50.0, "cogs": 10.0,
                   "profit": 40.0}
    named = [{"key": "x", "name": "行X", "label": "X", "kind": "d", "value": 2.0},
             {"key": "y", "name": "行Y", "label": "Y", "kind": "f", "formula": ("行X", "行X")}]
    assert R.pl_rows(named)[1]["value"] == 4.0
    with pytest.raises(ValueError):
        R.pl_rows([{"key": "z", "label": "Z", "kind": "f"}])


def test_version_parts_and_document():
    v = R.version_id("/x/20261021_233002_m202609", "20261022_010203")
    assert R.version_parts(v) == ("20261021_233002_m202609", "20261022_010203")
    assert R.version_parts("junk") == (None, None)
    doc = R.document("ML", "ML-MX", "MX", "MXN", 2, "S1", "2026-09", "bill", v, "2026-10-22T01:02:03",
                     ["/x/20261021_233002_m202609"], None, {"level": R.FAILED, "reasons": ["x"]},
                     missing=[{"form": "a", "cause": "UNKNOWN", "explain": "dropped"}, {"cause": "no form"}])
    assert not R.validate(doc)
    assert doc["run"] == "20261021_233002_m202609" and doc["workbook"] is None
    assert doc["missing"] == [{"form": "a", "cause": "UNKNOWN", "in_report": None}]
    assert doc["pl"] == [] and doc["inventory"] is None and set(doc["code"]) == {"commit", "fingerprint"}
    assert doc["flags"] == {"provisional": [], "data_incomplete": False}
