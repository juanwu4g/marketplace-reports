# -*- coding: utf-8 -*-
"""Store-month: prepare (load → derive → validate → SKU costs), then the store
workbook; a country's prepared stores → the combined workbook. Each workbook is
recalculated when LibreOffice is there."""
import datetime
import os
import time
import traceback

from core import report_family as RF
from core import report_kit as K
from core import sku_costs as SC
from core import layout
from core.paths import ROOT
from platforms.mercadolibre.bill_report import combined as CB
from platforms.mercadolibre.bill_report import model as MD
from platforms.mercadolibre.bill_report import workbook as WB
from platforms.mercadolibre.bill_report import mapping as M

PLATFORM = "ML"
DEFAULT_OUT = layout.reports_root()
SUFFIX = "_账单基准"


def workbook_name(store, month):
    return "%s_%s%s.xlsx" % (store, month, SUFFIX)


def combined_name(month, region):
    """The bill basis' 全店汇总: the order-based name plus the suffix."""
    return "%s_ML_全店汇总_%s%s.xlsx" % (str(region).upper(), month, SUFFIX)


def site_of(folder, site=None):
    """--site wins; else run_meta.json's "site"; else None (the caller must ask)."""
    if site:
        return str(site).lower()
    meta = MD.R.run_meta(folder)
    return (meta.get("site") or "").lower() or None


def cost_items(S):
    """One SC.Item per real SKU of ⑦ (pseudo rows and unmapped listings have no cost)."""
    items = []
    for k, x in S["sku"].items():
        if k.startswith("（") or k.startswith(M.UNMAPPED_PREFIX):
            continue
        items.append(SC.Item(k, S["sku_titles"].get(k, ""), int(round(x["units"])),
                             int(round(x["cost_units_excluded"])), x["net_gmv"]))
    return items


def load_costs(sku_costs):
    return SC.load(sku_costs) if isinstance(sku_costs, str) else sku_costs


def prepare(store, folder, month, site, table=None):
    """Everything a workbook needs for one store. Returns (S, checks)."""
    S = MD.load(store, folder, month, site)
    MD.derive(S)
    checks = MD.validate(S)
    S["cost"] = None
    if table is not None:
        S["cost"] = SC.apply(table, PLATFORM, S["prof"].country, S["prof"].currency, month, cost_items(S))
        checks += SC.checks(table, [(store, S["cost"])], PLATFORM, S["prof"].country)
    return S, checks


def _ctx(month, checks, table, build_date):
    return {"month": month, "period_cn": MD.period_cn(month), "checks": checks, "sku_costs": table,
            "built": build_date or datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}


def _finish(out, recalc):
    rc = scan = None
    if recalc:
        rc = K.recalc(out)
        if rc[0]:
            scan = K.scan_formula_errors(out)
    return rc, scan


def write_store(S, checks, out_dir=None, table=None, recalc=True, build_date=None):
    ctx = _ctx(S["month"], checks, table, build_date)
    out_dir = out_dir or DEFAULT_OUT
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, workbook_name(S["store"], S["month"]))
    WB.build(S, ctx, out)
    rc, scan = _finish(out, recalc)
    return {"out": out, "checks": checks, "S": S, "recalc": rc, "scan": scan}


def write_combined(prepared, month, region, out_dir=None, table=None, recalc=True, build_date=None,
                   order_combined="", name=None):
    """prepared: [(S, checks)] of one country's stores. name: the file name
    (default combined_name; core/finalize passes the 未定稿 one)."""
    Ss = [S for S, _ in prepared]
    checks = [c for _, cs in prepared for c in cs]
    ctx = _ctx(month, checks, table, build_date)
    ctx["order_combined"] = order_combined
    out_dir = out_dir or DEFAULT_OUT
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, name or combined_name(month, region))
    CB.build(Ss, ctx, out)
    rc, scan = _finish(out, recalc)
    return {"out": out, "checks": checks, "recalc": rc, "scan": scan}


def run(store, folder, month, site, out_dir=None, sku_costs=None, recalc=True, build_date=None):
    """One store workbook. sku_costs: None (no table), a path, or an SC.Table.
    Returns dict(out, checks, S, recalc, scan)."""
    table = load_costs(sku_costs)
    S, checks = prepare(store, folder, month, site, table)
    return write_store(S, checks, out_dir, table, recalc, build_date)


# ------------------------------------------------------------------ results for run_reports

def outcome(res):
    """The order report's outcome words (financial_report.OUTCOME_LABELS) for a
    bill-basis build: crashed, formula_errors, checks_failed or ok."""
    if res.get("crashed"):
        return "crashed"
    scan = res.get("scan")
    if scan is not None and scan[1]:
        return "formula_errors"
    if RF.failed_errors(res.get("checks") or []):
        return "checks_failed"
    return "ok"


def guarded(label, fn, *a, **kw):
    """Run one build; never raise. Returns a result dict with outcome, seconds,
    label and (on a crash) the error text."""
    t0 = time.time()
    try:
        res = fn(*a, **kw)
    except Exception as e:                       # noqa: BLE001 — one store must not stop the rest
        res = {"crashed": True, "checks": [], "out": "", "output": "%s: %s" % (e.__class__.__name__, e),
               "traceback": traceback.format_exc()}
    res["seconds"] = int(time.time() - t0)
    res["label"] = label
    res["outcome"] = outcome(res)
    return res
