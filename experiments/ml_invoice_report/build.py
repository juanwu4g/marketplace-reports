# -*- coding: utf-8 -*-
"""One store-month: load → derive → validate → (SKU costs) → workbook → recalc."""
import datetime
import os

from core import report_kit as K
from core import sku_costs as SC
from core.paths import ROOT
from experiments.ml_invoice_report import model as MD
from experiments.ml_invoice_report import workbook as WB
from experiments.ml_invoice_report import mapping as M

PLATFORM = "ML"
DEFAULT_OUT = os.path.join(ROOT, "downloads", "data", "reports", "experiments")


def workbook_name(store, month):
    return "%s_%s_账单基准报表.xlsx" % (store, month)


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


def run(store, folder, month, site, out_dir=None, sku_costs=None, recalc=True, build_date=None):
    """Build the workbook. sku_costs: None (no table), a path, or an SC.Table.
    Returns dict(out, checks, S, recalc, scan)."""
    S = MD.load(store, folder, month, site)
    MD.derive(S)
    checks = MD.validate(S)
    table = SC.load(sku_costs) if isinstance(sku_costs, str) else sku_costs
    S["cost"] = None
    if table is not None:
        S["cost"] = SC.apply(table, PLATFORM, S["prof"].country, S["prof"].currency, month, cost_items(S))
        checks += SC.checks(table, [(store, S["cost"])], PLATFORM, S["prof"].country)
    ctx = {"month": month, "period_cn": MD.period_cn(month), "checks": checks, "sku_costs": table,
           "built": build_date or datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}
    out_dir = out_dir or DEFAULT_OUT
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, workbook_name(store, month))
    WB.build(S, ctx, out)
    rc = scan = None
    if recalc:
        rc = K.recalc(out)
        if rc[0]:
            scan = K.scan_formula_errors(out)
    return {"out": out, "checks": checks, "S": S, "recalc": rc, "scan": scan}
