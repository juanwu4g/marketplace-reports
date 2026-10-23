# -*- coding: utf-8 -*-
"""The order basis' numbers for the monthly results file (core/results, schema 1),
taken from the same store dict S and the same row builders that fill the workbook:
② from sheet_pl (build_pl_rows, pl_value, PL_FORMULAS, pl_memo), ④ from
sheet_fees.fee_groups, ⑤ from sheet_returns.status_groups, ⑦ from
sheet_sku.sku_final_rows (表五), ⑨ from sheet_inventory.stock_by_sku.

Every amount goes through catalog.money, as the sheets write it, and every label
through site.localize (Peru: IVA → IGV …), as localize_workbook does, so a results
file and its workbook show the same number under the same name."""
import pandas as pd

from core import results
from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.site import localize
from platforms.mercadolibre.report.catalog import FEE_BY_ES, REVERSAL_TO_PARENT, STATUS_ZH, classify, money
from platforms.mercadolibre.report.sheet_fees import fee_groups
from platforms.mercadolibre.report.sheet_inventory import stock_by_sku
from platforms.mercadolibre.report.sheet_pl import (PL_FORMULA_KEYS, PL_FORMULAS, build_pl_rows, pl_memo,
                                                    pl_value)
from platforms.mercadolibre.report.sheet_returns import status_groups
from platforms.mercadolibre.report.sheet_sku import sku_final_rows

# ② row of an order-level fee bucket (catalog.FEE_CATALOG)
ORDER_BUCKET_ROW = {"commission": "com_neg", "shipping": "shipcost_neg", "returns": "dev_neg"}


def pl(S):
    formulas = dict(PL_FORMULAS)
    spec = []
    for key, label, es, kind, _note in build_pl_rows():
        spec.append({"key": key or PL_FORMULA_KEYS[label], "name": label, "label": localize(label),
                     "original": localize(es or ""),
                     "kind": kind, "formula": formulas.get(label),
                     "value": money(pl_value(S, key)) if kind in ("d", "in") else None})
    return results.pl_rows(spec)


def memo(S, pl_doc=None):
    """备查项 amounts. The sensitivity row is the sheet's formula: 经营贡献毛利 − the
    first memo row (代扣代缴税金)."""
    pl_doc = pl_doc or pl(S)
    contrib = dict((r["key"], r["value"]) for r in pl_doc)["contrib"]
    out, first = [], None
    for key, (label, _es, fn, special, _note) in pl_memo([S]):
        if special == "sens":
            v = contrib - (first or 0.0)
        else:
            v = money(fn(S))
        if first is None:
            first = v
        out.append({"key": key, "label": localize(label), "value": v})
    return out


def fee_row(detalle):
    """The ② row (results "pl" key) a billing Detalle goes to; None when unregistered."""
    _zh, bucket, order_level, is_rev, known = classify(detalle)
    if not known:
        return None
    if order_level:
        return ORDER_BUCKET_ROW.get(bucket)
    parent = REVERSAL_TO_PARENT.get(detalle, detalle) if is_rev else detalle
    return "fee:" + parent if parent in FEE_BY_ES else None


def fees(S):
    """④ by Detalle (both settlement channels together), ② sign (a charge is
    negative). ④ is by charge date: it is not ②'s order-date fee rows."""
    g = fee_groups(S)
    out, at = [], {}
    for _, x in g.iterrows():
        d = x["detalle"]
        if d not in at:
            at[d] = len(out)
            out.append({"name": d, "label": localize(classify(d)[0]), "pl_key": fee_row(d), "amount": 0.0})
        out[at[d]]["amount"] -= money(x["sum"])
    return out


def returns(S):
    """⑤ by order status (售后/异常订单): count of orders and the refund in ② sign."""
    g = status_groups(S)
    if g is None:
        return []
    out = []
    for _, x in g.iterrows():
        zh = STATUS_ZH.get(x["estado"], (x["estado"],))[0]
        out.append({"key": x["estado"], "label": localize(zh), "count": int(x["单数"]),
                    "amount": -money(x["退款"])})
    return out


def sku(S):
    """⑦ 表五: units are the cost units (cancelled and resold returns already out,
    so units_returned is 0); fees = commission + withheld tax + shipping, ② sign;
    contribution before goods cost, as 表五's 净贡献毛利 without 采购成本."""
    cost = S.get("cost")
    out = []
    for a in sku_final_rows(S):
        gmv, ship, com, tax = money(a["GMV"]), money(a["运费收入"]), money(a["佣金"]), money(a["代扣税"])
        env, ref, cup = money(a["配送费"]), money(a["退款"]), money(a["券"])
        out.append({"sku": a["SKU"], "title": a["名称"], "units": float(a["件数"]), "units_returned": 0.0,
                    "net_sales": gmv - ref - cup, "fees": -(com + tax + env),
                    "contribution": gmv + ship - com - tax - env - ref - cup,
                    "cost": results.sku_cost(cost, a["SKU"])})
    return out


def inventory(S):
    """⑨ 表二's Full snapshot by SKU: stock = 库存合计, sold_30d = the platform's
    30-day sales. None without Full or without a stock file."""
    ST = S.get("stock")
    if not site.HAS_FULL or ST is None or not len(ST):
        return None
    agg = stock_by_sku(ST)
    agg = agg[agg["sku"].astype(str).str.strip() != ""]
    rows = [{"sku": str(x["sku"]), "title": "" if pd.isna(x["title"]) else str(x["title"]),
             "stock": float(x["total_u"]), "sold_30d": float(x["sales30_u"])} for _, x in agg.iterrows()]
    snap = S.get("stock_snap")
    return {"as_of": snap.strftime("%Y-%m-%d") if snap is not None and pd.notna(snap) else None, "rows": rows}


def figures(S):
    p = pl(S)
    return {"pl": p, "memo": memo(S, p), "fees": fees(S), "returns": returns(S), "sku": sku(S),
            "inventory": inventory(S)}
