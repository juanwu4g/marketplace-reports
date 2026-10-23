# -*- coding: utf-8 -*-
"""The bill basis' numbers for the monthly results file (core/results, schema 1),
taken from the same S and the same row builders that fill the store workbook:
② from workbook.pl_rows + S["m"]["pl"], 备查项 from workbook.pl_memo, ④ from the
ledger rows of the period (S["P"]) by Detalle, ⑤ from S["revs"], ⑦ from S["sku"],
⑨ from workbook.inventory_rows."""
from core import report_family as RF
from core import results
from platforms.mercadolibre.bill_report import catalog as C
from platforms.mercadolibre.bill_report import mapping as M
from platforms.mercadolibre.bill_report import status as ST
from platforms.mercadolibre.bill_report import workbook as WB


def pl(S):
    cost = S.get("cost")
    spec = []
    for row in WB.pl_rows(S):
        if row.kind == "d":
            v = S["m"]["pl"][row.key]
        elif row.kind == "in":                   # 商品采购成本: the cost table's cogs, as the sheet writes it
            v = cost.cogs if cost is not None else 0.0
        else:
            v = None
        spec.append({"key": row.key, "label": row.label, "original": row.original, "kind": row.kind,
                     "value": v, "formula": row.formula})
    return results.pl_rows(spec)


def memo(S):
    fam = RF.Family([S])
    return [{"key": key, "label": label, "value": float(fn(S))} for key, (label, _o, fn, _n) in WB.pl_memo(fam, S)]


def fees(S):
    """④: every Detalle with rows charged in the month; amount in ② sign (a
    charge is negative), the ② row it goes to."""
    P_ = S["P"]
    out = []
    for d in dict.fromkeys(list(P_["detalle"])):
        rows = P_[P_["detalle"] == d]
        e, known = C.entry(d)
        out.append({"name": d, "label": e.cn if known else "⚠ 未登记科目", "pl_key": e.line,
                    "amount": -float(rows["amount"].sum())})
    return out


def returns(S):
    """⑤ 一: the month's reversals by class (② sign, an estimate as in ②)."""
    out = []
    for c in ST.REVERSAL_CLASSES:
        ev = [e for e in S["revs"] if e["in_period"] and e["cls"] == c]
        out.append({"key": c, "label": ST.CLASS_CN[c], "count": len(ev),
                    "amount": float(sum(e["amount"] for e in ev))})
    return out


def sku(S):
    """⑦ 表一, one entry per row (pseudo rows and unmapped sales included, as there)."""
    out = []
    cost = S.get("cost")
    for k, _vals in WB.sku_rows(S, None):
        x = S["sku"][k]
        real = not (k.startswith("（") or k.startswith(M.UNMAPPED_PREFIX))
        out.append({"sku": k, "title": S["sku_titles"].get(k, "") or "", "units": float(x["units"]),
                    "units_returned": float(x["units_rev"]), "net_sales": float(x["net_gmv"]),
                    "fees": float(x["direct"] + x["alloc"]), "contribution": float(x["contrib"]),
                    "cost": results.sku_cost(cost, k) if real else None})
    return out


def inventory(S):
    """⑨'s Full snapshot: stock = the four sellable/held states, sold_30d = the
    platform's last-30-days sales. None when the country has no Full or the
    folder has no stock file."""
    if not S["prof"].has_full or S.get("stock") is None:
        return None
    rows = []
    for vals, _flag in WB.inventory_rows(S):
        rows.append({"sku": vals[0], "title": vals[1] or "", "stock": float(sum(vals[2:6])),
                     "sold_30d": float(vals[7])})
    return {"as_of": str(S.get("stock_snapshot") or "") or None, "rows": rows}


def figures(S):
    return {"pl": pl(S), "memo": memo(S), "fees": fees(S), "returns": returns(S), "sku": sku(S),
            "inventory": inventory(S)}
