# -*- coding: utf-8 -*-
"""What an accountant adds to the invoice basis, from files already downloaded.

Per reading (账单期 / 自然月切) of the invoice basis:

* actual refunds from MercadoPago next to the reversal-ratio estimate, per reversed
  sale, with the estimate's error;
* seller-funded coupons, dispute charges and withheld tax from MercadoPago;
* the accrual of charges on this reading's sales that a later invoice already in
  the folder carries;
* cost of goods when a SKU cost table is configured;
* the net-of-VAT figures (vat.py proves the charges include VAT).

Per month: the three dates of every billed sale (order / billing / delivery) and
the revenue split under each rule; the cash bridge from the adjusted result to the
money MercadoPago released; the materiality summary for finance.

Every function returns plain data; workbook.py lays it out. Nothing raises on a
missing MercadoPago file — each number says where it came from."""

import glob
import os
import re

import pandas as pd

from core import sku_costs as SC
from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (CH_NO, FEE_BY_ES, PAGOS_HEADER_ROW,
    SPANISH_MONTHS, V_SKU, fmt)

from platforms.mercadolibre.report.loading import _srcid

from experiments.ml_invoice_basis import mercadopago as MP
from experiments.ml_invoice_basis.common import key
from experiments.ml_invoice_basis.compare import CAT, month_orders, order_key
from experiments.ml_invoice_basis.invoice import billed_sales

ST_ACTUAL = "MercadoPago 实际退款"
ST_UNPAID = "收款未入账（下单后未付款或随即取消）：按全额冲销"
ST_NOPAY = "MercadoPago 没有这笔收款（下单后未付款或随即取消）：按全额冲销"
ST_OUT = "MercadoPago 文件未覆盖（早于流水起始日）：沿用估算"
ST_NOFILE = MP.NO_FILES + "：沿用估算"
ST_NOSALE = "冲销行没有销售号，无法关联收款：沿用估算"
STATUS_ORDER = (ST_ACTUAL, ST_UNPAID, ST_NOPAY, ST_OUT, ST_NOFILE, ST_NOSALE)
DATE_RULES = (("order", "按下单日"), ("billing", "按计费日（账单）"), ("delivery", "按交付日"))
NO_DATE = "未交付 / 无日期"

# Default materiality (the CLI can change both).
DEFAULT_PCT = 1.0          # % of the current basis' net revenue
DEFAULT_ABS = 1000.0       # in the store's currency


# ════════════════════════════════════════════════════════════════════════
# links between billed sales and MercadoPago payments
# ════════════════════════════════════════════════════════════════════════

def links(B):
    s2p, p2s, tot = MP.sale_payments(B)
    return {"s2p": s2p, "p2s": p2s, "tot": tot}


def _found(ks, lk, pay):
    return [p for p in lk["s2p"].get(ks, ()) if len(pay) and p in pay.index]


def _wsum(ks, lk, pay, col, pays=None):
    pays = _found(ks, lk, pay) if pays is None else pays
    return float(sum(MP.weight(ks, p, lk["p2s"], lk["tot"]) * float(pay.at[p, col]) for p in pays))


# ════════════════════════════════════════════════════════════════════════
# actual refunds next to the estimate
# ════════════════════════════════════════════════════════════════════════

def actual_reversals(inv, B, mp, lk):
    """One row per reversed sale of the reading: estimate, actual, error, status.

    A sale reversed in several invoices gets its actual refund in proportion to
    the commission reversed in this reading (reversed commission here ÷ all
    reversed commission of the sale in the downloaded invoices)."""
    R = inv["reversals"]
    pay, tol = mp["pay"], site.CN_TOL
    com = B[B["is_rev"] & (B["bucket"] == "commission") & (B["ksale"] != "")]
    rev_all = (-com["amount"]).groupby(com["ksale"]).sum().to_dict()
    span0 = mp["span"][0] if mp.get("span") else None
    out = []
    if not len(R):
        return pd.DataFrame(columns=["ksale", "kpack", "sdate", "cdate", "sale_total", "ratio", "est",
                                     "est_units", "actual", "dispute", "coupon_back", "err",
                                     "status", "pays", "title", "alloc"])
    R = R.assign(_g=[k if k else "row:%s" % i for k, i in zip(R["ksale"], R["idx"])])
    for g, x in R.groupby("_g", sort=False):
        ks = x["ksale"].iloc[0]
        revc = float(x["rev_commission"].sum())
        total = float(x["sale_total"].iloc[0])
        alloc = min(1.0, revc / rev_all[ks]) if ks and rev_all.get(ks, 0) > 0 else 1.0
        est = float(x["est_sales"].sum())
        sdate = x["sdate"].dropna().iloc[0] if x["sdate"].notna().any() else pd.NaT
        actual = dispute = cback = None
        pays = _found(ks, lk, pay) if ks else []
        if not ks:
            status = ST_NOSALE
        elif mp["source"] is None:
            status = ST_NOFILE
        elif pays:
            settled = sum(float(pay.at[p, "settle"]) for p in pays)
            refund = alloc * _wsum(ks, lk, pay, "refund", pays)
            dispute = alloc * _wsum(ks, lk, pay, "dispute", pays)
            cback = alloc * _wsum(ks, lk, pay, "coupon_back", pays)
            if settled <= tol and refund + dispute <= tol:
                status, actual, dispute, cback = ST_UNPAID, alloc * total, 0.0, 0.0
            else:
                status, actual = ST_ACTUAL, refund + dispute
        elif span0 is not None and pd.notna(sdate) and sdate.normalize() >= span0:
            status, actual, dispute, cback = ST_NOPAY, alloc * total, 0.0, 0.0
        else:
            status = ST_OUT
        out.append({"ksale": ks, "kpack": x["kpack"].iloc[0], "sdate": sdate,
                    "cdate": x["cdate"].max(), "sale_total": total,
                    "ratio": float(x["ratio"].clip(upper=1.0).sum()) if len(x) > 1
                    else float(x["ratio"].iloc[0]),
                    "est": est, "est_units": float(x["est_units"].sum()), "actual": actual,
                    "dispute": dispute, "coupon_back": cback,
                    "err": (est - actual) if actual is not None else None, "status": status,
                    "pays": ", ".join(sorted(lk["s2p"].get(ks, ()))) if ks else "",
                    "title": x["title"].iloc[0], "alloc": alloc})
    out = pd.DataFrame(out)
    for c in ("actual", "dispute", "coupon_back", "err"):      # None -> NaN, numeric dtype
        out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


def reversal_totals(A):
    known = A[A["actual"].notna()] if len(A) else A
    unk = A[A["actual"].isna()] if len(A) else A
    tol = site.CN_TOL
    t = {"n": int(len(A)), "est": float(A["est"].sum()) if len(A) else 0.0,
         "n_known": int(len(known)), "actual_known": float(known["actual"].sum()) if len(known) else 0.0,
         "est_known": float(known["est"].sum()) if len(known) else 0.0,
         "n_unknown": int(len(unk)), "est_unknown": float(unk["est"].sum()) if len(unk) else 0.0,
         "dispute": float(known["dispute"].fillna(0).sum()) if len(known) else 0.0,
         "coupon_back": float(known["coupon_back"].fillna(0).sum()) if len(known) else 0.0,
         "n_exact": int((known["err"].abs() <= tol).sum()) if len(known) else 0,
         "status": {s: int((A["status"] == s).sum()) if len(A) else 0 for s in STATUS_ORDER}}
    t["value"] = t["actual_known"] + t["est_unknown"]
    t["err"] = t["est_known"] - t["actual_known"]
    t["abs_err"] = float(known["err"].abs().sum()) if len(known) else 0.0
    return t


# ════════════════════════════════════════════════════════════════════════
# coupons, disputes, withheld tax (MercadoPago)
# ════════════════════════════════════════════════════════════════════════

def dev_fees(B):
    """{payment: [amounts of 'Cargo por devolución' rows and their reversals]}."""
    out = {}
    if "Pago" not in B.columns:
        return out
    d = B[B["parent_detalle"].astype(str) == "Cargo por devolución"]
    for pg, a in zip(d["Pago"].map(_srcid), d["amount"]):
        if pg:
            out.setdefault(pg, []).append(float(a))
    return out


def mp_on_sales(inv, A, mp, lk, after=None, B=None):
    """Coupons on the reading's billed sales (and given back on its reversed sales),
    dispute charges and withheld tax on their payments."""
    pay = mp["pay"]
    sales = [k for k in inv["sales"]["ksale"] if k] if len(inv["sales"]) else []
    gross, n_cpn, n_miss, n_unk, tax = 0.0, 0, 0, 0, 0.0
    for ks in sales:
        pays = _found(ks, lk, pay)
        if not pays:
            n_miss += 1
            continue
        if not all(bool(pay.at[p, "coupon_known"]) for p in pays):
            n_unk += 1
        c = _wsum(ks, lk, pay, "coupon", pays)
        gross += c
        n_cpn += int(abs(c) > site.COUPON_TOL)
        tax += -_wsum(ks, lk, pay, "tax_all", pays)
    back = float(A["coupon_back"].fillna(0).sum()) if len(A) else 0.0
    both = set(sales) | (set(A["ksale"]) - {""} if len(A) else set())
    seen, dch, n_dch, n_dev = set(), 0.0, 0, 0
    dev = dev_fees(B) if B is not None else {}
    for ks in both:
        for p in _found(ks, lk, pay):
            if p not in seen:
                seen.add(p)
                v = float(pay.at[p, "dispute_charge"])
                if abs(v) > site.ZERO_TOL:
                    dch += v
                    n_dch += 1
                    n_dev += int(any(abs(abs(v) - abs(a)) <= site.CN_TOL for a in dev.get(p, ())))
    claims = 0
    if after is not None and len(after):
        op = after[(after["flow"] == "claim") & (after["status"] == "opened")]
        ids = set(sales)
        pk = set(inv["sales"]["kpack"]) - {""} if len(inv["sales"]) else set()
        claims = int(op["order"].isin(ids | pk).sum())
    return {"available": mp["source"] is not None, "source": mp["source"],
            "coupon_available": mp["source"] == "settlement",
            "coupon_gross": gross, "coupon_back": back, "coupon_net": gross - back,
            "n_coupon": n_cpn, "n_sales": len(sales), "n_missing": n_miss, "n_unknown": n_unk,
            "dispute_charge": dch, "n_dispute_charge": n_dch, "n_dispute_is_dev": n_dev,
            "tax": tax, "open_claims": claims}


def coupon_check(D, mp, lk):
    """Ventas' derived coupon (the current report) vs the settlement residual, order
    by order, for the month's orders billed in any downloaded invoice."""
    pay = mp["pay"]
    if mp["source"] != "settlement" or not len(pay):
        return None
    B = D["ctx"]["B"]
    by_order = {}
    com = B[(~B["is_rev"]) & (B["bucket"] == "commission") & (B["ksale"] != "")]
    for ks, kp in zip(com["ksale"], com["kpack"]):
        ok = order_key(D["orders"], D["parent_of"], kp, ks)
        if ok:
            by_order.setdefault(ok, set()).add(ks)
    n = eq = 0
    tv = tm = 0.0
    for k, o in month_orders(D["orders"], D["month"]).items():
        sales = by_order.get(k, set())
        pays = set()
        for ks in sales:
            pays |= set(_found(ks, lk, pay))
        if not pays:
            continue
        m = sum(float(pay.at[p, "coupon"]) for p in pays)
        n += 1
        tv += o["coupon"]
        tm += m
        eq += int(abs(m - o["coupon"]) <= site.CN_TOL)
    return {"n": n, "eq": eq, "ventas": tv, "mp": tm}


# ════════════════════════════════════════════════════════════════════════
# accrual: later invoices already in the folder
# ════════════════════════════════════════════════════════════════════════

def accrual(inv, B, docs):
    """Order-level charges (not reversals) on this reading's billed sales that an
    invoice later than the reading already carries."""
    view = inv["view"]
    ks = set(inv["sales"]["ksale"]) - {""} if len(inv["sales"]) else set()
    rows = B[B["ksale"].isin(ks) & (~B["is_rev"]) & B["order_level"]
             & (~B.index.isin(view["rows"].index)) & (B["cdate"] > view["end"])]
    is_open = dict((d["name"], d["open"]) for d in docs)
    lines = []
    for (pdet, f), g in rows.groupby([rows["parent_detalle"].astype(str), rows["_file"]]):
        lines.append({"es": pdet, "zh": FEE_BY_ES.get(pdet, ("⚠ 未登记科目",))[0], "file": f,
                      "open": bool(is_open.get(f)), "n": int(len(g)), "amount": float(g["amount"].sum()),
                      "from": g["cdate"].min(), "to": g["cdate"].max()})
    lines.sort(key=lambda x: (x["file"], -abs(x["amount"])))
    return {"lines": lines, "total": float(rows["amount"].sum()), "n": int(len(rows)),
            "n_sales": int(rows["ksale"].nunique()), "rows": rows}


# ════════════════════════════════════════════════════════════════════════
# three dates per billed sale
# ════════════════════════════════════════════════════════════════════════

DELIVERED_COL = "Fecha entregado"
_DAY_RE = re.compile(r"(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)")


def parse_delivery(v, ref):
    """'13 de octubre | 11:40' (no year) -> that day in the year closest to the
    order date (an order of 30 Dec delivered '3 de enero' is next year's)."""
    if not isinstance(v, str):
        return pd.NaT
    m = _DAY_RE.search(v)
    if not m or m.group(2).lower() not in SPANISH_MONTHS or ref is None or pd.isna(ref):
        return pd.NaT
    d, mo = int(m.group(1)), SPANISH_MONTHS[m.group(2).lower()]
    cands = []
    for y in (ref.year - 1, ref.year, ref.year + 1):
        try:
            cands.append(pd.Timestamp(y, mo, d))
        except ValueError:
            pass
    if not cands:
        return pd.NaT
    return min(cands, key=lambda t: abs((t - ref.normalize()).days))


def delivery_map(V):
    """{Ventas id: delivery date}; a pack's items inherit the parent row's date."""
    out = {}
    if DELIVERED_COL not in V.columns:
        return out, False
    for oid, v, ref in zip(V["oid"], V[DELIVERED_COL], V["fecha"]):
        k = key(oid)
        d = parse_delivery(v, ref)
        if k and pd.notna(d):
            out[k] = d
    return out, True


def three_dates(D):
    """Every billed sale of the downloaded invoices: order date (Ventas, else the
    invoice's sale date), billing date (commission charge), delivery date (Ventas);
    the month's revenue under each rule; the sales touching the month."""
    B = D["ctx"]["B"]
    S = billed_sales(B)
    month = D["month"]
    dm, has_col = delivery_map(D["S"]["ventas_all"])
    orders, parent_of = D["orders"], D["parent_of"]
    rows = []
    for _, s in S.iterrows():
        ok = order_key(orders, parent_of, s["kpack"], s["ksale"])
        o = orders.get(ok) if ok else None
        od = o["date"] if o is not None and pd.notna(o["date"]) else s["sdate"]
        dd = dm.get(s["ksale"]) or (dm.get(ok) if ok else None)
        rows.append({"ksale": s["ksale"], "okey": ok, "order": od, "billing": s["cdate"],
                     "delivery": dd if dd is not None else pd.NaT, "total": s["total"],
                     "units": s["units"], "status": o["status"] if o is not None else "Ventas 查无",
                     "title": s["title"], "from_ventas": o is not None})
    T = pd.DataFrame(rows, columns=["ksale", "okey", "order", "billing", "delivery", "total", "units",
                                    "status", "title", "from_ventas"])
    ym = lambda d: d.strftime("%Y-%m") if d is not None and pd.notna(d) else NO_DATE
    for r, _lab in DATE_RULES:
        T["m_" + r] = T[r].map(ym) if len(T) else []
    months = sorted(set().union(*[set(T["m_" + r]) for r, _l in DATE_RULES])) if len(T) else []
    months = [m for m in months if m != NO_DATE] + ([NO_DATE] if NO_DATE in months else [])
    split = []
    for mo in months:
        d = {"month": mo}
        for r, _lab in DATE_RULES:
            x = T[T["m_" + r] == mo]
            d[r] = float(x["total"].sum())
            d["n_" + r] = int(len(x))
        split.append(d)
    touch = T[(T["m_order"] == month) | (T["m_billing"] == month) | (T["m_delivery"] == month)] \
        if len(T) else T
    in_month = dict((r, float(T[T["m_" + r] == month]["total"].sum()) if len(T) else 0.0)
                    for r, _l in DATE_RULES)
    return {"table": touch.sort_values(["order", "ksale"]) if len(touch) else touch,
            "split": split, "in_month": in_month, "has_delivery": has_col,
            "n_all": int(len(T)), "n_no_delivery": int(T["delivery"].isna().sum()) if len(T) else 0}


# ════════════════════════════════════════════════════════════════════════
# cost of goods (only when a SKU cost table is configured)
# ════════════════════════════════════════════════════════════════════════

def load_cost_table():
    """(table or None, note). load_configured() reads only reports.sku_costs of
    config.json; any problem reading it is a note, never a failure."""
    try:
        return SC.load_configured(), ""
    except Exception as e:                      # StoreConfigError and friends
        return None, "成本表读不出来：%s" % str(e)[:120]


def cogs(inv, A, D, table):
    """Units of the reading's billed sales minus the units of fully reversed sales,
    per SKU, priced with the cost table."""
    if table is None:
        return None
    V = D["S"]["ventas_all"]
    sku_of = {}
    if V_SKU in V.columns:
        for oid, s in zip(V["oid"], V[V_SKU]):
            k = key(s)
            if key(oid) and k:
                sku_of[key(oid)] = k
    orders, parent_of = D["orders"], D["parent_of"]

    def sku(ks, kp):
        if ks in sku_of:
            return sku_of[ks]
        ok = order_key(orders, parent_of, kp, ks)
        o = orders.get(ok) if ok else None
        return o["skus"][0] if o is not None and len(o["skus"]) == 1 else ""
    agg = {}
    for _, s in inv["sales"].iterrows():
        k = sku(s["ksale"], s["kpack"]) or "（无 SKU）"
        a = agg.setdefault(k, {"name": s["title"], "units": 0.0, "ret": 0.0, "sales": 0.0})
        a["units"] += float(s["units"])
        a["sales"] += float(s["total"])
    if len(A):
        for _, r in A[A["ratio"] >= 0.999].iterrows():
            k = sku(r["ksale"], r["kpack"]) or "（无 SKU）"
            a = agg.setdefault(k, {"name": r["title"], "units": 0.0, "ret": 0.0, "sales": 0.0})
            a["ret"] += float(r["est_units"])
    items = [SC.Item(k, a["name"], int(round(a["units"])), int(round(a["ret"])), a["sales"])
             for k, a in sorted(agg.items())]
    res = SC.apply(table, "ML", site.SITE.upper(), site.CUR, D["month"], items)
    return {"res": res, "cogs": res.cogs, "coverage": res.coverage_text(),
            "units": sum(i.units for i in items), "returned": sum(i.returned for i in items)}


# ════════════════════════════════════════════════════════════════════════
# per reading: everything together
# ════════════════════════════════════════════════════════════════════════

def reading(inv, D, mp, lk, table):
    B = D["ctx"]["B"]
    A = actual_reversals(inv, B, mp, lk)
    t = reversal_totals(A)
    m = mp_on_sales(inv, A, mp, lk, mp.get("after"), B)
    acc = accrual(inv, B, D["docs"])
    n = inv["n"]
    coupon = m["coupon_net"] if m["coupon_available"] else 0.0
    # MercadoPago's zero-amount DISPUTE deductions are the invoice's own return fee
    # (Cargo por devolución, same payment, same amount): already in the fees, not added.
    adj = t["est"] - t["value"] - coupon - acc["total"]
    contrib_adj = n["contribution"] + adj
    vat = D["vat"]
    r = vat["rate"]
    rev_gross = (n["gross_sales"] - t["value"] + n["ship_income"] - n["ship_income_rev"] - coupon)
    fees_gross = n["fee_total"] + acc["total"]
    rev_net = rev_gross / (1 + r)
    fees_net = fees_gross / (1 + r) if vat["credit"] else fees_gross
    c = cogs(inv, A, D, table)
    return {"A": A, "rev": t, "mp": m, "accrual": acc, "coupon": coupon,
            "contrib_adj": contrib_adj,
            "vat": {"rev_gross": rev_gross, "rev_net": rev_net, "output": rev_gross - rev_net,
                    "fees_gross": fees_gross, "fees_net": fees_net, "input": fees_gross - fees_net,
                    "contrib_net": rev_net - fees_net},
            "cogs": c}


# ════════════════════════════════════════════════════════════════════════
# cash: from the adjusted result to the money MercadoPago released
# ════════════════════════════════════════════════════════════════════════

def pagos_in_month(folder, p0, p1):
    """Invoice payments (not credit notes) made in the month, from every
    Reporte_Pagos_Facturas file of the folder."""
    out, seen = [], set()
    for p in sorted(glob.glob(os.path.join(folder, "Reporte_Pagos_Facturas_*.xlsx"))):
        if os.path.basename(p).startswith("~$"):
            continue
        try:
            P = pd.read_excel(p, sheet_name="Pagos y notas de crédito",
                              header=PAGOS_HEADER_ROW).dropna(how="all")
        except Exception:
            continue
        if not {"Tipo de pago", "Importe total"} <= set(P.columns):
            continue
        tp = P["Tipo de pago"].astype(str)
        paid = P[~tp.str.contains("Nota", na=False) & tp.str.strip().ne("") & tp.ne("nan")]
        dt = pd.to_datetime(paid.get("Fecha de pago / Emisión de NC"), errors="coerce")
        for no, d, amt, kind in zip(paid.get("Número de pago", pd.Series("", index=paid.index)), dt,
                                    pd.to_numeric(paid["Importe total"], errors="coerce").fillna(0.0),
                                    paid["Tipo de pago"]):
            k = (key(no), d)
            if k in seen or pd.isna(d) or not (p0 <= d <= p1):
                continue
            seen.add(k)
            out.append({"file": os.path.basename(p), "no": key(no), "day": d, "amount": float(amt),
                        "kind": str(kind)})
    return out


def _classify(R, pagos_all):
    sale = R["type"].isin(MP.SALE_TYPES)
    out = pd.Series("other", index=R.index)
    out[(R["type"] == "SETTLEMENT") & (R["trans"] < 0)] = "outgoing"
    out[sale & R["sid"].isin(pagos_all) & (out != "outgoing")] = "sale"
    ml = R["unit"].astype(str).str.contains("Mercado Libre", case=False, na=False)
    out[sale & ~R["sid"].isin(pagos_all) & ml & (out != "outgoing")] = "sale_other"
    return out


def cash_bridge(D, cal, CR):
    """Lines (dicts) of the bridge for the calendar month; see workbook.cash_sheet."""
    mp, lk = D["mp"], D["links"]
    p0, p1 = D["S"]["p0"], D["S"]["p1"]
    inv = D["invs"]["calendar"]
    rows = inv["view"]["rows"]
    no_rows = rows[(rows["channel"] == CH_NO) | (rows["is_rev"] & (rows["parent_channel"] == CH_NO))]
    L = []

    def add(k, label, src, v, kind="m", note="", sum_of=None):
        L.append({"key": k, "label": label, "src": src, "value": v, "kind": kind, "note": note,
                  "sum_of": sum_of})
    add("c1", "经营贡献毛利（账单口径 · 自然月切 · 调整后，含税）", "③ 损益表", CR["contrib_adj"],
        note="已含实际退款、优惠券和应计费用")
    add("c2", "加回：另行付款的账单级费用（不从货款里扣）", "账单 · 通道 No",
        float(no_rows["amount"].sum()),
        note="广告、仓储等“月底账单支付”的费用不从每笔货款里扣，平台另开账单向卖家收（见下面的 Pagos_Facturas）")
    add("c3", "加回：应计费用（下期账单才扣）", "下期账单", CR["accrual"]["total"],
        note="已在损益里计入、但还没扣款")
    have = mp["source"] == "settlement"
    why_none = MP.NO_FILES if not any(mp["files"].values()) else \
        "本次下载没有 settlement_v2（%s）" % MP.source_text(mp)
    pay = mp["pay"]
    sales = (set(inv["sales"]["ksale"]) if len(inv["sales"]) else set()) | \
        (set(CR["A"]["ksale"]) if len(CR["A"]) else set())
    sales.discard("")
    P = set()
    for ks in sales:
        P |= set(_found(ks, lk, pay))
    if not D["vat"]["withholding"]:
        tax_v, tax_note = 0.0, "本国平台不代扣税"
    elif have:
        tax_v = float(sum(pay.at[p, "tax_all"] for p in P))
        tax_note = "可抵税款，不是费用；正式凭证是 CFDI de retenciones（未下载）"
    else:
        tax_v, tax_note = None, why_none
    add("c4", "减：平台代扣税（这些收款上 MercadoPago 扣下的 TAXES_AMOUNT，扣除退款退回的）",
        "settlement_v2", tax_v, note=tax_note)
    add("c5", "＝ 本期开票销售应收货款净额（推算）", "", None, "sub", sum_of=["c1", "c2", "c3", "c4"])
    if not have:
        add("c6", "实际：这些销售的收款在 MercadoPago 的净额", "settlement_v2", None, note=why_none)
        add("c7", "差额（未解释）", "", None, "tot", note=why_none)
        L.append({"key": "stop", "label": why_none, "kind": "note"})
        return {"lines": L, "have": False, "pagos": pagos_in_month(D["folder"], p0, p1),
                "withdraw": _withdraw(mp, p0, p1), "release_check": _release_report(mp, p0, p1)}
    Rw = mp["rows"]
    cls = _classify(Rw, set(lk["p2s"]))
    inM = (Rw["tdate"] >= p0) & (Rw["tdate"] <= p1)
    mine = Rw["sid"].isin(P) & cls.isin(["sale"])
    add("c6", "实际：这些销售的收款在 MercadoPago 的净额（全部日期，REAL_AMOUNT）", "settlement_v2",
        float(Rw[mine]["real"].sum()),
        note="本期开票或冲销的销售的收款号（账单的 Pago 列）在结算流水里的全部收支：收款、退款、争议")
    add("c7", "差额：实际 − 推算", "", None, "tot", sum_of=["c6", "-c5"],
        note="下面逐项拆开；每一项都是两份数据的实测差，不是倒挤")
    parts = cash_parts(D, CR, inv, P, Rw[mine])
    for k, label, v, note in parts:
        add(k, label, "settlement_v2 ↔ 账单", v, "part", note=note)
    add("c7z", "余数（拆解后仍未解释）", "", None, "chk",
        sum_of=["c7"] + ["-" + k for k, _l, _v, _n in parts],
        note="拆解是逐笔做的；余数只来自四舍五入。"
             "上面各项里“收款金额 ≠ 开票售价”一项本身还要逐笔查（见说明）")
    add("c8", "去掉：上述收款中交易日不在本月的部分（取反）", "settlement_v2",
        -float(Rw[mine & ~inM]["real"].sum()),
        note="本月开票的销售上月就收款（取反为负），或退款在下月（取反为正）")
    add("c9", "加：本月交易日、不属于本期开票销售的销售收支", "settlement_v2",
        float(Rw[inM & cls.isin(["sale", "sale_other"]) & ~mine]["real"].sum()),
        note="别的月份开票（或账单未下载）的销售在本月收款/退款")
    add("c10", "加：本月其他 MercadoPago 收支（运费、返现、扣收欠款等，不挂销售收款号）", "settlement_v2",
        float(Rw[inM & (cls == "other")]["real"].sum()))
    add("c11", "＝ 本月 MercadoPago 销售相关净流入（按交易日）", "", None, "sub",
        sum_of=["c6", "c8", "c9", "c10"])
    add("c11k", "核对：结算流水本月合计（不含卖家对外付款）", "settlement_v2",
        float(Rw[inM & (cls != "outgoing")]["real"].sum()), "info")
    add("c11r", "残差（应为 0）", "", None, "chk", sum_of=["c11", "-c11k"],
        note="分类是穷尽的，所以这里必须为 0")
    nr = Rw[cls != "outgoing"]
    later = inM & ~((Rw["rdate"] >= p0) & (Rw["rdate"] <= p1))
    earlier = (Rw["tdate"] < p0) & (Rw["rdate"] >= p0) & (Rw["rdate"] <= p1)
    add("c12", "减：本月交易、下月以后才释放（MONEY_RELEASE_DATE 在月末之后）", "settlement_v2",
        -float(Rw[later & (cls != "outgoing")]["real"].sum()))
    add("c13", "加：以前交易、本月释放", "settlement_v2", float(Rw[earlier & (cls != "outgoing")]["real"].sum()))
    add("c14", "＝ 本月释放给卖家的货款（结算流水，按释放日）", "", None, "tot", sum_of=["c11", "c12", "c13"])
    rel = _release_report(mp, p0, p1, set(Rw[cls != "outgoing"]["sid"]) - {""})
    add("c14k", "核对：结算流水按释放日直接加总", "settlement_v2",
        float(nr[(nr["rdate"] >= p0) & (nr["rdate"] <= p1)]["real"].sum()), "info")
    add("c14r", "残差（应为 0）", "", None, "chk", sum_of=["c14", "-c14k"])
    if rel is not None:
        add("c15", "对照：Liberaciones 报表本月释放的货款（同一批收款号）", "reserve-release", rel["net"],
            "info", note="同一批收款号（不含卖家对外付款）在 Liberaciones 里按 DATE 落在本月的行，不含提现")
        add("c15d", "两份报表的差（未解释）", "", None, "info", sum_of=["c15", "-c14"],
            note="逐笔收款两份报表的合计基本相等；差在日期：Liberaciones 把争议冻结/解冻（reserve_for_*）"
                 "记在各自的日子，流水只记一个释放日。差额不强行抹平")
    return {"lines": L, "have": True, "pagos": pagos_in_month(D["folder"], p0, p1),
            "withdraw": _withdraw(mp, p0, p1), "release_check": rel,
            "outgoing": float(Rw[inM & (cls == "outgoing")]["real"].sum())}


def cash_parts(D, CR, inv, P, r):
    """The difference actual − expected of the bridge, element by element.

    r: settlement rows (SETTLEMENT / REFUND / DISPUTE) of the payments P of this
    reading's billed and reversed sales. Withheld tax is taken from the same rows
    on both sides, so it is not a part."""
    lk = D["links"]
    n = inv["n"]
    billed = set(inv["sales"]["ksale"]) - {""} if len(inv["sales"]) else set()
    rev = set(CR["A"]["ksale"]) - {""} if len(CR["A"]) else set()
    g1 = {p for p in P if lk["p2s"].get(p, set()) & billed}
    st = r["type"] == "SETTLEMENT"
    rows = inv["view"]["rows"]
    ol = rows[rows["order_level"] & (rows["channel"] != CH_NO)
              & ~(rows["is_rev"] & (rows["parent_channel"] == CH_NO))]
    ours = ol["ksale"].isin(billed | rev)
    ship = n["ship_income"] - n["ship_income_rev"]
    g1t = r[st & r["sid"].isin(g1)].groupby("sid")["trans"].sum()
    tot = lk["tot"]
    eq = sum(1 for p, t in g1t.items()
             if abs(t - sum(tot.get(x, 0.0) for x in lk["p2s"].get(p, ()) if x in billed)) <= site.CN_TOL)
    return [
        ("k1", "收款金额 − （开票售价 + 买家运费），本期开票的销售",
         float(g1t.sum()) - n["gross_sales"] - ship,
         "%d 笔收款里 %d 笔收款金额 = 开票售价。不等的：买家运费有时在同一笔收款里、有时单独一笔"
         "（SETTLEMENT_SHIPPING，不挂销售收款号）；一笔销售分两次付款时账单只记其中一个收款号；"
         "未付款就取消的销售没有收款。这一项要逐笔查" % (len(g1t), eq)),
        ("k2", "冲销的销售在别期开票：它们当初的收款也在“实际”里",
         float(r[st & ~r["sid"].isin(g1)]["trans"].sum()),
         "本期冲销、但销售在上期账单里：收入不在本期损益，收款却在这批收款号的流水里"),
        ("k3", "退款：MercadoPago 实际退给买家的 − 损益里的实际退款",
         float(r[~st]["trans"].sum()) + CR["rev"]["value"],
         "多为本期开票的销售已经退款，但账单的冲销还没出（在下期账单或尚未冲销）；"
         "损益按冲销所在的期间算退款"),
        ("k4", "费用：MercadoPago 从货款里扣的 − 账单上这些销售的订单级费用（本期）",
         float(r["fee"].sum()) + float(ol[ours]["amount"].sum()),
         "配送费、退货处理费在另一期账单计费，平台退回的费用与账单冲销不在同一期。"
         "（MercadoPago 争议里金额为 0、只扣费的行就是账单的退货处理费，算在这一项）"),
        ("k5", "本期账单里、但销售不在本期开票或冲销的订单级费用", float(ol[~ours]["amount"].sum()),
         "损益扣了、但对应的收款不在这批收款号里"),
        ("k6", "优惠券：MercadoPago − 损益", -float(r["imp"].sum()) + CR["coupon"],
         "损益的退回优惠券只算本期冲销的销售"),
    ]


RELEASE_SKIP = ("", "payout", "reserve_for_payout", "reserve_for_payment")


def _release_report(mp, p0, p1, sids=None):
    """Liberaciones rows dated in the month (no payouts / payout reserves); with
    `sids`, only those payments (the same ones the settlement side counts)."""
    r = mp.get("release")
    if r is None or not len(r):
        return None
    x = r[(r["day"] >= p0) & (r["day"] <= p1) & ~r["desc"].isin(RELEASE_SKIP)]
    if sids is not None:
        x = x[x["sid"].isin(sids)]
    return {"net": float(x["net"].sum()), "n": int(len(x))}


def _withdraw(mp, p0, p1):
    w = mp.get("withdraw")
    if w is None or not len(w):
        return None
    x = w[(w["day"] >= p0) & (w["day"] <= p1) & (w["status"] == "approved")]
    return {"amount": float(x["amount"].sum()), "n": int(len(x))}


# ════════════════════════════════════════════════════════════════════════
# materiality summary (② for finance)
# ════════════════════════════════════════════════════════════════════════

REASON = {"time": "时间差", "other": "时间差", "cycle": "时间差", "rev": "估算与实际",
          "absent": "账单里没有", "source": "取数来源不同", "same": "取数来源不同",
          "na": "取数来源不同", "unit": "取数来源不同"}
REASON_VAT = "含税/不含税列报"

SHORT = {
    "gmv": "现行算本月下单的订单；账单算本月计费的销售。差额来自月初月末的订单（下单与计费不在同一月）、"
           "Ventas 不列的随即取消的订单，以及两边金额不同的订单",
    "refund": "现行用 Ventas 的退款净额、按下单月；账单按冲销的计费日、用冲销比例估算冲掉的销售额",
    "cn_adj": "现行报表修正 Ventas 对部分退款重复扣佣金；账单口径不经过 Ventas，不需要",
    "ship_in": "两边都是买家付的运费，差额来自期间（下单月 ↔ 计费月）",
    "coupon": "账单里没有卖家优惠券；下面“调整为实际”用 MercadoPago 的实际数补上",
    "commission": "现行取本月订单的佣金（不论哪张账单计费）；账单取本月计费的佣金",
    "shipping": "现行取 Ventas 的配送费；账单取本月计费的配送费",
    "returns": "两边按不同的月份归属退货处理费",
    "order_rev": "账单把取消/退款退回的佣金、配送费单列；现行把它们抵在退款净额里",
    "unc_order": "字典里没有的订单级费用",
    "unc_bill": "字典里没有的账单级费用",
}


def summary_rows(D):
    """Rows of ②: a walk from the current contribution to the invoice basis
    (calendar month), each step with its amount and reason category."""
    cal = D["acc"]["calendar"]
    out = []
    S = D["S"]
    out.append({"key": "start", "kind": "start", "label": "现行订单口径 · 经营贡献毛利（含税）",
                "value": S["contrib"], "cat": "", "why": "月度报表 ② 的数（同一个下载目录）"})
    for ln in D["lines"]:
        if ln["kind"] != "m":
            continue
        inv = ln["inv"].get("calendar")
        diff = (inv or 0.0) - (ln["cur"] or 0.0)
        if abs(diff) <= site.ZERO_TOL:
            continue
        cats = ln["cats"] if isinstance(ln["cats"], list) else []
        cat = REASON.get(cats[0], "取数来源不同") if cats else "取数来源不同"
        k = ln["key"].split(":")[0]
        why = SHORT.get(ln["key"]) or (("按计费日归入本月的同一批账单行；差额来自" + CAT[cats[0]])
                                       if k == "fee" and cats else CAT.get(cats[0], "") if cats else "")
        out.append({"key": ln["key"], "kind": "diff", "label": ln["label"], "value": diff,
                    "cur": ln["cur"], "inv": inv, "cat": cat, "why": why})
    out.append({"key": "mid", "kind": "sub", "label": "＝ 账单口径 · 经营贡献毛利（冲销按估算，含税）",
                "value": D["invs"]["calendar"]["n"]["contribution"], "cat": "",
                "why": "③ 损益表的同名行（自然月切）"})
    t = cal["rev"]
    have = cal["mp"]["available"]
    nof = MP.source_text(D["mp"])
    out.append({"key": "adj_rev", "kind": "diff", "label": "冲销：估算 → MercadoPago 实际",
                "value": t["est"] - t["value"], "cat": "估算与实际",
                "why": ("%d 笔冲销里 %d 笔有实际退款记录，估算比实际%s %s；其余 %d 笔沿用估算"
                        % (t["n"], t["n_known"], "多" if t["err"] >= 0 else "少", fmt(abs(t["err"]), D["dec"]),
                           t["n_unknown"])) if have else nof + "，沿用估算"})
    out.append({"key": "adj_coupon", "kind": "diff", "label": "减：卖家承担优惠券（MercadoPago 实际）",
                "value": -cal["coupon"], "cat": "账单里没有",
                "why": ("账单没有优惠券；取自 MercadoPago（%s），%d 笔销售有优惠券" % (nof, cal["mp"]["n_coupon"]))
                if cal["mp"]["coupon_available"] else nof + "：优惠券无法得到，按 0 计（不是没有优惠券）"})
    out.append({"key": "adj_accrual", "kind": "diff", "label": "减：应计费用（下期账单已出现）",
                "value": -cal["accrual"]["total"], "cat": "时间差",
                "why": "本月销售的费用、在下期账单里才计费（%d 行）；还没出账单的无法得知" % cal["accrual"]["n"]})
    out.append({"key": "adj", "kind": "sub", "label": "＝ 账单口径 · 经营贡献毛利（调整后，含税）",
                "value": cal["contrib_adj"], "cat": "", "why": "③ 损益表“调整后”行（自然月切）"})
    v = cal["vat"]
    out.append({"key": "vat", "kind": "diff", "label": "减：增值税列报（收入与费用都改为不含税）",
                "value": v["contrib_net"] - cal["contrib_adj"], "cat": REASON_VAT,
                "why": "收入中的销项税 %s 不是卖家的钱；费用中的进项税 %s %s" % (
                    fmt(v["output"], D["dec"]), fmt(v["input"], D["dec"]),
                    "可以抵扣" if D["vat"]["credit"] else "不能抵扣（按含税计费用）")})
    out.append({"key": "end", "kind": "end", "label": "＝ 账单口径 · 经营贡献毛利（调整后，不含税）",
                "value": v["contrib_net"], "cat": "", "why": "③ 损益表的不含税列"})
    return out


def base_net_revenue(D):
    """The current basis' net revenue (GMV − refunds + adjustment + shipping − coupons):
    the denominator of the % column on ②."""
    keys = ("gmv", "refund", "cn_adj", "ship_in", "coupon")
    return float(sum((ln["cur"] or 0.0) for ln in D["lines"] if ln["key"] in keys))


def flagged(rows, base, pct=DEFAULT_PCT, absv=DEFAULT_ABS):
    """The rows over the threshold (for the console; ② recomputes it in Excel)."""
    out = []
    for r in rows:
        if r["kind"] != "diff":
            continue
        if abs(r["value"]) >= absv or (base and abs(r["value"]) >= abs(base) * pct / 100.0):
            out.append(r)
    return out
