# -*- coding: utf-8 -*-
"""The invoice basis: everything computed from Facturación + Notas de crédito alone.

Two readings of "the month" (both from the same invoice rows):

* **invoice period** (账单期) — the documents named for the month
  (`Reporte_Facturacion_MercadoLibre_Sept2026.xlsx` + the Notas de crédito of the
  same name). Mexico and Peru: the calendar month. Chile: from the store's cut-off
  day of the previous month to the day before it in this month (DAB_CL: Aug 8 –
  Sep 7 for "Sept2026"); the cut-off day is inferred per store by the report's own
  infer_cycle_day().
* **calendar cut** (自然月切) — rows of all downloaded invoices whose charge date
  (Fecha del cargo) is in the calendar month. In Chile that takes Sep 1–7 from the
  "Sept" invoice and Sep 8–30 from the "Oct" invoice.

Revenue comes from the 'Cargo por venta' rows: one row per billed sale (Número de
venta) carrying the sale's total, units, unit price, sale date and listing.
Reversals ('Anulación del cargo por venta', in Facturación or Notas de crédito)
undo a share of a sale: the share is the reversed commission over the original
commission of that sale — MercadoLibre returns the commission in proportion to
what it refunds (verified against MercadoPago cash in the main report's
credit_note_audit). Buyer-paid shipping comes from 'Envío a cargo del cliente' on
the 'Cargo por envíos' rows (equal to Ventas' "Ingresos por envío" order by order
on the real data: DAB_CL 592/592, BOCINA_TA02 13/13).
"""

import glob
import os

import pandas as pd

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (BILLING_HEADER_ROW, FEE_BY_ES, FEE_CATALOG,
    F_CHANNEL, F_DETALLE, F_STATE, MONTH_FULL_ES, file_month_token, month_bounds)
from platforms.mercadolibre.report.loading import (BILL_OPEN_RE, bill_cycles, infer_cycle_day,
    load_billing)

from experiments.ml_invoice_basis.common import FAC_GLOB, NC_GLOB, ExperimentError, key

SALE = "Cargo por venta"
SHIP = "Cargo por envíos de Mercado Libre"
C_TOTAL = "Total de la venta"
C_QTY = "Cantidad vendida"
C_PRICE = "Precio unitario"
C_TITLE = "Título de publicación"
C_PUB = "Número de publicación"
C_BUYER_SHIP = "Envío a cargo del cliente"
C_PCT = "Porcentaje por categoría"
C_UNIT_FEE = "Costo por unidad vendida"

LINK_ID = "按冲销号关联（Cargo que bonifica → Número del cargo）"
LINK_SALE = "按销售号关联（Número de venta，原佣金在已下载的任一账单里）"
LINK_RATE = ("原佣金不在已下载的账单里：按同一商品（Número de publicación）的类目费率重算原佣金"
             "（与冲销额恰好相等即全额冲销）")
LINK_NONE = "无法关联：按全额冲销计"
LINK_ORDER = (LINK_ID, LINK_SALE, LINK_RATE, LINK_NONE)


def bill_name(ym):
    return "%s %d" % (MONTH_FULL_ES[ym[1]], ym[0]) if ym else "（文件名认不出月份）"


def _num(df, col):
    if col in df.columns:
        return pd.to_numeric(df[col], errors="coerce")
    return pd.Series(float("nan"), index=df.index)


def documents(folder):
    """Every Facturación / Notas de crédito file: name, kind, which bill, still open?"""
    out = []
    for pat, doc in ((FAC_GLOB, "FAC"), (NC_GLOB, "NC")):
        for p in sorted(glob.glob(os.path.join(folder, pat))):
            name = os.path.basename(p)
            if name.startswith("~$"):
                continue
            is_open = False
            try:
                top = pd.read_excel(p, sheet_name="REPORT", header=None, nrows=BILLING_HEADER_ROW)
                is_open = any(isinstance(v, str) and BILL_OPEN_RE.search(v)
                              for v in top.values.ravel())
            except Exception:
                pass
            out.append({"path": p, "name": name, "doc": doc, "bill": file_month_token(name),
                        "open": is_open})
    return out


def load(folder, month, B=None):
    """All invoice rows of a folder, with the keys this experiment needs.

    B may be the frame derive() already loaded (S["billing_all"]) so that row
    labels are shared with the current basis; it is copied, never changed."""
    y, m = int(month[:4]), int(month[5:7])
    if B is None:
        try:
            B = load_billing(folder)
        except IOError:
            raise ExperimentError("下载目录里没有任何 Facturación 账单文件：%s" % folder)
    B = B.copy()
    B["_bill"] = B["_file"].map(lambda f: file_month_token(str(f)))
    for src, dst in (("k_sale", "ksale"), ("k_pack", "kpack"), ("charge_no", "kcharge"),
                     ("reverses", "krev")):
        B[dst] = B[src].map(key) if src in B.columns else ""
    B["bucket"] = B["parent_detalle"].map(
        lambda d: FEE_BY_ES.get(str(d), (None, "unclassified", None))[1])
    B["det"] = B[F_DETALLE].astype(str)
    B["order_level"] = B["sdate"].notna()
    cycle = 1 if site.BILL_CYCLE_DAY == 1 else infer_cycle_day(B, folder)
    p0, p1 = month_bounds(month)
    return {"folder": folder, "month": month, "y": y, "m": m, "p0": p0, "p1": p1,
            "B": B, "docs": documents(folder), "cycle_day": cycle}


def nominal_span(y, m, day):
    """The days the invoice named (y, m) covers by the cut-off rule."""
    p0 = pd.Timestamp(y, m, 1)
    if day == 1:
        return p0, (p0 + pd.offsets.MonthEnd(1)).normalize()
    return (p0 - pd.DateOffset(months=1)).replace(day=day), p0.replace(day=day - 1)


def views(ctx):
    """The two readings of the month: row selections plus what is missing / open."""
    B, y, m, day = ctx["B"], ctx["y"], ctx["m"], ctx["cycle_day"]
    docs = ctx["docs"]
    fac = {d["bill"]: d for d in docs if d["doc"] == "FAC"}
    nc = {d["bill"]: d for d in docs if d["doc"] == "NC"}

    # ---- invoice period: the documents named for the month ----
    s0, s1 = nominal_span(y, m, day)
    rows = B[B["_bill"].map(lambda b: b == (y, m))]
    issues = []
    d = fac.get((y, m))
    if not d:
        issues.append("没有 %s 的 Facturación（Reporte_Facturacion_MercadoLibre_…%d.xlsx）"
                      % (bill_name((y, m)), y))
    elif d["open"]:
        issues.append("%s 账单还没关账（文件顶部写着“podría tener cambios hasta el cierre”），"
                      "数字还会变" % bill_name((y, m)))
    if (y, m) not in nc:
        issues.append("没有 %s 的 Notas de crédito 文件（当期没有贷记单时平台不提供，"
                      "这本身不算缺）" % bill_name((y, m)))
    period = {"key": "period", "name": "账单期", "rows": rows, "start": s0, "end": s1,
              "bills": [(y, m)], "issues": issues,
              "complete": bool(d), "closed": bool(d) and not d["open"],
              "files": [x["name"] for x in docs if x["bill"] == (y, m)]}

    # ---- calendar cut: charge dates in the calendar month, from every invoice ----
    p0, p1 = ctx["p0"], ctx["p1"]
    crow = B[(B["cdate"] >= p0) & (B["cdate"] <= p1)]
    need = bill_cycles(y, m, day)
    cissues, complete, closed = [], True, True
    for by, bm, c0, c1 in need:
        d = fac.get((by, bm))
        a, b = max(c0, p0), min(c1, p1)
        if not d:
            complete = closed = False
            cissues.append("缺 %s 账单（覆盖 %s–%s）：本月 %s–%s 的计费没有数据"
                           % (bill_name((by, bm)), c0.strftime("%m/%d"), c1.strftime("%m/%d"),
                              a.strftime("%m/%d"), b.strftime("%m/%d")))
        elif d["open"]:
            closed = False
            cissues.append("%s 账单（覆盖 %s–%s）还没关账，本月 %s–%s 的数字还会变"
                           % (bill_name((by, bm)), c0.strftime("%m/%d"), c1.strftime("%m/%d"),
                              a.strftime("%m/%d"), b.strftime("%m/%d")))
    cal = {"key": "calendar", "name": "自然月切", "rows": crow, "start": p0, "end": p1,
           "bills": [(by, bm) for by, bm, _a, _b in need], "issues": cissues,
           "complete": complete, "closed": closed,
           "files": sorted(set(crow["_file"].astype(str)))}
    return period, cal


# ════════════════════════════════════════════════════════════════════════
# One reading -> numbers
# ════════════════════════════════════════════════════════════════════════

def _first(g, col):
    v = g[col].dropna() if col in g.columns else pd.Series(dtype=object)
    return v.iloc[0] if len(v) else None


def billed_sales(rows):
    """One row per billed sale (Número de venta) of the commission charges.

    'Cargo por venta' is the anchor (it carries the sale); an affiliate charge
    on the same sale only adds commission. Sales without a sale number cannot
    be told apart and are kept one per charge row."""
    c = rows[(~rows["is_rev"]) & (rows["bucket"] == "commission")].copy()
    if not len(c):
        return pd.DataFrame(columns=["ksale", "kpack", "sdate", "cdate", "units", "price", "total",
                                     "commission", "title", "pub", "state", "channel", "doc",
                                     "file", "charge_no"])
    c["_anchor"] = (c["det"] != SALE).astype(int)
    c = c.sort_values(["_anchor", "cdate"], kind="stable")
    c["_k"] = [k if k else "row:%s" % i for k, i in zip(c["ksale"], c.index)]
    out = []
    for k, g in c.groupby("_k", sort=False):
        a = g.iloc[0]
        out.append({
            "ksale": a["ksale"], "kpack": a["kpack"], "sdate": a["sdate"], "cdate": a["cdate"],
            "units": float(pd.to_numeric(pd.Series([a.get(C_QTY)]), errors="coerce").fillna(0)[0]),
            "price": float(pd.to_numeric(pd.Series([a.get(C_PRICE)]), errors="coerce").fillna(0)[0]),
            "total": float(pd.to_numeric(pd.Series([a.get(C_TOTAL)]), errors="coerce").fillna(0)[0]),
            "commission": float(g["amount"].sum()),
            "title": _first(g, C_TITLE) or "", "pub": key(_first(g, C_PUB)),
            "state": key(a.get(F_STATE)), "channel": key(a.get(F_CHANNEL)),
            "doc": a["_doc"], "file": a["_file"], "charge_no": a["kcharge"]})
    return pd.DataFrame(out)


def _parent_maps(B):
    """Lookups over every downloaded invoice: charge no -> row, sale -> commission,
    listing -> the fee terms seen on it."""
    by_charge = {}
    for i, k in zip(B.index, B["kcharge"]):
        if k:
            by_charge[k] = i
    com = B[(~B["is_rev"]) & (B["bucket"] == "commission")]
    by_sale = {}
    for k, amt, tot, qty in zip(com["ksale"], com["amount"], _num(com, C_TOTAL), _num(com, C_QTY)):
        if not k:
            continue
        s = by_sale.setdefault(k, [0.0, None, None])
        s[0] += float(amt)
        if s[1] is None and pd.notna(tot):
            s[1], s[2] = float(tot), (float(qty) if pd.notna(qty) else None)
    # Fee terms seen per listing: (category %, fixed fee per unit, sale date).
    # The category % of a listing changes over time (BOCINA_TA02 MLM5954700702:
    # 14.5 % in early September, 11 % from October), so a median would move
    # whenever a later invoice is downloaded; the terms themselves do not.
    terms = {}
    cv = com[com["det"] == SALE]
    tot, qty = _num(cv, C_TOTAL), _num(cv, C_QTY).fillna(1.0)
    pct = _num(cv, C_PCT)
    unit = _num(cv, C_UNIT_FEE).fillna(0.0)
    if C_PUB in cv.columns:
        for p, a, t, q, pc, u, d in zip(cv[C_PUB].map(key), cv["amount"], tot, qty, pct, unit,
                                        cv["sdate"]):
            if not p or not (pd.notna(t) and t > 0):
                continue
            if pd.isna(pc):                       # no % column: the charge's own rate
                pc, u = float(a) / float(t), 0.0
            terms.setdefault(p, set()).add((round(float(pc), 6), round(float(u) / max(q, 1.0), 4),
                                            d if pd.notna(d) else pd.NaT))
    return by_charge, by_sale, terms


def _from_terms(terms, amt, total, qty, sdate, tol):
    """Original commission of a sale whose charge is not downloaded, from the
    listing's fee terms: a term that reproduces the reversed amount exactly means
    a full reversal; otherwise the term of the nearest sale date."""
    cands = [(round(pc * total + u * qty, 2), d) for pc, u, d in terms]
    for c, _d in cands:
        if abs(c - amt) <= tol:
            return c
    ref = sdate if pd.notna(sdate) else None
    cands.sort(key=lambda x: (abs((x[1] - ref).days) if ref is not None and pd.notna(x[1])
                              else 10 ** 6, x[0]))
    return cands[0][0] if cands else None


def reversals(rows, B):
    """Sale reversals in these rows: how much of which sale was undone, and how
    each was linked to the original charge."""
    rv = rows[rows["is_rev"] & (rows["parent_detalle"].astype(str) == SALE)]
    by_charge, by_sale, terms = _parent_maps(B)
    out = []
    for i, r in rv.iterrows():
        amt = -float(r["amount"])
        own_total = pd.to_numeric(pd.Series([r.get(C_TOTAL)]), errors="coerce")[0]
        own_qty = pd.to_numeric(pd.Series([r.get(C_QTY)]), errors="coerce")[0]
        orig = total = qty = None
        parent_no, parent_date, parent_file = r["krev"], None, None
        j = by_charge.get(r["krev"]) if r["krev"] else None
        if j is not None and not B.at[j, "is_rev"]:
            method = LINK_ID
            orig = float(B.at[j, "amount"])
            total = pd.to_numeric(pd.Series([B.at[j, C_TOTAL] if C_TOTAL in B.columns else None]),
                                  errors="coerce")[0]
            qty = pd.to_numeric(pd.Series([B.at[j, C_QTY] if C_QTY in B.columns else None]),
                                errors="coerce")[0]
            parent_date, parent_file = B.at[j, "cdate"], B.at[j, "_file"]
        elif r["ksale"] in by_sale:
            method = LINK_SALE
            orig, total, qty = by_sale[r["ksale"]]
        elif terms.get(key(r.get(C_PUB))) and pd.notna(own_total) and own_total > 0:
            method = LINK_RATE
            orig = _from_terms(terms[key(r.get(C_PUB))], amt, float(own_total),
                               float(own_qty) if pd.notna(own_qty) else 1.0, r["sdate"],
                               site.CN_TOL)
        else:
            method = LINK_NONE
        if total is None or pd.isna(total):
            total = own_total
        if qty is None or pd.isna(qty):
            qty = own_qty
        total = float(total) if pd.notna(total) else 0.0
        qty = float(qty) if pd.notna(qty) else 0.0
        ratio = 1.0 if not orig or orig <= 0 else min(1.0, amt / orig)
        out.append({"idx": i, "ksale": r["ksale"], "kpack": r["kpack"], "sdate": r["sdate"],
                    "cdate": r["cdate"], "doc": r["_doc"], "file": r["_file"],
                    "charge_no": r["kcharge"], "parent_no": parent_no,
                    "parent_date": parent_date, "parent_file": parent_file,
                    "rev_commission": amt, "orig_commission": orig, "ratio": ratio,
                    "sale_total": total, "units": qty, "est_sales": ratio * total,
                    "est_units": ratio * qty, "method": method,
                    "title": r.get(C_TITLE) if isinstance(r.get(C_TITLE), str) else ""})
    return pd.DataFrame(out, columns=["idx", "ksale", "kpack", "sdate", "cdate", "doc", "file",
                                      "charge_no", "parent_no", "parent_date", "parent_file",
                                      "rev_commission", "orig_commission", "ratio", "sale_total",
                                      "units", "est_sales", "est_units", "method", "title"])


def buyer_shipping(rows, B):
    """Buyer-paid shipping ('Envío a cargo del cliente') billed and reversed here.

    The column sits on both the commission and the shipping row of a sale with
    the same value; it is taken from the shipping rows (from the commission row
    only for a sale without a shipping row). A reversal of the shipping charge
    gives the buyer's shipping back in the same proportion."""
    if C_BUYER_SHIP not in rows.columns:
        return 0.0, 0.0, 0
    e = _num(rows, C_BUYER_SHIP).fillna(0.0)
    ship = rows[(~rows["is_rev"]) & (rows["det"] == SHIP)]
    got = float(e[ship.index].sum())
    n = int((e[ship.index] > 0).sum())
    have = set(ship["ksale"])
    cv = rows[(~rows["is_rev"]) & (rows["det"] == SALE) & ~rows["ksale"].isin(have)]
    got += float(e[cv.index].sum())
    n += int((e[cv.index] > 0).sum())
    back = 0.0
    rv = rows[rows["is_rev"] & (rows["parent_detalle"].astype(str) == SHIP)]
    by_charge = dict((k, i) for i, k in zip(B.index, B["kcharge"]) if k)
    eb = _num(B, C_BUYER_SHIP).fillna(0.0)
    for i, r in rv.iterrows():
        j = by_charge.get(r["krev"])
        ev = float(e[i]) or (float(eb[j]) if j is not None else 0.0)
        if not ev:
            continue
        ratio = 1.0
        if j is not None and float(B.at[j, "amount"]) > 0:
            ratio = min(1.0, -float(r["amount"]) / float(B.at[j, "amount"]))
        back += ev * ratio
    return got, back, n


def fee_table(rows):
    """Fees by FEE_CATALOG line (gross, reversals, net) + unregistered Detalle values."""
    lines = []
    for es, zh, bucket, order_level in FEE_CATALOG:
        f = rows[rows["parent_detalle"].astype(str) == es]
        g, r = f[~f["is_rev"]], f[f["is_rev"]]
        lines.append({"es": es, "zh": zh, "bucket": bucket, "order_level": order_level,
                      "gross": float(g["amount"].sum()), "n_gross": int(len(g)),
                      "rev": float(r["amount"].sum()), "n_rev": int(len(r)),
                      "net": float(f["amount"].sum())})
    unk = rows[~rows["parent_detalle"].astype(str).isin(FEE_BY_ES)]
    unknown = []
    for det, g in unk.groupby("det"):
        unknown.append({"es": det, "n": int(len(g)), "amount": float(g["amount"].sum()),
                        "order_level": bool(g["order_level"].any()),
                        "channels": ", ".join(sorted(set(g["channel"].map(key)) - {""}))})
    unknown.sort(key=lambda u: -abs(u["amount"]))
    return lines, unknown


def analyze(view, B):
    """Every invoice-basis number of one reading (period or calendar cut)."""
    rows = view["rows"]
    sales = billed_sales(rows)
    rev = reversals(rows, B)
    ship_in, ship_back, ship_n = buyer_shipping(rows, B)
    lines, unknown = fee_table(rows)
    fee_total = float(rows["amount"].sum())
    by_bucket = {}
    for ln in lines:
        by_bucket.setdefault(ln["bucket"], {"gross": 0.0, "rev": 0.0, "net": 0.0})
        for k in ("gross", "rev", "net"):
            by_bucket[ln["bucket"]][k] += ln[k]
    unk_order = sum(u["amount"] for u in unknown if u["order_level"])
    unk_bill = sum(u["amount"] for u in unknown if not u["order_level"])
    gross = float(sales["total"].sum()) if len(sales) else 0.0
    rsales = float(rev["est_sales"].sum()) if len(rev) else 0.0
    n = {
        "rows": int(len(rows)), "n_fac": int((rows["_doc"] == "FAC").sum()),
        "n_nc": int((rows["_doc"] == "NC").sum()),
        "date_min": rows["cdate"].min() if len(rows) else pd.NaT,
        "date_max": rows["cdate"].max() if len(rows) else pd.NaT,
        "n_sales": int(len(sales)), "units": float(sales["units"].sum()) if len(sales) else 0.0,
        "gross_sales": gross, "rev_sales": rsales, "net_sales": gross - rsales,
        "n_rev": int(len(rev)), "rev_units": float(rev["est_units"].sum()) if len(rev) else 0.0,
        "rev_full": int((rev["ratio"] >= 0.999).sum()) if len(rev) else 0,
        "rev_methods": {m: int((rev["method"] == m).sum()) for m in LINK_ORDER} if len(rev)
        else {m: 0 for m in LINK_ORDER},
        "ship_income": ship_in, "ship_income_rev": ship_back, "ship_n": ship_n,
        "fee_total": fee_total, "unk_order": unk_order, "unk_bill": unk_bill,
        "channel": {c: float(rows[rows["channel"].map(key) == c]["amount"].sum())
                    for c in ("Si", "No", "No aplica")},
    }
    n["sale_months"] = (sales["sdate"].dt.to_period("M").astype(str).value_counts().to_dict()
                        if len(sales) else {})
    n["contribution"] = n["net_sales"] + ship_in - ship_back - fee_total
    return {"view": view, "sales": sales, "reversals": rev, "fees": lines, "unknown": unknown,
            "by_bucket": by_bucket, "n": n}
