# -*- coding: utf-8 -*-
"""The current order basis (read-only through derive()) and its comparison with
the invoice basis: orders of the month <-> billed sales, line by line.

derive() and report_build.build_context() are called exactly as the monthly
report calls them; nothing here changes how they compute."""

import re
from collections import OrderedDict

import pandas as pd

from platforms.mercadolibre import report_build
from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (FEE_BY_ES, FEE_CATALOG, V_SKU, V_STATUS,
    V_TITLE, V_UNITS, fmt, month_bounds, num)
from platforms.mercadolibre.report.derived import derive

from experiments.ml_invoice_basis.common import key
from experiments.ml_invoice_basis.invoice import bill_name

PACK_RE = re.compile(r"^\s*Paquete de (\d+) productos?", re.I)
CANCEL_RE = re.compile(r"cancel", re.I)


def current_basis(store, folder, month):
    """The monthly report's own numbers for this store/month/folder (site already set)."""
    p0, p1 = month_bounds(month)
    S = derive(store, folder, p0, p1)
    ctx = report_build.build_context([S], month)
    S["contrib"] = ctx["contrib"][store]
    return S


# ════════════════════════════════════════════════════════════════════════
# Ventas -> orders (a multi-product pack folded into its parent row)
# ════════════════════════════════════════════════════════════════════════

def order_index(V):
    """{order key: record} and {pack child id: parent key}.

    Ventas writes a pack as a parent row "Paquete de N productos" that carries
    the money, followed by N item rows with units and SKU but no money and no
    shared id. The invoice bills each item (Número de venta = the item row's id)
    and gives the pack in Número de paquete (= the parent row's id), so a billed
    sale is matched on the pack first, then on its own id."""
    cols = {c: (num(V, c) if c in V.columns else pd.Series(0.0, index=V.index))
            for c in (site.V_REVENUE, site.V_REFUND, site.V_TOTAL, site.V_SHIP_INCOME,
                      site.V_SHIP_COST, site.V_COMMISSION, V_UNITS)}
    st = V[V_STATUS].astype(str) if V_STATUS in V.columns else pd.Series("", index=V.index)
    sku = V[V_SKU] if V_SKU in V.columns else pd.Series("", index=V.index)
    title = V[V_TITLE] if V_TITLE in V.columns else pd.Series("", index=V.index)
    coupon = V["coupon"] if "coupon" in V.columns else pd.Series(0.0, index=V.index)
    orders, parent_of = OrderedDict(), {}
    parent, pending = None, 0
    for i in V.index:
        oid = key(V.at[i, "oid"])
        if not oid:
            continue
        s = st[i] if isinstance(st[i], str) and st[i] != "nan" else ""
        m = PACK_RE.match(s)
        if m:
            parent, pending = oid, int(m.group(1))
            rec = _rec(orders, oid, V.at[i, "fecha"], s, pack=True)
        elif pending > 0 and parent:
            pending -= 1
            parent_of[oid] = parent
            rec = orders[parent]
            rec["item_status"].append(s)
        else:
            parent, pending = None, 0
            rec = _rec(orders, oid, V.at[i, "fecha"], s)
        rec["gmv"] += float(cols[site.V_REVENUE][i])
        rec["refund"] += -float(cols[site.V_REFUND][i])
        rec["total"] += float(cols[site.V_TOTAL][i])
        rec["ship_income"] += float(cols[site.V_SHIP_INCOME][i])
        rec["ship_cost"] += -float(cols[site.V_SHIP_COST][i])
        rec["com_col"] += -float(cols[site.V_COMMISSION][i])
        rec["units"] += float(cols[V_UNITS][i])
        rec["coupon"] += float(coupon[i]) if pd.notna(coupon[i]) else 0.0
        k = key(sku[i])
        if k and k not in rec["skus"]:
            rec["skus"].append(k)
        if not rec["title"] and isinstance(title[i], str):
            rec["title"] = title[i]
    for rec in orders.values():
        if rec["pack"] and rec["item_status"]:
            rec["status"] = " / ".join(sorted(set(rec["item_status"])))
    return orders, parent_of


def _rec(orders, oid, date, status, pack=False):
    if oid not in orders:          # exchanges repeat the same id on two rows
        orders[oid] = {"oid": oid, "date": date, "status": status, "pack": pack,
                       "item_status": [], "gmv": 0.0, "refund": 0.0, "total": 0.0,
                       "ship_income": 0.0, "ship_cost": 0.0, "com_col": 0.0, "units": 0.0,
                       "coupon": 0.0, "skus": [], "title": ""}
    return orders[oid]


def order_key(orders, parent_of, kpack, ksale):
    if kpack and kpack in orders:
        return kpack
    if ksale in parent_of:
        return parent_of[ksale]
    if ksale in orders:
        return ksale
    if kpack in parent_of:
        return parent_of[kpack]
    return ""


def month_orders(orders, month):
    p0, p1 = month_bounds(month)
    end = p1 + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return OrderedDict((k, r) for k, r in orders.items()
                       if pd.notna(r["date"]) and p0 <= r["date"] <= end)


# ════════════════════════════════════════════════════════════════════════
# Per-sale reconciliation of one invoice reading against the month's orders
# ════════════════════════════════════════════════════════════════════════

def billed_where(B, orders, parent_of):
    """{order key: [(charge date, bill)]} of every non-reversed commission charge."""
    com = B[(~B["is_rev"]) & (B["bucket"] == "commission")]
    out = {}
    for kp, ks, d, b in zip(com["kpack"], com["ksale"], com["cdate"], com["_bill"]):
        k = order_key(orders, parent_of, kp, ks)
        if k:
            out.setdefault(k, []).append((d, b))
    return out


def reconcile(inv, orders, parent_of, month, where, B):
    """Billed sales of one reading vs the orders of the month (order date).

    The GMV difference splits exactly into: billed here but not this month's
    orders (+), this month's orders not billed here (−), and amount differences
    of the orders on both sides."""
    y, m = int(month[:4]), int(month[5:7])
    mo = month_orders(orders, month)
    sales = inv["sales"].copy()
    sales["okey"] = [order_key(orders, parent_of, kp, ks)
                     for kp, ks in zip(sales.get("kpack", []), sales.get("ksale", []))]
    fully_rev = set(B[B["is_rev"] & (B["bucket"] == "commission")]["ksale"])
    extra = []
    for _, s in sales[~sales["okey"].isin(list(mo))].iterrows():
        o = orders.get(s["okey"]) if s["okey"] else None
        if not o:
            code = "absent_rev" if s["ksale"] in fully_rev else "absent"
            why = ("Ventas 里查无此单，账单上有它的冲销 —— 多为下单后随即取消的订单（Ventas 不列这类订单）；"
                   "也可能是 Ventas 导出缺数据" if code == "absent_rev" else
                   "Ventas 里查无此单 —— Ventas 导出缺数据，或超出了导出窗口")
            odate = pd.NaT
        else:
            odate = o["date"]
            if pd.isna(odate):
                code, why = "nodate", "Ventas 里有，但没有下单日期"
            elif (odate.year, odate.month) < (y, m):
                code, why = "prev", "%d 年 %d 月的订单，本期才计费" % (odate.year, odate.month)
            else:
                code, why = "next", "%d 年 %d 月的订单，本期已计费" % (odate.year, odate.month)
        extra.append({"ksale": s["ksale"], "okey": s["okey"], "sdate": s["sdate"],
                      "cdate": s["cdate"], "odate": odate, "units": s["units"],
                      "total": s["total"], "title": s["title"], "state": s["state"],
                      "code": code, "why": why})
    billed = set(sales["okey"]) - {""}
    missing = []
    for k, o in mo.items():
        if k in billed:
            continue
        w = where.get(k, [])         # billed, but not in this reading (else it would match)
        if w:
            code = "elsewhere"
            why = "在 %s 计费" % "、".join(sorted(set(
                "%s 账单（%s）" % (bill_name(b), d.strftime("%m/%d") if pd.notna(d) else "?")
                for d, b in w)))
        elif CANCEL_RE.search(o["status"] or ""):
            code, why = "cancel", "订单已取消，平台没有计费（多为发货前取消）"
        elif abs(o["gmv"]) < 0.005:
            code, why = "zero", "Ventas 上金额为 0"
        else:
            code, why = "unbilled", "已下载的账单里都没有这笔的佣金（尚未出账单，或免佣金）"
        missing.append({"okey": k, "odate": o["date"], "status": o["status"], "units": o["units"],
                        "gmv": o["gmv"], "refund": o["refund"], "sku": ", ".join(o["skus"]),
                        "title": o["title"], "code": code, "why": why})
    matched = []
    tol = site.CN_TOL
    g = sales[sales["okey"].isin(list(mo))].groupby("okey").agg(
        total=("total", "sum"), units=("units", "sum"))
    for k, r in g.iterrows():
        o = mo[k]
        matched.append({"okey": k, "inv_total": float(r["total"]), "gmv": o["gmv"],
                        "inv_units": float(r["units"]), "units": o["units"],
                        "diff": float(r["total"]) - o["gmv"], "status": o["status"],
                        "odate": o["date"]})
    E = pd.DataFrame(extra, columns=["ksale", "okey", "sdate", "cdate", "odate", "units", "total",
                                     "title", "state", "code", "why"])
    M = pd.DataFrame(missing, columns=["okey", "odate", "status", "units", "gmv", "refund", "sku",
                                       "title", "code", "why"])
    T = pd.DataFrame(matched, columns=["okey", "inv_total", "gmv", "inv_units", "units", "diff",
                                       "status", "odate"])
    return {"sales": sales, "extra": E, "missing": M, "matched": T,
            "n_month_orders": len(mo),
            "month_gmv": float(sum(o["gmv"] for o in mo.values())),
            "month_units": float(sum(o["units"] for o in mo.values())),
            "differ": T[T["diff"].abs() > tol]}


def rec_summary(rec):
    """The numbers behind a revenue line's explanation."""
    E, M, T = rec["extra"], rec["missing"], rec["matched"]
    cnt = lambda df, c: int((df["code"] == c).sum()) if len(df) else 0
    return {
        "extra_n": int(len(E)), "extra_total": float(E["total"].sum()) if len(E) else 0.0,
        "extra_units": float(E["units"].sum()) if len(E) else 0.0,
        "extra_prev": cnt(E, "prev"), "extra_next": cnt(E, "next"),
        "extra_absent": cnt(E, "absent") + cnt(E, "absent_rev") + cnt(E, "nodate"),
        "missing_n": int(len(M)), "missing_gmv": float(M["gmv"].sum()) if len(M) else 0.0,
        "missing_units": float(M["units"].sum()) if len(M) else 0.0,
        "missing_elsewhere": cnt(M, "elsewhere"), "missing_cancel": cnt(M, "cancel"),
        "missing_other": cnt(M, "unbilled") + cnt(M, "zero"),
        "matched_n": int(len(T)), "matched_diff": float(T["diff"].sum()) if len(T) else 0.0,
        "matched_unit_diff": float((T["inv_units"] - T["units"]).sum()) if len(T) else 0.0,
        "differ_n": int(len(rec["differ"])),
    }


# ════════════════════════════════════════════════════════════════════════
# Line-by-line comparison
# ════════════════════════════════════════════════════════════════════════

CAT = {
    "time": "月初/月末时间差（下单日 ≠ 计费日）",
    "other": "别的月份开票的销售",
    "rev": "冲销估算 ↔ 实际退款",
    "absent": "账单不含此项",
    "cycle": "智利账单周期",
    "source": "取数来源不同（Ventas ↔ 账单）",
    "unit": "计数单位不同",
    "same": "同源同期",
    "na": "账单口径不需要",
}


def _split(view_rows, L, bucket, rev):
    """Commission-type rows: billed in this reading but not linked to the month's
    orders (+) and linked to them but billed outside this reading (−)."""
    a = view_rows[(view_rows["bucket"] == bucket) & (view_rows["is_rev"] == rev)]
    b = L[(L["bucket"] == bucket) & (L["is_rev"] == rev)] if len(L) else L
    inv_only = float(a[~a.index.isin(b.index)]["amount"].sum())
    cur_only = float(b[~b.index.isin(a.index)]["amount"].sum()) if len(b) else 0.0
    return inv_only, cur_only


def _rev_text(rs):
    s = rs
    return ("本期开票、不是本月订单 +%d 笔 %s（上月订单 %d、下月订单 %d、Ventas 查无 %d）；"
            "本月订单、不在本期开票 −%d 笔 %s（在别的账单计费 %d、取消未计费 %d、其他 %d）；"
            "两边都有的订单金额差 %s"
            % (s["extra_n"], fmt(s["extra_total"]), s["extra_prev"], s["extra_next"],
               s["extra_absent"], s["missing_n"], fmt(s["missing_gmv"]), s["missing_elsewhere"],
               s["missing_cancel"], s["missing_other"], fmt(s["matched_diff"])))


def comparison(S, invs, recs, cycle_day):
    """Rows for ③: (key, label, original term, kind, current, {view: value},
    category keys, {view: note}). kind: 'n' count, 'm' money, 'sub' subtotal,
    'tot' total. Money is signed as on the P&L (deductions negative)."""
    m = S["m"]
    L = S["billing_linked"]
    if "bucket" not in L.columns:
        L = L.assign(bucket=L["parent_detalle"].map(
            lambda d: FEE_BY_ES.get(str(d), (None, "unclassified", None))[1]))
    cl = cycle_day != 1
    views = list(invs)
    rows = []

    def add(k, label, es, kind, cur, vals, cats, notes):
        rows.append({"key": k, "label": label, "es": es, "kind": kind, "cur": cur,
                     "inv": vals, "cats": cats, "notes": notes})

    def per(fn):
        return dict((v, fn(v)) for v in views)

    def cyc(v, cats):
        return cats + (["cycle"] if cl and v == "period" else [])

    rs = dict((v, rec_summary(recs[v])) for v in views)
    add("orders", "订单 / 开票销售笔数", "# de venta ↔ Número de venta", "n", m["orders_unique"],
        per(lambda v: invs[v]["n"]["n_sales"]), ["unit", "time", "other"],
        per(lambda v: "现行数的是本月下单的 Ventas 行（多件包裹的母行和每件子行各算一行）；账单一件商品一笔"
                      "（Número de venta）。按订单对照：" + _rev_text(rs[v])))
    add("units", "件数", "Unidades ↔ Cantidad vendida", "n", m["units"],
        per(lambda v: invs[v]["n"]["units"]), ["time", "other"],
        per(lambda v: "本期开票、不是本月订单 +%s 件；本月订单、不在本期开票 −%s 件；两边都有的订单件数差 %s"
                      % (fmt(rs[v]["extra_units"], 0), fmt(rs[v]["missing_units"], 0),
                         fmt(rs[v]["matched_unit_diff"], 0))))
    add("gmv", "商品销售额 (GMV)", "Ingresos por productos ↔ Total de la venta", "m", m["ing"],
        per(lambda v: invs[v]["n"]["gross_sales"]), ["time", "other"], per(lambda v: _rev_text(rs[v])))
    add("refund", "减：取消与退款", "Anulaciones y reembolsos ↔ Anulación del cargo por venta",
        "m", -m["refund"], per(lambda v: -invs[v]["n"]["rev_sales"]), ["rev", "time"],
        per(lambda v: "现行 = Ventas 退款列：退给买家的净额（平台退回的佣金/运费/税已抵掉），按下单月；"
                      "账单 = 本期冲销的销售额估算（冲销佣金 ÷ 原佣金 × 售价），按冲销的计费日。"
                      "%d 笔冲销，其中全额 %d 笔；关联方法见 ⑤" % (invs[v]["n"]["n_rev"],
                                                    invs[v]["n"]["rev_full"])))
    add("cn_adj", "加：部分退款重复扣佣金调整", "(现行报表修正)", "m", S["cn"]["partial_total"],
        per(lambda v: 0.0), ["na"],
        per(lambda v: "修正的是 Ventas 退款列的重复扣佣金；账单口径不用 Ventas，按比例估算冲销，不需要这一行"))
    add("ship_in", "运费收入（买家支付）", "Ingresos por envío ↔ Envío a cargo del cliente", "m",
        m["env_ing"],
        per(lambda v: invs[v]["n"]["ship_income"] - invs[v]["n"]["ship_income_rev"]),
        ["time", "source"] + (["absent"] if site.BONIF_COL else []),
        per(lambda v: "账单的“Envío a cargo del cliente”与 Ventas 运费收入逐单相等（实测），差额来自期间；"
                      "本期 %d 笔，冲销退回 %s%s" % (
                          invs[v]["n"]["ship_n"], fmt(invs[v]["n"]["ship_income_rev"]),
                          "。秘鲁的 Flex 配送补贴（%s）不在账单里" % fmt(m.get("bonif", 0.0))
                          if site.BONIF_COL else "")))
    add("coupon", "减：卖家承担优惠券", "Descuentos y bonificaciones（由 Total 倒推）", "m",
        -m["coupon"], per(lambda v: None), ["absent"],
        per(lambda v: "账单里没有买家优惠券（账单的 Valor del descuento 是平台给卖家费用的折扣，已含在费用净额里）"))
    add("_net_sales", "净收入小计", "", "sub", None, {}, [], {})

    def order_fee(bucket, label, es, cur, note_extra=""):
        def note(v):
            a, b = _split(invs[v]["view"]["rows"], L, bucket, False)
            return ("现行取本月订单的这项费用（不论哪张账单计费）；账单取本期计费的。本期计费、不是本月订单的 +%s；"
                    "本月订单、在本期以外计费的 −%s%s" % (fmt(a), fmt(b), note_extra))
        add(bucket, label, es, "m", cur,
            per(lambda v: -invs[v]["by_bucket"].get(bucket, {}).get("gross", 0.0)),
            ["time"] + (["source"] if bucket == "shipping" else []), per(note))

    order_fee("commission", "减：平台销售佣金（毛额）", "Cargo por venta (+ afiliados)", -m["com_gross"])
    order_fee("shipping", "减：平台配送费", "Costos de envío ↔ Cargo por envíos de Mercado Libre",
              -m["env_cost"], "。另外现行取的是 Ventas 的 Costos de envío，与账单上本月订单的配送费差 %s"
              % fmt(m["env_cost"] - m["env_gross"]))
    add("returns", "减：退货处理费（净）", "Cargo por devolución", "m", -m["dev_net"],
        per(lambda v: -invs[v]["by_bucket"].get("returns", {}).get("net", 0.0)), ["time"],
        per(lambda v: "现行 = 本月订单关联的退货处理费及其冲销；账单 = 本期计费的退货处理费及其冲销"))
    add("order_rev", "加：取消/退款退回的佣金与配送费", "Anulación del cargo por venta / envíos", "m",
        0.0, per(lambda v: -(invs[v]["by_bucket"].get("commission", {}).get("rev", 0.0)
                             + invs[v]["by_bucket"].get("shipping", {}).get("rev", 0.0))),
        ["rev"],
        per(lambda v: "现行口径没有这一行：退回的佣金/配送费已抵在“取消与退款”的净额里；"
                      "账单口径把它们单列（冲销毛额与费用冲销分开）"))
    add("unc_order", "减：订单级未登记科目 ⚠", "(未在 FEE_CATALOG 登记)", "m", -m["unc_order"],
        per(lambda v: -invs[v]["n"]["unk_order"]), ["same"],
        per(lambda v: "字典里没有的订单级费用，逐项见 ④"))
    add("_order_fees", "订单级费用小计", "", "sub", None, {}, [], {})
    for es, zh, bucket, order_level in FEE_CATALOG:
        if order_level:
            continue
        net = dict((v, sum(f["net"] for f in invs[v]["fees"] if f["es"] == es)) for v in views)
        add("fee:" + es, "减：" + zh, es, "m", -m["bill_fee_net"].get(es, 0.0),
            dict((v, -net[v]) for v in views),
            per(lambda v: cyc(v, ["same"])),
            per(lambda v: ("现行按计费日归自然月；账单期是 %s–%s" % (
                invs[v]["view"]["start"].strftime("%m/%d"), invs[v]["view"]["end"].strftime("%m/%d"))
                if cl and v == "period" else "两边都是按计费日归入本月的同一批账单行")))
    add("unc_bill", "减：账单级未登记科目 ⚠", "(未在 FEE_CATALOG 登记)", "m", -m["unc_bill"],
        per(lambda v: -invs[v]["n"]["unk_bill"]), per(lambda v: cyc(v, ["same"])),
        per(lambda v: "字典里没有的账单级费用，逐项见 ④"))
    add("_bill_fees", "月度账单费用小计", "", "sub", None, {}, [], {})
    add("_contrib", "经营贡献毛利（未扣商品成本）", "", "tot", S["contrib"],
        per(lambda v: invs[v]["n"]["contribution"]), [], {})
    return rows


def unavailable(S):
    """What the invoices cannot give: (label, original term, current value, why)."""
    m = S["m"]
    out = []
    if site.WITHHOLDING:
        out.append(("平台代扣代缴税金 (IVA/ISR)", "Retenciones (en Cargo por venta e impuestos)",
                    m["tax_ret"], "账单只有纯佣金；代扣税只在 Ventas 的“佣金及税金”列里和佣金混在一起，"
                                  "要用 Ventas 减账单才得到。现行报表不把它当费用（是可抵税款），但报税要用"))
    out.append(("卖家承担优惠券", "Descuentos y bonificaciones", m["coupon"],
                "账单没有买家优惠；现行报表由 Ventas 的 Total 倒推"))
    out.append(("实际退给买家的金额（净额）", "Anulaciones y reembolsos", m["refund"],
                "账单只有被冲销的佣金/配送费；退给买家多少钱只能按冲销比例估算（见 ② 的冲销行）"))
    if site.BONIF_COL:
        out.append(("Flex 配送补贴（秘鲁）", "Descuentos y bonificaciones", m.get("bonif", 0.0),
                    "平台付给卖家的配送补贴，在 Ventas 和 MercadoPago 里，不在账单里"))
    out.append(("订单应到账净额（Ventas Total）", "Total", m["total"],
                "买家付款扣掉直扣费用后应到账的钱；账单没有收款信息"))
    mp = S.get("mp")
    if mp:
        out.append(("MercadoPago 本月实际到账（结算流水）", "settlement REAL_AMOUNT", mp.get("real_all", 0.0),
                    "现金在 MercadoPago，不在账单里"))
    else:
        out.append(("MercadoPago 本月实际到账", "settlement REAL_AMOUNT", None,
                    "本次下载没有可用的 MercadoPago 结算流水；无论如何账单里都没有现金"))
    out.append(("商品采购成本", "", None, "两种口径的平台数据都没有（要商品成本表）"))
    return out
