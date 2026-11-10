# -*- coding: utf-8 -*-
"""
⑪ 账单逐单明细 — Liverpool's settlement bill, order by order, with the shipping fee
and the seller SKU of each order on the same row (owner's requirement 2026-11-10:
"reports based on the settled bills"; finance used to build this by hand from three
downloads).

    the bill      "Ventas Marketplace" (LV_ventas): one row per order of a sale or
                  return document — price, Liverpool's discount, commission, net payment
    shipping      "Costos logísticos" (LV_costos_logisticos), looked up by order number
    extra charges "Balance de cuenta" (LV_balance), looked up by its Remisión (order)
    SKU           the order list (LV_pedidos), looked up by order number; an order
                  older than the list falls back to Liverpool's article number

What the bill's "Pago neto" is (measured on 2026-09 and 2026-10, document by document
against the ledger): Venta con descuento − Comisión Dilisa − the surcharge for a
non-Liverpool card (Sobrecomis % T.Exter). It does not deduct the withheld taxes.

Rules:
* The bill decides the month: every row of the month's bill is a row here.
* An order's shipping fee is shown once, on its first sale row. The month's own
  list is one column (it ties to the month); a fee charged in a neighbouring month
  (found in that month's download folder when there is one) is a separate column,
  for the order's full picture, and is not part of this month's totals.
* Nothing is dropped: the month's shipping fees and extra charges whose order is
  not in the month's bill are listed below the table.
* Liverpool's own discount ("Descuento") lowers the bill's net payment, and Liverpool
  pays it back to the seller on a separate ledger document (C2, same Referencia):
  it is its own column, so an order's net is what the seller really gets.
* The ledger stays the proof: the sheet ends with bill net − the month's other
  charges = ②'s contribution, and the difference is a check (bill_tie).
"""
import glob
import os

import pandas as pd
from openpyxl.utils import get_column_letter

from core import report_family as RF
from core import report_kit as K
from core.report_kit import F, fill
from platforms.liverpool import catalog as C
from platforms.liverpool import periods
from platforms.liverpool import readers as R

SHEET = "⑪ 账单逐单明细"
TIE_TOL = 1.00          # MXN: the bill rounds the card surcharge per order (≤ 0.03 a row)
SKU_FROM_ORDERS, SKU_FROM_ARTICLE, SKU_NONE = "订单列表", "按 Liverpool 商品号", "未找到"


def _previous(month):
    y, m = periods.parse_month(month)
    return "%04d-%02d" % (y - (m == 1), 12 if m == 1 else m - 1)


def neighbour_shipping(folder, month):
    """order -> (amount, charge date) from the logistic-cost lists of the months next
    to this one, when the store's data folder holds them (<store>/<YYYY-MM>/<run>/).
    A folder that is not laid out that way has no neighbours."""
    out = {}
    month_dir = os.path.dirname(os.path.abspath(folder))
    if os.path.basename(month_dir) != month:
        return out
    for mon in (periods.next_month(month), _previous(month)):
        hits = sorted(glob.glob(os.path.join(os.path.dirname(month_dir), mon, "*",
                                             R.LOGISTICS % periods.stamp(mon))))
        if not hits:
            continue
        try:
            lst = R.read_logistics(hits[-1])
        except Exception:                      # noqa: BLE001 — a neighbour is a courtesy
            continue
        for r in lst.itertuples(index=False):
            out.setdefault(r.order, (float(r.amount or 0.0), r.charged))
    return out


def derive(S, sku_maps):
    """S["bill"] = {"rows", "ship_only", "extra_only"} and S["m"]["bill"] (totals and
    the tie to ②). sku_maps = (order -> (sku, title, qty, article), article -> (sku, title))."""
    v, m = S["sales"], S["m"]
    if v is None or not len(v):
        S["bill"] = None
        m["bill"] = None
        return
    by_order, by_article = sku_maps
    o = S["orders"]
    status = dict(zip(o["order"], o["status"])) if o is not None and len(o) else {}
    g = S["logistics"]
    ship = g.groupby("order").agg(amount=("amount", "sum"), charged=("charged", "min")) \
        if g is not None and len(g) else pd.DataFrame(columns=["amount", "charged"])
    bal = S["balance"]
    bal = bal[bal["order"] != ""] if bal is not None and len(bal) else None
    extra = bal.groupby("order")["amount"].sum() if bal is not None and len(bal) else pd.Series(dtype=float)
    near = neighbour_shipping(S["folder"], S["month"])
    L = S["L"]
    wh = L[L["line"].isin(C.WITHHOLDING_LINES)].groupby("doc")["amount"].sum()
    doc_gross = v.assign(_w=v["gross"].abs()).groupby("doc")["_w"].sum()

    v = v.assign(_ret=(v["gross"] < 0).astype(int)).sort_values(["date", "doc", "_ret"], kind="stable")
    # the row that carries the order's shipping: its first sale row, else its first row
    carrier = {}
    for i, r in zip(v.index, v.itertuples(index=False)):
        if r.order not in carrier or (carrier[r.order][1] and r.gross > 0):
            carrier[r.order] = (i, r.gross <= 0)
    rows = []
    for i, r in zip(v.index, v.itertuples(index=False)):
        hit = by_order.get(r.order)
        if hit and hit[0]:
            sku, title, qty, src = hit[0], hit[1], hit[2], SKU_FROM_ORDERS
        elif r.article in by_article:
            (sku, title), qty, src = by_article[r.article], 1.0, SKU_FROM_ARTICLE
        else:
            sku, title, qty, src = "LV-%s" % r.article, "", 1.0, SKU_NONE
        first = carrier[r.order][0] == i
        s_m = -float(ship.loc[r.order, "amount"]) if first and r.order in ship.index else 0.0
        s_n, s_n_date = 0.0, None
        if first and r.order not in ship.index and r.order in near:
            s_n, s_n_date = -near[r.order][0], near[r.order][1]
        x_m = -float(extra[r.order]) if first and r.order in extra.index else 0.0
        w = float(wh.get(r.doc, 0.0))
        share = abs(r.gross) / float(doc_gross[r.doc]) if float(doc_gross.get(r.doc, 0.0)) else 0.0
        rows.append({
            "date": r.date, "doc": r.doc, "order": r.order, "kind": "退货" if r.gross < 0 else "销售", "first": first,
            "store_lv": r.store_lv, "article": r.article, "sku": sku, "title": title,
            "qty": -qty if r.gross < 0 else qty, "status": status.get(r.order, ""), "sku_from": src,
            "gross": float(r.gross), "discount": float(r.discount), "net": float(r.net),
            "commission": -float(r.commission), "card": -(float(r.net) - float(r.commission) - float(r.net_pay)),
            "net_pay": float(r.net_pay), "promo": -float(r.discount), "ship": s_m, "ship_near": s_n, "ship_near_date": s_n_date,
            "ship_date": ship.loc[r.order, "charged"] if s_m else None, "extra": x_m,
            "withheld": w * share})
    rows = pd.DataFrame(rows)
    billed = set(rows["order"])
    od = o.drop_duplicates("order").set_index("order") if o is not None and len(o) else None

    def describe(order):
        hit = by_order.get(order)
        return (hit[0] if hit else "", hit[1] if hit else "",
                od.loc[order, "status"] if od is not None and order in od.index else "",
                od.loc[order, "created"] if od is not None and order in od.index else None)

    ship_only = []
    if g is not None:
        for r in g[~g["order"].isin(billed)].itertuples(index=False):
            sku, title, st, created = describe(r.order)
            ship_only.append({"order": r.order, "charged": r.charged, "amount": -float(r.amount),
                              "carrier": r.carrier, "sku": sku, "title": title, "status": st, "created": created})
    extra_only = []
    full = S["balance"]
    if full is not None:
        for r in full[~full["order"].isin(billed) | (full["order"] == "")].itertuples(index=False):
            sku, title, st, created = describe(r.order) if r.order else ("", "", "", None)
            extra_only.append({"order": r.order, "registered": r.registered, "amount": -float(r.amount),
                               "reason": r.reason, "sku": sku, "title": title, "status": st})
    S["bill"] = {"rows": rows, "ship_only": pd.DataFrame(ship_only), "extra_only": pd.DataFrame(extra_only)}

    pl = m["pl"]
    posted = set(L.loc[L["line"].isin((C.SALES, C.REFUND)), "doc"])
    in_ledger = rows[rows["doc"].isin(posted)]
    other = dict((k, float(pl[k])) for k in (C.LOGISTICS, C.LOGISTICS_EXTRA, C.LOGISTICS_ADJ, C.MSI,
                                              C.SUBSCRIPTION, C.OTHER, C.UNREGISTERED))
    B = {"rows": int(len(rows)), "orders": int(rows["order"].nunique()),
         "sales_n": int((rows["kind"] == "销售").sum()), "returns_n": int((rows["kind"] == "退货").sum()),
         "gross": float(rows["gross"].sum()), "discount": float(rows["discount"].sum()),
         "net": float(rows["net"].sum()), "commission": float(rows["commission"].sum()),
         "card": float(rows["card"].sum()), "net_pay": float(rows["net_pay"].sum()),
         "promo": float(rows["promo"].sum()), "promo_ledger": float(pl[C.PROMO]),
         "ship": float(rows["ship"].sum()), "ship_n": int((rows["ship"] != 0).sum()),
         "ship_near": float(rows["ship_near"].sum()), "ship_near_n": int((rows["ship_near"] != 0).sum()),
         "extra": float(rows["extra"].sum()),
         "ship_only": float(sum(x["amount"] for x in ship_only)), "ship_only_n": len(ship_only),
         "extra_only": float(sum(x["amount"] for x in extra_only)), "extra_only_n": len(extra_only),
         "ship_list": -float(g["amount"].sum()) if g is not None and len(g) else 0.0,
         "no_ship_n": int(((rows["kind"] == "销售") & rows["first"] & (rows["ship"] == 0)
                           & (rows["ship_near"] == 0)).sum()),
         "sku_fallback_n": int((rows["sku_from"] != SKU_FROM_ORDERS).sum()),
         "not_in_ledger_n": int(len(rows) - len(in_ledger)),
         "net_pay_ledger": float(in_ledger["net_pay"].sum()),
         "other": other, "withheld": float(rows["withheld"].sum())}
    B["computed"] = B["net_pay_ledger"] + B["promo_ledger"] + sum(other.values())
    B["resid"] = m["contrib"] - B["computed"]
    m["bill"] = B


def check(S):
    """The bill's tie to ② as a check record (warn: the figures of ② come from the
    ledger and stand on their own; a difference means the bill and the ledger disagree)."""
    B = S["m"].get("bill")
    return K.check("bill_tie", S["store"], "账单逐单合计 ＝ ② 经营贡献毛利",
                   True if not B else abs(B["resid"]) <= TIE_TOL,
                   "没有 Ventas Marketplace（账单）" if not B else
                   "账单净额 %s ＋ 促销补偿 %s ＋ 本月其他费用 %s ＝ %s，② %s，差 %s%s" % (
                       _f(B["net_pay_ledger"]), _f(B["promo_ledger"]), _f(sum(B["other"].values())), _f(B["computed"]),
                       _f(S["m"]["contrib"]), _f(B["resid"]),
                       "；账单里有 %d 行的单据不在本月账本里" % B["not_in_ledger_n"] if B["not_in_ledger_n"] else ""),
                   severity="warn", skipped=not B)


def _f(x):
    return "{:,.2f}".format(float(x))


def _d(x):
    return str(x.date()) if x is not None and pd.notna(x) else ""


_T = F(10)
_CALC = dict(fmt=RF.MONEY, border=True, fill=K.C_CALC)
COLS = [RF.Col("店铺", 18, font=_T), RF.Col("入账日", 11, font=_T), RF.Col("单据号", 12, font=_T),
        RF.Col("订单号", 13, font=_T), RF.Col("类型", 6, font=_T), RF.Col("Liverpool 商品号", 13, font=_T),
        RF.Col("SKU", 18, font=_T), RF.Col("商品", 40, font=_T), RF.Col("件数", 6, font=_T),
        RF.Col("订单状态", 12, font=_T),
        RF.Col("含税售价\nVenta con impuestos", 14, RF.MONEY), RF.Col("Liverpool 折扣\nDescuento", 12, RF.MONEY),
        RF.Col("折后售价\nVenta con descuento", 14, RF.MONEY), RF.Col("佣金\nComisión Dilisa", 12, RF.MONEY),
        RF.Col("外部银行卡附加佣金\n（账单内已扣）", 13, RF.MONEY), RF.Col("账单净额\nPago neto", 14, RF.MONEY),
        RF.Col("Liverpool 促销补偿\n（另单返还卖家）", 13, RF.MONEY),
        RF.Col("发货物流费\n（本月扣）", 12, RF.MONEY), RF.Col("补充物流费\nBalance de cuenta", 12, RF.MONEY),
        RF.Col("扣物流后净额", 14, formula="=P{r}+Q{r}+R{r}+S{r}", **_CALC),
        RF.Col("发货物流费\n（邻月扣，不计入本月）", 13, RF.MONEY), RF.Col("物流扣费日", 11, font=_T),
        RF.Col("代扣税\n（备查，不计入）", 12, RF.MONEY), RF.Col("SKU 来源", 16, font=F(9, False, K.C_GREY))]
N = len(COLS)
SUBTITLE = ("以 Liverpool 的结算账单（Ventas Marketplace）为底：账单的每一行在这里一行，按订单号对上发货物流费"
            "（Costos logísticos）、补充物流费（Balance de cuenta）和 SKU（订单列表）。扣项为负数；金额 MXN，含 IVA。"
            "账单净额 ＝ 折后售价 − 佣金 − 外部银行卡附加佣金；Liverpool 自己出资的折扣另单返还卖家。表后列出对不到账单的物流费，以及和②的勾稽。")


def sheet(wb, fam, ctx):
    ws = fam.table_sheet(wb, SHEET, "账单逐单明细（%s）" % ctx["period_cn"], SUBTITLE, COLS)
    ws.row_dimensions[4].height = 44
    r = 5
    for S in fam.stores:
        bill, B, st = S.get("bill"), S["m"].get("bill"), S["store"]
        if not bill:
            ws.cell(row=r, column=1, value="%s：没有 Ventas Marketplace（账单），本页无法生成" % st).font = \
                F(10, True, K.C_RED)
            r += 2
            continue
        first = r
        for x in bill["rows"].itertuples(index=False):
            cells = fam.write_row(ws, r, COLS, [
                st, _d(x.date), x.doc, x.order, x.kind, x.article, x.sku, x.title, x.qty, x.status,
                x.gross, x.discount, x.net, x.commission, x.card, x.net_pay, x.promo, x.ship, x.extra, None,
                x.ship_near, _d(x.ship_date) or _d(x.ship_near_date), x.withheld, x.sku_from])
            if x.kind == "退货":
                for c in cells[:10]:
                    c.font = F(10, False, K.C_RED)
            if x.sku_from != SKU_FROM_ORDERS:
                cells[6].fill = fill(K.C_WARN)
            r += 1
        last = r - 1
        ws.cell(row=r, column=1, value="%s 合计" % st).font = F(10, True)
        ws.cell(row=r, column=4, value="%d 行，%d 个订单（销售 %d 行，退货 %d 行）" % (
            B["rows"], B["orders"], B["sales_n"], B["returns_n"])).font = F(9, False, K.C_GREY)
        for j in list(range(11, 22)) + [23]:
            col = get_column_letter(j)
            c = ws.cell(row=r, column=j, value="=SUM(%s%d:%s%d)" % (col, first, col, last))
            c.number_format, c.font, c.fill, c.border = fam.money, F(10, True), fill(K.C_TOTAL), K.BOX
        r += 2
        r = _tie(ws, fam, r, S, B)
        r = _leftover(ws, fam, r, st, "本月扣的发货物流费，订单不在本月账单里（多为上月入账的订单、退货取回）",
                      ["订单号", "扣费日", "金额", "承运商", "SKU", "商品", "订单状态", "下单日"],
                      [(x.order, _d(x.charged), x.amount, x.carrier, x.sku, x.title, x.status, _d(x.created))
                       for x in bill["ship_only"].itertuples(index=False)], B["ship_only"])
        r = _leftover(ws, fam, r, st, "本月的补充物流费与其他扣款（Balance de cuenta），订单不在本月账单里（DIGITAL 是订阅费，在②的订阅费一行）",
                      ["订单号（Remisión）", "登记日", "金额", "原因（Motivo）", "SKU", "商品", "订单状态", ""],
                      [(x.order, _d(x.registered), x.amount, x.reason, x.sku, x.title, x.status, "")
                       for x in bill["extra_only"].itertuples(index=False)], B["extra_only"])
        r += 1
    ws.freeze_panes = "E5"


def _tie(ws, fam, r, S, B):
    """Bill net -> ②'s contribution, line by line."""
    r = RF.section(ws, r, "%s：从账单净额到 ② 经营贡献毛利" % S["store"], 10)
    o = B["other"]
    ledger_only = o[C.LOGISTICS] - B["ship"] - B["ship_only"]
    lines = [
        ("账单净额合计（Pago neto）", B["net_pay_ledger"],
         "＝ 折后售价 %s − 佣金 %s − 外部银行卡附加佣金 %s%s" % (
             _f(B["net"]), _f(-B["commission"]), _f(-B["card"]),
             "；另有 %d 行的单据不在本月账本里，未计入" % B["not_in_ledger_n"] if B["not_in_ledger_n"] else ""), False),
        ("加：Liverpool 出资的促销补偿（账本，另一张单据返还）", B["promo_ledger"],
         "账单里的 Liverpool 折扣合计 %s：买家少付的部分由 Liverpool 补给卖家，不在账单净额里" % _f(B["discount"]), False),
        ("减：发货物流费（本月账本）", o[C.LOGISTICS],
         "逐单对到账单 %s（%d 单）；订单不在本月账单 %s（%d 单，见下）；账本有、物流费明细里没有 %s" % (
             _f(B["ship"]), B["ship_n"], _f(B["ship_only"]), B["ship_only_n"], _f(ledger_only)), False),
        ("减：补充物流费", o[C.LOGISTICS_EXTRA],
         "逐单对到账单 %s；其余 %d 笔见下" % (_f(B["extra"]), B["extra_only_n"]), False),
        ("加/减：物流费调整（往月）", o[C.LOGISTICS_ADJ], "", False),
        ("减：无息分期（MSI）成本", o[C.MSI], "", False),
        ("减：订阅费（Cuota digital）", o[C.SUBSCRIPTION], "", False),
        ("减：其他调整与未登记科目", o[C.OTHER] + o[C.UNREGISTERED], "", False),
        ("＝ 按账单算出的经营贡献毛利", B["computed"], "", True),
        ("② 经营贡献毛利（账本）", S["m"]["contrib"], "", False),
        ("差额（应为 0）", B["resid"], "账单按单四舍五入外部银行卡附加佣金，允许 ±%s" % _f(TIE_TOL), True),
        ("备查：代扣税（IVA 8% ＋ ISR 2.5%，预缴税款，不计入贡献）", S["m"]["pl"]["withheld"], "", False),
        ("备查：销售订单里本月和邻月都没有发货物流费的", B["no_ship_n"],
         "单；多为已取消、已退款或门店自提的订单", False),
        ("备查：邻月扣的发货物流费（不计入本月）", B["ship_near"], "%d 单" % B["ship_near_n"], False),
    ]
    for label, value, note, bold in lines:
        ws.cell(row=r, column=1, value=label).font = F(10, bold)
        c = ws.cell(row=r, column=11, value=value)
        c.border = K.BOX
        c.font = F(10, bold)
        if isinstance(value, float):
            c.number_format = fam.money
        if bold:
            c.fill = fill(K.C_TOTAL if "差额" in label else K.C_SUB)
        ws.cell(row=r, column=12, value=note).font = F(9, False, K.C_GREY)
        r += 1
    return r + 1


def _leftover(ws, fam, r, store, title, headers, rows, total):
    r = RF.section(ws, r, "%s：%s" % (store, title), 10)
    if not rows:
        ws.cell(row=r, column=1, value="没有。").font = F(10, False, K.C_GREY)
        return r + 2
    at = (4, 2, 11, 6, 7, 8, 10, 3)            # the main table's columns these values sit under
    for j, h in zip(at, headers):
        if h:
            c = ws.cell(row=r, column=j, value=h)
            c.font, c.fill, c.border = F(9, True), fill(K.C_SUB), K.BOX
    r += 1
    for vals in rows:
        for j, v in zip(at, vals):
            c = ws.cell(row=r, column=j, value=v)
            c.font = _T
            if j == 11:
                c.number_format, c.border = fam.money, K.BOX
        r += 1
    ws.cell(row=r, column=1, value="小计").font = F(10, True)
    c = ws.cell(row=r, column=11, value=total)
    c.number_format, c.font, c.fill, c.border = fam.money, F(10, True), fill(K.C_TOTAL), K.BOX
    return r + 2
