# -*- coding: utf-8 -*-
"""Reventa ("vendido por Mercado Libre") — what both report bases share.

MercadoLibre Mexico moves invited listings to Reventa: at the buyer's click the
seller sells the item to MercadoLibre, who resells it. Seen on EWTTO_SM, whose
whole store switched on 2026-10-26 (platforms/mercadolibre/README.md, "Reventa").
In the sales export a Reventa row differs from a direct sale:

  "Modelo de venta"                 "Reventa" (a direct sale says "Venta directa")
  "Ingresos por productos"          the net amount MercadoLibre pays the seller
  "Cargo por venta e impuestos",
  "Costos de envío"                 0 — no commission, shipping or withholding
  the unit-price column             the seller's suggested price; the header became
                                    "Precio unitario para venta directa / Precio
                                    sugerido para reventa" for every store
  the monthly bill                  no "Cargo por venta" / "Cargo por envíos" row

Here: the two header facts (model column, unit-price column), and the per-SKU
comparison table appended to ⑦ when a month has Reventa sales. Each basis reads
its own files and decides its own accounting; this module only knows the
export's columns and writes the table.

This file exists twice, byte for byte: report/reventa.py (order basis) and
bill_report/reventa.py (bill basis). The two bases never import each other or a
common MercadoLibre module (tests/test_ml_bill_report.py), so each has its own
copy; tests/test_ml_reventa.py fails when the copies differ. Change both.
"""
import collections

from openpyxl.utils import get_column_letter

from core import report_kit as K

MODEL_HEADER = "Modelo de venta"
# the unit-price header: until 2026-10 / since Reventa (same column, renamed)
UNIT_PRICE_HEADS = ("Precio unitario de venta", "Precio unitario para venta directa")

LABEL_CN = "Reventa（平台买断转售）"
DIRECT_CN = "直接销售"


def is_reventa(value):
    """True for the "Modelo de venta" value of a Reventa sale."""
    return str(value if value is not None else "").strip().lower() == "reventa"


def unit_price_column(columns):
    """The export's unit-price header among `columns` (old or new name), or None."""
    for head in UNIT_PRICE_HEADS:
        for c in columns:
            if str(c).startswith(head):
                return c
    return None


# ------------------------------------------------------------------ per-SKU comparison

Row = collections.namedtuple("Row", "reventa sku title units list_value income deductions refund")
Row.__doc__ = """One sales-export row of the month.
list_value   units × unit price (direct: the buyer's price; Reventa: the suggested price)
income       "Ingresos por productos" (direct: the buyer's price; Reventa: what ML pays)
deductions   direct: charge and taxes + shipping cost + shipping income (negative = cost);
             Reventa: 0
refund       "Anulaciones y reembolsos" (≤ 0)"""


def compare(rows, tol=0.005):
    """{sku: {...}} over the rows **without a refund** (a cancelled or returned row
    would distort the per-unit figures), plus the count of rows left out.
    Keys per SKU: title, d_units, d_list, d_ded, d_net, r_units, r_list, r_net."""
    out = collections.OrderedDict()
    skipped = 0
    for r in rows:
        if abs(r.refund or 0.0) > tol:
            skipped += 1
            continue
        x = out.setdefault(r.sku or "（无 SKU）", collections.defaultdict(float))
        if r.title and not x.get("title"):
            x["title"] = r.title
        if r.reventa:
            x["r_units"] += r.units or 0.0
            x["r_list"] += r.list_value or 0.0
            x["r_net"] += r.income or 0.0
        else:
            x["d_units"] += r.units or 0.0
            x["d_list"] += r.list_value or 0.0
            x["d_ded"] += r.deductions or 0.0
            x["d_net"] += (r.income or 0.0) + (r.deductions or 0.0)
    return out, skipped


COMPARE_HEAD = ["SKU", "商品",
                "直接：件数", "直接：标价销售额", "直接：平台扣款（佣金＋代扣税＋运费）", "直接：到手净额",
                "直接：每件到手", "直接：到手 ÷ 标价",
                "Reventa：件数", "Reventa：建议售价合计", "Reventa：平台支付净额", "Reventa：每件到手",
                "Reventa：到手 ÷ 建议售价", "每件到手变化（Reventa ÷ 直接 − 1）"]
COMPARE_TITLE = "Reventa 与直接销售对比（按 SKU，本月下单、没有取消或退款的订单行）"
COMPARE_NOTE = ("来源：销售报表 Ventas 的“Modelo de venta”列。直接销售：标价 ＝ 买家价（Ingresos por productos），"
                "到手 ＝ 标价 − 佣金及代扣税 − 运费（销售报表口径，代扣税以后可抵减）。Reventa：平台在买家下单时向本店买断，"
                "“Ingresos por productos”就是平台付给本店的净额，没有佣金、运费和代扣税；建议售价 ＝ 件数 × "
                "“Precio sugerido para reventa”。广告、Full 仓储等店铺级费用两种方式都照常收取，不在本表。%s")


def write_compare(ws, r, table, skipped, money, first_col=1):
    """Write the comparison at row r of ws. Returns the next free row."""
    F, fill = K.F, K.fill
    n = len(COMPARE_HEAD)
    c = ws.cell(row=r, column=first_col, value=COMPARE_TITLE)
    c.font = F(11, True, "FFFFFF")
    for j in range(first_col, first_col + n):
        ws.cell(row=r, column=j).fill = fill(K.C_BAND)
    r += 1
    ws.cell(row=r, column=first_col, value=COMPARE_NOTE % (
        "有取消或退款的 %d 行不在本表。" % skipped if skipped else "")).font = F(9, False, K.C_GREY)
    r += 1
    K.header_cells(ws, r, COMPARE_HEAD, first_col=first_col)
    ws.row_dimensions[r].height = 42
    r += 1
    first = r
    L = lambda j: get_column_letter(first_col + j)

    def line(r, vals, bold=False):
        fmts = [None, None, K.INT, money, money, money, money, K.PCT, K.INT, money, money, money, K.PCT, K.PCT]
        row = {0: vals[0], 1: vals[1], 2: vals[2], 3: vals[3], 4: vals[4], 5: vals[5],
               6: "=IF(%s%d=0,0,%s%d/%s%d)" % (L(2), r, L(5), r, L(2), r),
               7: "=IF(%s%d=0,0,%s%d/%s%d)" % (L(3), r, L(5), r, L(3), r),
               8: vals[6], 9: vals[7], 10: vals[8],
               11: "=IF(%s%d=0,0,%s%d/%s%d)" % (L(8), r, L(10), r, L(8), r),
               12: "=IF(%s%d=0,0,%s%d/%s%d)" % (L(9), r, L(10), r, L(9), r),
               13: '=IF(OR(%s%d=0,%s%d=0),0,%s%d/%s%d-1)' % (L(6), r, L(11), r, L(11), r, L(6), r)}
        if bold:
            row[13] = None          # a store-wide change would only show the product mix
        for j in range(n):
            c = ws.cell(row=r, column=first_col + j, value=row[j])
            c.border = K.BOX
            if fmts[j]:
                c.number_format = fmts[j]
            if bold:
                c.font = F(10, True)
                c.fill = fill(K.C_TOTAL)
            elif j in (6, 7, 11, 12, 13):
                c.fill = fill(K.C_CALC)

    order = sorted(table, key=lambda k: -(table[k]["r_net"] + table[k]["d_net"]))
    for k in order:
        x = table[k]
        line(r, [k, x.get("title", ""), x["d_units"], x["d_list"], x["d_ded"], x["d_net"],
                 x["r_units"], x["r_list"], x["r_net"]])
        r += 1
    if order:
        s = lambda j: "=SUM(%s%d:%s%d)" % (L(j), first, L(j), r - 1)
        line(r, ["合计", "", s(2), s(3), s(4), s(5), s(8), s(9), s(10)], bold=True)
        r += 1
    return r + 1
