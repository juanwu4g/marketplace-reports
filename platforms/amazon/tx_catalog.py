# -*- coding: utf-8 -*-
"""
Every line of Amazon Mexico's monthly transaction report ("Transacciones",
FORMS.md §1) -> where each of its money columns goes in the monthly report.

A row of that report is identified by `tipo` + `descripción` (for order and
refund rows the description is the product name, so only `tipo` counts). Each
row spreads its `total` over up to 15 money columns; this catalogue says, per
row kind and per column:

    (P&L line or memo line, IVA mode)

IVA modes — how much of the amount is IVA (the rate is the site's, 16 %):

    NET   the amount is already net of IVA (Amazon split it out in another column,
          or it carries no IVA: sales, shipping, promotions, reimbursements)
    INCL  IVA-inclusive: net = amount / 1.16, the rest is input IVA (Amazon's fee
          columns: referral, FBA, storage, removal, inbound transport, subscription)
    IVA   the whole amount is input IVA (the `otro` column of an advertising row)

Lines are the fixed rows of ② (report_build.PL_ROWS) or the memo lines that are
not profit and loss (output IVA, withholdings, payouts) — see LINES.

A row kind that is not here, or a known row with money in a column the catalogue
does not expect, is an **error** check in the report: the amount goes to the
visible "未登记科目 ⚠" row (so every bridge still ties) and the check lists the
rows. A new Amazon charge never vanishes silently into "other".

Spanish is primary: the downloader switches the UI back to Spanish before it
requests the report. The English names (as an English UI writes them) are kept
next to each entry so a person can find them, not for matching.
"""
import collections
import re
import unicodedata

# ---------------------------------------------------------------- lines

# P&L lines (keys of report_build.PL_ROWS)
SALES = "sales"                 # 商品销售收入
REFUND = "refund"               # 退货退款
SHIPPING = "shipping"           # 运费收入
GIFTWRAP = "giftwrap"           # 礼品包装收入
PROMO = "promo"                 # 促销折扣
OTHER_INCOME = "other_income"   # FBA 库存赔付
COMMISSION = "commission"       # 销售佣金
FBA_FEE = "fba_fee"             # FBA 配送费
STORAGE = "storage"             # FBA 月度仓储费
STORAGE_LT = "storage_lt"       # FBA 长期仓储费
LOGISTICS = "logistics_other"   # 入仓运输、移除/弃置等其他物流费
ADS = "ads"                     # 广告费
SUBSCRIPTION = "subscription"   # 订阅费
OTHER_FEES = "other_fees"       # 其他平台费用
UNREGISTERED = "unregistered"   # 未登记 ⚠

# Not profit and loss: memo lines (② 备查项) and the cash bridge (③ 桥 C)
IVA_OUT = "iva_out"             # 销项 IVA（商品、运费、礼品包装、促销所含）
IVA_WITHHELD = "iva_withheld"   # Amazon 代扣 IVA（预缴）
ISR = "isr"                     # Amazon 代扣 ISR（上月的，本月从余额扣）
IVA_ADJ = "iva_adj"             # IVA 调整（Ajuste del IVA del producto）
PAYOUT = "payout"               # 转账到银行

INCOME_LINES = (SALES, REFUND, SHIPPING, GIFTWRAP, PROMO)
FEE_LINES = (COMMISSION, FBA_FEE, STORAGE, STORAGE_LT, LOGISTICS, ADS, SUBSCRIPTION,
             OTHER_FEES, UNREGISTERED)
PL_LINES = INCOME_LINES + FEE_LINES + (OTHER_INCOME,)
TAX_LINES = (IVA_OUT, IVA_WITHHELD, ISR, IVA_ADJ)
CASH_LINES = (PAYOUT,)
LINES = PL_LINES + TAX_LINES + CASH_LINES

# line -> family category (docs/REPORT_FAMILY.md §5, plus the revenue side)
LINE_CATEGORY = {
    SALES: "sales", REFUND: "refund", SHIPPING: "shipping_income", GIFTWRAP: "sales",
    PROMO: "promotion", OTHER_INCOME: "other_income", COMMISSION: "commission",
    FBA_FEE: "fulfilment_storage", STORAGE: "fulfilment_storage",
    STORAGE_LT: "fulfilment_storage", LOGISTICS: "fulfilment_storage", ADS: "ads",
    SUBSCRIPTION: "subscription", OTHER_FEES: "other", UNREGISTERED: "⚠ 未登记",
    IVA_OUT: "tax", IVA_WITHHELD: "tax_withheld", ISR: "tax_withheld", IVA_ADJ: "tax",
    PAYOUT: "payout",
}

# line -> group of Amazon's monthly summary PDF ("Resumen"), for ③ bridge D.
# FBA inventory reimbursements are split by sign there (credits under Ingresos as
# "Crédito de inventario del FBA", clawbacks under Gastos as "Ajustes"): SIGNED.
SIGNED = "signed"
SUMMARY_GROUP = dict([(k, "ingresos") for k in INCOME_LINES] +
                     [(k, "gastos") for k in FEE_LINES] +
                     [(k, "impuesto") for k in TAX_LINES] +
                     [(PAYOUT, "transferencias"), (OTHER_INCOME, SIGNED)])
SUMMARY_GROUPS = ("ingresos", "gastos", "transferencias", "impuesto")

NET, INCL, IVA = "net", "incl", "iva"

# ---------------------------------------------------------------- columns

# The money columns of the report, in file order (`total` is their sum).
MONEY_COLS = [
    "ventas de productos", "impuesto de ventas de productos", "créditos de envío",
    "impuesto de abono de envío", "créditos por envoltorio de regalo",
    "impuesto de créditos de envoltura", "Tarifa reglamentaria",
    "Impuesto sobre tarifa reglamentaria", "descuentos promocionales",
    "impuesto de reembolsos promocionales", "impuesto de retenciones en la plataforma",
    "tarifas de venta", "tarifas fba", "tarifas de otra transacción", "otro",
]
TOTAL_COL = "total"
# Spanish column -> Chinese, for ④
COL_ZH = {
    "ventas de productos": "商品价款（不含 IVA）",
    "impuesto de ventas de productos": "商品销项 IVA",
    "créditos de envío": "买家支付的运费（不含 IVA）",
    "impuesto de abono de envío": "运费销项 IVA",
    "créditos por envoltorio de regalo": "礼品包装费（不含 IVA）",
    "impuesto de créditos de envoltura": "礼品包装销项 IVA",
    "Tarifa reglamentaria": "监管费",
    "Impuesto sobre tarifa reglamentaria": "监管费所含 IVA",
    "descuentos promocionales": "促销折扣（不含 IVA）",
    "impuesto de reembolsos promocionales": "促销折扣冲减的销项 IVA",
    "impuesto de retenciones en la plataforma": "Amazon 代扣 IVA",
    "tarifas de venta": "销售佣金（含 IVA）",
    "tarifas fba": "FBA 费用（含 IVA）",
    "tarifas de otra transacción": "其他交易费",
    "otro": "其他",
}

# What an order line (`Pedido`) does with each column. A refund line (`Reembolso`)
# is the same with the product price going to the refund row. Fee refunds net in
# the fee's own row; the refund administration fee Amazon keeps is inside
# "tarifas de venta" of the refund line, so it nets in the commission row.
_ORDER_COLS = {
    "ventas de productos": (SALES, NET),
    "impuesto de ventas de productos": (IVA_OUT, NET),
    "créditos de envío": (SHIPPING, NET),
    "impuesto de abono de envío": (IVA_OUT, NET),
    "créditos por envoltorio de regalo": (GIFTWRAP, NET),
    "impuesto de créditos de envoltura": (IVA_OUT, NET),
    # Not seen in Mexico yet (always 0 in 2026). Amazon prints the fee and its tax in
    # separate columns, so the fee is taken as net and its tax as input IVA.
    "Tarifa reglamentaria": (OTHER_FEES, NET),
    "Impuesto sobre tarifa reglamentaria": (OTHER_FEES, IVA),
    "descuentos promocionales": (PROMO, NET),
    "impuesto de reembolsos promocionales": (IVA_OUT, NET),
    "impuesto de retenciones en la plataforma": (IVA_WITHHELD, NET),
    "tarifas de venta": (COMMISSION, INCL),
    "tarifas fba": (FBA_FEE, INCL),
    # "Incluye reintegros de envío, retenciones de envío y tarifas de cobro de
    # impuestos de venta" — 0 on every order line so far; a fee, IVA-inclusive.
    "tarifas de otra transacción": (OTHER_FEES, INCL),
}
_REFUND_COLS = dict(_ORDER_COLS, **{"ventas de productos": (REFUND, NET)})


Entry = collections.namedtuple("Entry", "key tipo pattern zh en line mode cols")
Entry.__doc__ = """One kind of transaction line.

key      stable id (④ groups by it)
tipo     Amazon's `tipo`
pattern  regex on the normalised `descripción` (None: any description)
zh / en  Chinese name / the English UI's name (for people, not matching)
line     the line every money column goes to (when cols is None)
mode     the IVA mode of that line
cols     {column: (line, mode)} for rows that spread over several columns; a column
         not listed is unregistered
"""

CATALOG = []


def _add(key, tipo, pattern, zh, en, line=None, mode=NET, cols=None):
    CATALOG.append(Entry(key, tipo, re.compile(pattern) if pattern else None, zh, en,
                         line, mode, cols))


def norm(s):
    """lower case, no accents, single spaces: 'Tarifa  por almacenamiento' and
    'tarifa por almacenamiento' are the same."""
    if s is None:
        return ""
    s = str(s)
    if s.lower() == "nan":
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.lower().split())


# ---------------------------------------------------------------- the catalogue
# Order matters: the first entry whose tipo and pattern match wins.

_add("order", "Pedido", None, "订单", "Order", cols=_ORDER_COLS)
_add("refund", "Reembolso", None, "退款", "Refund", cols=_REFUND_COLS)

# Ajuste — FBA inventory reimbursements; the reason is in brackets
_add("fba_reimbursement", "Ajuste", r"^reembolso de inventario de logistica de amazon",
     "FBA 库存赔付", "FBA Inventory Reimbursement", OTHER_INCOME, NET)
# Ajuste without description: ISR withholding / IVA adjustment (classify())
_add("isr", "Ajuste", r"^$", "Amazon 代扣 ISR（上月）", "Income tax withheld", ISR, NET)
_add("iva_adj", "Ajuste", r"^$", "IVA 调整（Ajuste del IVA del producto）",
     "VAT adjustment", IVA_ADJ, NET)

# Tarifa de servicio
_add("ads", "Tarifa de servicio", r"^costo de la publicidad", "广告费", "Cost of Advertising",
     cols={"tarifas de otra transacción": (ADS, NET), "otro": (ADS, IVA)})
_add("subscription", "Tarifa de servicio", r"^suscripcion", "专业卖家订阅费", "Subscription",
     SUBSCRIPTION, INCL)

# Tarifas de transacciones de Logística de Amazon
_add("storage", "Tarifas de transacciones de Logística de Amazon",
     r"^tarifa por almacenamiento de inventario de fba", "FBA 月度仓储费",
     "FBA storage fee", STORAGE, INCL)
_add("storage_lt", "Tarifas de transacciones de Logística de Amazon",
     r"^tarifa por almacenamiento prolongado de fba", "FBA 长期仓储费",
     "FBA long-term storage fee", STORAGE_LT, INCL)
_add("removal", "Tarifas de transacciones de Logística de Amazon",
     r"^solicitud de retiro de inventario", "FBA 库存移除/弃置费",
     "FBA removal / disposal fee", LOGISTICS, INCL)
_add("partnered_carrier", "Tarifas de transacciones de Logística de Amazon",
     r"^tarifa de envio con transportista asociado", "FBA 合作承运人入仓运费",
     "FBA inbound partnered carrier fee", LOGISTICS, INCL)

# Tarifa de inventario FBA
_add("inbound_transport", "Tarifa de inventario FBA",
     r"^cargos por transporte de envios a amazon", "FBA 入仓运输费",
     "FBA inbound transportation fee", LOGISTICS, INCL)

# Trasferir (sic) — payout to the seller's bank account
_add("payout", "Trasferir", r"^para la cuenta que termina en", "转账到银行账户",
     "Transfer to bank account", PAYOUT, NET)

BY_KEY = dict((e.key, e) for e in CATALOG)

# An Ajuste without description and without SKU larger than this (MXN) is the ISR
# withholding Amazon takes around the 11th for the previous month (Aug 11 2026:
# −12 886.33, Sep 11: −11 814.82); smaller ones are "Ajuste del IVA del producto"
# (+2.45, +2.69). ③ bridge D checks the split against the summary PDF.
ISR_MIN = 100.0

# The reasons of FBA inventory reimbursements -> Chinese (⑥)
REIMBURSE_REASON_ZH = {
    "devolucion de cliente": "买家退货后未退回仓库 / 退回不可售",
    "extraviado en el almacen": "仓内丢失",
    "danado en el almacen": "仓内损坏",
    "extraviado: envio entrante": "入库途中丢失",
    "incidente de atencion al cliente": "客服事件赔付",
    "ajuste general": "一般调整（多为冲回此前的赔付）",
}


def reimburse_reason(desc):
    """'Reembolso de inventario de Logística de Amazon (Devolución de cliente)'
    -> ('Devolución de cliente', '买家退货后未退回仓库 / 退回不可售')."""
    m = re.search(r"\(([^)]*)\)\s*$", str(desc or ""))
    reason = m.group(1).strip() if m else ""
    return reason, REIMBURSE_REASON_ZH.get(norm(reason), "（未识别的原因）")


def classify(tipo, desc, sku=None, total=0.0):
    """The Entry of one transaction line, or None (unregistered)."""
    t, d = norm(tipo), norm(desc)
    if t == norm("Ajuste") and not d:
        if norm(sku):
            return None                       # an adjustment of a product without reason
        return BY_KEY["isr"] if abs(float(total or 0.0)) > ISR_MIN else BY_KEY["iva_adj"]
    for e in CATALOG:
        if e.key in ("isr", "iva_adj"):
            continue
        if norm(e.tipo) != t:
            continue
        if e.pattern is None or e.pattern.search(d):
            return e
    return None


def column_target(entry, col):
    """(line, mode) of one money column of a line of this kind; unregistered when
    the catalogue does not expect money in that column."""
    if entry is None:
        return UNREGISTERED, NET
    if entry.cols is None:
        return entry.line, entry.mode
    return entry.cols.get(col, (UNREGISTERED, NET))


def split(amount, mode, rate):
    """(net, iva) of an amount in an IVA mode, unrounded: the report rounds the
    net of each P&L line and takes the line's IVA as the rest (report_build.derive)."""
    if mode == INCL:
        net = amount / (1.0 + rate)
        return net, amount - net
    if mode == IVA:
        return 0.0, amount
    return amount, 0.0
