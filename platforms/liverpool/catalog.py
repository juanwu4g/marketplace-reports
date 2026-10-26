# -*- coding: utf-8 -*-
"""Every kind of line of Liverpool's supplier ledger ("Saldos Proveedor", the
"Cobros y Pagos Marketplace" export), mapped to a row of the report.

Liverpool keeps two accounts per marketplace supplier (FORMS.md):

  20040002  what Liverpool owes the seller: the sale price of every sale ("F/",
            credit), the taxes it withholds on it (Q6 IVA 8 %, Q8 ISR 2.5 % of the
            price net of IVA, debits), returns ("DEV", debit) with the withholding
            and the commission given back, logistics adjustments ("RW"), and the
            payments ("ZP")
  11900003  what the seller owes Liverpool: the commission of every sale ("Comisión %
            DILISA") and the surcharge for a non-Liverpool card ("Sobrecomis %
            T.Exter"), the shipping charges ("Rec Logística"), the extra logistics
            ("Balance logístico"), adjustments, the subscription, and the payments
            ("ZP", netted with the first account's)

Sign: a credit (negative) on either account is money for the seller, a debit
(positive) money for Liverpool. Every report amount is therefore -1 × the ledger
amount ("Importe Debe/Haber"), with IVA as Liverpool charges it.

lookup(account, doc_class, text, tax_code) -> (line, zh) or None (unregistered).
"""
import re

PAYABLE = "20040002"        # Liverpool owes the seller
RECEIVABLE = "11900003"     # the seller owes Liverpool

SALES = "sales"
PROMO = "promo"
REFUND = "refund"
COMMISSION = "commission"
CARD_FEE = "card_fee"
LOGISTICS = "logistics"
LOGISTICS_EXTRA = "logistics_extra"
LOGISTICS_ADJ = "logistics_adj"
MSI = "msi"
SUBSCRIPTION = "subscription"
OTHER = "other"
WH_IVA = "wh_iva"
WH_ISR = "wh_isr"
PAYMENT = "payment"
UNREGISTERED = "unregistered"

INCOME_LINES = (SALES, PROMO, REFUND)
FEE_LINES = (COMMISSION, CARD_FEE, LOGISTICS, LOGISTICS_EXTRA, LOGISTICS_ADJ, MSI, SUBSCRIPTION,
             OTHER, UNREGISTERED)
WITHHOLDING_LINES = (WH_IVA, WH_ISR)
# lines that belong to one sale document (and so to its order and SKU)
SALE_LINES = (SALES, PROMO, REFUND, COMMISSION, CARD_FEE, WH_IVA, WH_ISR)

# category of each line, for ④
LINE_CATEGORY = {SALES: "收入", PROMO: "收入", REFUND: "收入（冲减）", COMMISSION: "佣金",
                 CARD_FEE: "佣金", LOGISTICS: "物流", LOGISTICS_EXTRA: "物流", LOGISTICS_ADJ: "物流",
                 MSI: "支付成本", SUBSCRIPTION: "订阅", OTHER: "其他", WH_IVA: "代扣税金",
                 WH_ISR: "代扣税金", PAYMENT: "结算款"}

# (account, doc classes or None, text regex, tax codes or None, line, 中文, original term)
# First match wins. The text is matched after the leading "<store>/" of the
# commission lines (the Liverpool store that sold) is removed.
RULES = [
    (PAYABLE, ("ZP",), r"", None, PAYMENT, "结算款（付给卖家 / 被银行退回）", "ZP · Pago"),
    (RECEIVABLE, ("ZP",), r"", None, PAYMENT, "结算款（从付款中抵扣的费用 / 被银行退回）", "ZP · Pago"),
    (PAYABLE, ("YV",), r"^F/", ("D2",), SALES, "商品销售（买家实付，含 IVA）", "F/ · Venta"),
    (PAYABLE, ("YV",), r"^F/", ("C2",), PROMO, "Liverpool 出资的促销补偿", "F/ · Descuento Liverpool"),
    (PAYABLE, ("YV", "YD"), r"^(COM|F/)", ("Q6",), WH_IVA, "代扣 IVA（8%）", "Q6 · IVA retenido"),
    (PAYABLE, ("YV", "YD"), r"^(COM|F/)", ("Q8",), WH_ISR, "代扣 ISR（2.5%）", "Q8 · ISR retenido"),
    (PAYABLE, ("YD",), r"^DEV", None, REFUND, "退货退款", "DEV · Devolución"),
    (PAYABLE, ("YD",), r"^F/", ("C5",), COMMISSION, "退货退还的佣金", "F/ · Comisión devuelta"),
    (PAYABLE, ("RW",), r"^Ajuste log", None, LOGISTICS_ADJ, "物流费调整（往月）", "RW · Ajuste log"),
    (RECEIVABLE, ("YV",), r"^Comisi[oó]n % DILISA", None, COMMISSION, "销售佣金", "Comisión % DILISA"),
    (RECEIVABLE, ("YV",), r"^Sobrecomis % T\.?\s*Exter", None, CARD_FEE, "外部银行卡附加佣金",
     "Sobrecomis % T.Exter"),
    (RECEIVABLE, ("DY",), r"^Rec Log[ií]stica", None, LOGISTICS, "发货物流费", "Rec Logística"),
    (RECEIVABLE, ("DY",), r"^Balance log[ií]stico", None, LOGISTICS_EXTRA, "补充物流费",
     "Balance logístico (Cargo logístico complementario)"),
    (RECEIVABLE, ("DY",), r"^AJUSTE MSI", None, MSI, "无息分期（MSI）成本调整", "Ajuste MSI"),
    (RECEIVABLE, ("DY",), r"^Ajuste log", None, LOGISTICS_ADJ, "物流费调整（往月）", "Ajuste log"),
    (RECEIVABLE, ("DY",), r"^Suscripci[oó]n", None, SUBSCRIPTION, "订阅费（Cuota digital）",
     "Suscripción MKP"),
]
_COMPILED = [(a, c, re.compile(t, re.I), tc, line, zh, orig) for a, c, t, tc, line, zh, orig in RULES]


def clean_text(text):
    """'0028/Comisión % DILISA      0028' -> 'Comisión % DILISA'; spaces squeezed."""
    t = re.sub(r"\s+", " ", str(text or "")).strip()
    t = re.sub(r"^\d{3,4}/", "", t)
    return re.sub(r" \d{4}$", "", t).strip()


def lookup(account, doc_class, text, tax_code):
    """(line, 中文, original term) of one ledger line, or None when no rule matches."""
    account = str(account or "").strip()
    doc_class = str(doc_class or "").strip()
    tax_code = str(tax_code or "").strip()
    t = clean_text(text)
    for a, classes, rx, codes, line, zh, orig in _COMPILED:
        if a != account or (classes and doc_class not in classes):
            continue
        if codes and tax_code not in codes:
            continue
        if rx.search(t):
            return line, zh, orig
    return None
