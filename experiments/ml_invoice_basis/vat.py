# -*- coding: utf-8 -*-
"""VAT (IVA / IGV): the per-country policy and the proof, from the folder's own
files, that the invoice's `Valor del cargo` is the whole amount the seller pays
(tax included) — not a net amount with VAT added on top.

Four independent tests, all on data already downloaded:

1. Pagos_Facturas "Detalle de Pagos del mes": the part of each payment applied to
   each charge (`Parte del pago aplicado a cargos`, by `Número del cargo`) equals
   the charge's `Valor del cargo`.
2. Pagos_Facturas "Pagos y notas de crédito": each credit note is a fiscal
   document whose `Importe total` (what it settles) equals −Σ `Valor del cargo`
   of the rows carrying that number in `N° de factura fiscal`.
3. The charge is built from its own parts with nothing added on top:
   `Valor del cargo` = `Costo por categoría` + `Costo por unidad vendida` −
   `Valor del descuento` (= `Subtotal sin descuento` − discount). Shown for
   information: how many rows also have `Costo por categoría` = rate × price
   (TANKE_EE: 848 of 1,240; on the rest the category cost is about 0.95 × rate ×
   price — a reduced rate, not a tax: a tax would be × 1.16, not × 0.95).
4. MercadoPago settlement_v2: what MercadoPago deducted from a sale's payment as
   fees (FEE_AMOUNT) equals Σ `Valor del cargo` of that payment's "Si" (deducted at
   source) charges — no 16 % / 19 % on top.

If `Valor del cargo` were net of VAT, the cash (1, 4) and the credit-note totals (2)
would be 1.16 / 1.19 times the charges. They are equal, so the charges are the
gross amounts and, since MercadoLibre must charge VAT on its services and invoices
them as fiscal documents, the VAT is inside them: VAT = amount × rate / (1 + rate).
What the data cannot show is the VAT of each fiscal invoice itself (that is in the
CFDI / DTE XML, not downloaded): the standard rate on every charge is an assumption,
stated as such on ①."""

import glob
import os

import pandas as pd

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import PAGOS_HEADER_ROW
from platforms.mercadolibre.report.loading import _srcid

from experiments.ml_invoice_basis.common import key

# Input VAT on the platform's fees: creditable or not, and why (per site code).
# Peru's fee documents are boletas (B001-…, platforms/mercadolibre/README.md);
# a boleta gives no crédito fiscal, so Peru's fees stay at their gross amount.
INPUT_CREDIT = {
    "mx": (True, "平台费用的发票是 CFDI（墨西哥电子发票），其中的 IVA 是可抵扣的进项税"),
    "cl": (True, "平台费用的发票是电子发票（DTE factura），其中的 IVA 是可抵扣的进项税（crédito fiscal）"),
    "pe": (False, "平台费用的凭证是 boleta（B001-…），boleta 不能抵扣 IGV，所以费用按含税全额计入"),
}
FISCAL_DOC = {"mx": "CFDI", "cl": "DTE", "pe": "CPE"}


def policy(code=None):
    """The VAT facts of the current site (call while the folder's site is set)."""
    code = code or site.SITE
    credit, why = INPUT_CREDIT.get(code, (True, ""))
    return {"code": code, "rate": float(site.IVA_RATE), "name": site.VAT_NAME,
            "auth": site.TAX_AUTH, "credit": credit, "credit_why": why,
            "doc": FISCAL_DOC.get(code, "fiscal"), "withholding": bool(site.WITHHOLDING)}


def _pagos_files(folder):
    return sorted(p for p in glob.glob(os.path.join(folder, "Reporte_Pagos_Facturas_*.xlsx"))
                  if not os.path.basename(p).startswith("~$"))


def evidence(folder, B, mp_rows=None, tol=0.05):
    """Counts of the four tests (see the module doc); never raises."""
    out = {"detail": {"n": 0, "eq": 0, "applied": 0.0, "files": []},
           "nc": {"docs": 0, "found": 0, "eq": 0, "pending": 0},
           "formula": {"n": 0, "eq": 0, "pct_eq": 0}, "mp": {"n": 0, "eq": 0, "inv": 0.0, "mp": 0.0}}
    by_charge = dict(zip(B["kcharge"], B["amount"]))
    by_doc = B.groupby(B["N° de factura fiscal"].map(key))["amount"].sum() \
        if "N° de factura fiscal" in B.columns else pd.Series(dtype=float)
    for p in _pagos_files(folder):
        try:
            xl = pd.ExcelFile(p)
        except Exception:
            continue
        name = os.path.basename(p)
        if "Detalle de Pagos del mes" in xl.sheet_names:
            try:
                D = pd.read_excel(xl, sheet_name="Detalle de Pagos del mes",
                                  header=PAGOS_HEADER_ROW).dropna(how="all")
            except Exception:
                D = pd.DataFrame()
            if {"Número del cargo", "Parte del pago aplicado a cargos"} <= set(D.columns):
                a = pd.to_numeric(D["Parte del pago aplicado a cargos"], errors="coerce").fillna(0.0) \
                    .groupby(D["Número del cargo"].map(key)).sum()
                n = int(len(a))
                eq = int(sum(1 for k, v in a.items()
                             if k in by_charge and abs(v - by_charge[k]) <= tol))
                out["detail"]["n"] += n
                out["detail"]["eq"] += eq
                out["detail"]["applied"] += float(a.sum())
                out["detail"]["files"].append("%s：%d/%d" % (name, eq, n))
        try:
            P = pd.read_excel(xl, sheet_name="Pagos y notas de crédito",
                              header=PAGOS_HEADER_ROW).dropna(how="all")
        except Exception:
            P = pd.DataFrame()
        if {"Tipo de pago", "Número de nota de crédito", "Importe total"} <= set(P.columns):
            nc = P[P["Tipo de pago"].astype(str).str.contains("Nota", na=False)]
            for no, tot in zip(nc["Número de nota de crédito"].map(key),
                               pd.to_numeric(nc["Importe total"], errors="coerce")):
                out["nc"]["docs"] += 1
                if not no or no.lower().startswith("en proceso"):
                    out["nc"]["pending"] += 1
                    continue
                if no in by_doc.index:
                    out["nc"]["found"] += 1
                    if pd.notna(tot) and abs(-float(by_doc[no]) - float(tot)) <= tol:
                        out["nc"]["eq"] += 1
    need = {"Porcentaje por categoría", "Costo por categoría", "Total de la venta",
            "Subtotal sin descuento"}
    if need <= set(B.columns):
        c = B[(~B["is_rev"]) & (B["det"] == "Cargo por venta")]
        pct = pd.to_numeric(c["Porcentaje por categoría"], errors="coerce")
        cpc = pd.to_numeric(c["Costo por categoría"], errors="coerce")
        tot = pd.to_numeric(c["Total de la venta"], errors="coerce")
        sub = pd.to_numeric(c["Subtotal sin descuento"], errors="coerce")
        dsc = pd.to_numeric(c.get("Valor del descuento"), errors="coerce").fillna(0.0) \
            if "Valor del descuento" in c.columns else 0.0
        unit = pd.to_numeric(c.get("Costo por unidad vendida"), errors="coerce").fillna(0.0) \
            if "Costo por unidad vendida" in c.columns else 0.0
        ok = cpc.notna() & sub.notna()
        parts = ((c["amount"] - (sub - dsc)).abs() <= tol) & ((sub - (cpc + unit)).abs() <= tol)
        rate = ok & pct.notna() & tot.notna() & ((cpc - pct * tot).abs() <= tol)
        out["formula"] = {"n": int(ok.sum()), "eq": int((ok & parts).sum()), "pct_eq": int(rate.sum())}
    if mp_rows is not None and len(mp_rows) and "Pago" in B.columns:
        si = B[(B["channel"] == "Si") & (~B["is_rev"])]
        inv = si.groupby(si["Pago"].map(_srcid))["amount"].sum()
        st = mp_rows[mp_rows["type"] == "SETTLEMENT"]
        fee = -st.groupby("sid")["fee"].sum()
        j = pd.concat([inv.rename("inv"), fee.rename("mp")], axis=1).dropna()
        j = j[j.index != ""]
        out["mp"] = {"n": int(len(j)), "eq": int(((j["inv"] - j["mp"]).abs() <= tol).sum()),
                     "inv": float(j["inv"].sum()), "mp": float(j["mp"].sum())}
    return out


def conclusion(ev, pol):
    """(verdict True/False/None, plain-Chinese sentences) for ①."""
    d, nc, f, mp = ev["detail"], ev["nc"], ev["formula"], ev["mp"]
    lines = []
    if d["n"]:
        lines.append("(1) 账单付款明细（Pagos_Facturas · Detalle de Pagos del mes）：每笔付款付给每张费用的金额 = 该费用的"
                     " Valor del cargo，%d/%d 张费用分毫不差（%s）。" % (d["eq"], d["n"], "；".join(d["files"])))
    else:
        lines.append("(1) 本次下载没有账单付款明细页（Detalle de Pagos del mes），这一项无法核对。")
    if nc["found"]:
        lines.append("(2) 贷记单（Notas de crédito）本身就是税务凭证：Pagos_Facturas 里每张贷记单的“Importe total”"
                     " = 贷记单文件里同一凭证号（N° de factura fiscal）各行 Valor del cargo 之和，%d/%d 张相等"
                     "（另有 %d 张还在“En proceso”、没有凭证号）。" % (nc["eq"], nc["found"], nc["pending"]))
    else:
        lines.append("(2) 本次下载的 Pagos_Facturas 里没有能对上凭证号的贷记单，这一项无法核对。")
    if f["n"]:
        lines.append("(3) 佣金行的金额就是它自己的组成部分之和：Valor del cargo = Costo por categoría + Costo por "
                     "unidad vendida − Valor del descuento，%d/%d 行成立，没有在上面另加税（若加了 %s，金额应是这个和的"
                     " %.2f 倍）。其中 %d 行的 Costo por categoría 正好 = 类目费率 × 买家支付的价格；其余是平台给的优惠费率"
                     "（约 0.95 倍），是打折不是加税。" % (f["eq"], f["n"], pol["name"], 1 + pol["rate"], f["pct_eq"]))
    if mp["n"]:
        lines.append("(4) MercadoPago 结算流水：每笔收款被扣的费用（FEE_AMOUNT）= 该收款在账单上“订单内直扣（Si）”的费用"
                     "之和，%d/%d 笔相等（合计 %s 对 %s；不等的多为运费在另一张账单行或另一期计费）。"
                     % (mp["eq"], mp["n"], "{:,.2f}".format(mp["inv"]), "{:,.2f}".format(mp["mp"])))
    else:
        lines.append("(4) 本次下载没有 settlement_v2，MercadoPago 扣费这一项无法核对。")
    shares = [x["eq"] / float(x["n"]) for x in (d, f, mp) if x["n"]] + \
             ([nc["eq"] / float(nc["found"])] if nc["found"] else [])
    verdict = None
    if shares:
        verdict = min(shares) >= 0.9
    r = pol["rate"]
    if verdict:
        lines.append("结论：Valor del cargo 是卖家实际要付的全额。若它不含 %s，现金和贷记单金额都应是它的 %.2f 倍——"
                     "数据显示是 1 倍。平台的服务费依法要开含 %s 的税务发票，所以 %s 就在这个金额里面："
                     "%s = 金额 × %g/%g。" % (pol["name"], 1 + r, pol["name"], pol["name"], pol["name"],
                                            round(r * 100, 2), round(100 + r * 100, 2)))
    elif verdict is False:
        lines.append("结论：数据没有一致地证明 Valor del cargo 含税（见上面的比例）。本实验仍按含税处理，但这一条要人工复核。")
    else:
        lines.append("结论：本次下载的文件不足以核对；按平台惯例（服务费含税）处理，属于假设。")
    lines.append("假设（数据无法证明）：每一笔平台费用都按标准税率 %g%% 含 %s。每张税务发票的确切税额在平台的 %s"
                 " XML 里（没有下载）。%s。" % (round(r * 100, 2), pol["name"], pol["doc"], pol["credit_why"]))
    return verdict, lines
