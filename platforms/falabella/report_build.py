# -*- coding: utf-8 -*-
"""
Falabella monthly financial workbook — the sister of MercadoLibre's.

Same ten sheets, same look, same check levels (core/report_kit.py,
docs/REPORT_FAMILY.md); the content is Falabella's. Accounting rules are in
platforms/falabella/README.md ("Accounting rules"); the short version:

* **Basis: transaction date.** Every money line of the month comes from
  "Transacciones" filtered on Fecha de transacción in the month. Falabella books a
  sale when it is delivered and pays it out from there; fees, refunds and
  promotions carry their own dates. Order-date GMV (from the order list) is a KPI
  and is reconciled to the transaction basis in ③ bridge A.
* **Net of IVA.** The P&L uses "Monto (Sin IVA)". The seller is an IVA taxpayer:
  output IVA on sales and input IVA on Falabella's fee invoices are not income or
  cost. Both are shown as memo lines; the cash bridge (③ C) uses amounts with IVA.
* **Every transaction type is registered** (fee_catalog.py). An unregistered one is
  an error check and its amount goes to a visible "未登记科目 ⚠" row.
* **Product cost** is not in Falabella's data: a yellow input row in ② and a unit
  cost column in ⑦.

Inputs: one download folder per store (platforms/falabella/run_downloads.py):
Ordenes_*.xlsx, NewReportTransaction_*.xlsx, InvoiceReport_*.xlsx,
*_FSC_LIQ_*.xlsx, Product_details_*.xlsx, run_meta.json. Customer columns of the
order list are dropped on load and never reach the workbook.
"""
import datetime
import glob
import io
import json
import os
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from core import report_kit as K
from core.report_kit import F, fill, wrap, BOX
from platforms.falabella import fee_catalog as FC
from platforms.falabella import periods

SHEETS = K.SHEET_NAMES
CUR = "CLP"
MNY = K.money_format(0)
IVA_RATE = 0.19
TOL = 2.0                       # CLP; rounding tolerance per bridge

EXIT_OK, EXIT_CHECKS, EXIT_NO_STORE, EXIT_CRASH = 0, 1, 2, 70

# order-list columns we keep (everything else — names, RUT, addresses, phones,
# e-mails — is dropped on load)
ORDER_COLS = ["Item id", "SKU del vendedor", "SKU de Falabella", "Fecha de creacion",
              "N° orden", "Orden id", "Origen del pedido", "Moneda", "Documento requerido",
              "Metodo de pago", "Precio pagado", "Precio unitario", "Costo de envio",
              "Cupon/Descuentos", "Impuesto", "Producto", "Variacion", "Operador Logistico",
              "Tipo de envio", "Tipo de proveedor de envio", "Fecha limite de despacho",
              "Estado", "Tipo de Fulfillment", "Motivo de cancelacion"]
T_COLS = {"date": "Fecha de transacción", "type": "Tipo de transacción", "order": "Nº de orden",
          "item": "Id Artículo", "sku": "SKU vendedor", "name": "Nombre del producto",
          "net": "Monto (Sin IVA)", "iva": "IVA", "gross": "Monto con IVA",
          "pay": "Estado de pago", "stmt": "Nº estado de cuenta", "doc": "Tipo de Documento Tributario",
          "mode": "Modalidad logística", "status": "Estado de la orden",
          "created": "Fecha creación de la orden", "cat": "Categoría de transacciones",
          "pct": "% comisión"}

# ② fixed rows: (key, 中文, 平台原文, kind, 说明)   kind: d 取数 / sum 小计 / f 计算 / in 录入
PL_ROWS = [
    (FC.SALES, "商品销售收入", "Pago por precio del producto", "d",
     "本月结算的商品价款（不含 IVA），按交易日（≈ 送达日）归集；含 CMR 积分支付的部分"),
    (FC.SHIP_INCOME, "运费收入", "Pago de envío comprador / Pago por envío directo", "d",
     "买家支付的运费，以及平台付给卖家自配送订单的运费"),
    (None, "营业收入合计", "", "sum", "＝ 商品销售收入 ＋ 运费收入"),
    (FC.REFUND, "减：退货退款", "Descuento por devolución de producto", "d",
     "退货冲回的商品价款（含申诉裁定）"),
    (FC.PROMO, "加：平台出资促销与激励", "Aportes promocionales / Oportunidad Única", "d",
     "Falabella 出资的优惠券、CMR 卡活动补贴、销售目标奖励"),
    (FC.OTHER_INCOME, "加：丢失损坏赔付", "Compensación por extravío o producto dañado", "d",
     "物流丢失/损坏的赔付净额"),
    (None, "净销售收入", "", "f", "＝ 营业收入合计 − 退货退款 ＋ 平台促销 ＋ 赔付"),
    (FC.COMMISSION, "减：销售佣金", "Cobro por comisión por venta", "d",
     "按品类费率；退货订单的佣金退还已抵减"),
    (FC.SHIP_COFIN, "减：配送共担费", "Cobro por cofinanciamiento logístico", "d",
     "卖家承担的配送费（FBF 与自配送订单）"),
    (FC.SHIP_BUYER, "减：买家运费转付平台", "Reversa de pago de envío comprador", "d",
     "买家付的运费由平台配送，原样转付平台；与运费收入里买家付的部分相抵"),
    (FC.SHIP_PROMO, "减：免运费促销", "Promo envío falabella.com / Promo Envio Seller", "d",
     "平台免运促销先扣后还（净额通常为 0）；卖家出资的免运促销"),
    (FC.SHIP_REVERSE, "减：退货逆向物流", "Cobro por logística inversa", "d", "退货的取回运费"),
    (FC.FULFILMENT, "减：Fulfillment 仓储与操作", "Fulfillment by Falabella", "d",
     "FBF 仓储、超期仓储、拣货打包、揽收、取回库存等"),
    (FC.ADS, "减：广告推广", "Productos patrocinados / Banners …", "d", "店铺级费用，无法分摊到订单"),
    (FC.SUBSCRIPTION, "减：订阅与增值服务", "Plan marketplace 360 / Mi tienda …", "d", ""),
    (FC.PENALTY, "减：罚款与违规扣费", "Comisión por cancelación / Penalty …", "d",
     "卖家原因取消的佣金罚款、尺寸申报不实等"),
    (FC.OTHER, "减：其他调整", "Corrección de pagos / cobros", "d", ""),
    ("unregistered", "减：未登记科目 ⚠", "(fee_catalog 里没有)", "d",
     "字典里没有的交易类型。金额照常计入以保证勾稽，但需人工确认科目归属，见 ⑩ 页"),
    (None, "平台费用小计", "", "sum", "从结算款中直接扣除"),
    (None, "经营贡献毛利（未扣商品成本）", "", "f", "＝ 净销售收入 − 平台费用小计"),
    ("cogs", "减：商品采购成本", "", "in",
     "⚠ 平台数据不含此项：在【⑦ SKU盈利分析】填单位成本，或直接在此填总额"),
    (None, "税前净利润", "", "f", "＝ 经营贡献毛利 − 商品采购成本"),
]
INCOME_LINES = (FC.SALES, FC.SHIP_INCOME, FC.REFUND, FC.PROMO, FC.OTHER_INCOME)
FEE_LINES = (FC.COMMISSION, FC.SHIP_COFIN, FC.SHIP_BUYER, FC.SHIP_PROMO, FC.SHIP_REVERSE,
             FC.FULFILMENT, FC.ADS, FC.SUBSCRIPTION, FC.PENALTY, FC.OTHER, "unregistered")

ACCOUNTANT_QUESTIONS = [
    "销项 IVA：Fulfillment by Falabella 订单的 boleta/factura 由 Falabella 以卖家名义开具（见【③】桥 D "
    "与 Liquidación）；卖家自配送（FBS）订单的销售单据由谁开、是否已开，请会计师确认，避免漏报或重复报。",
    "进项 IVA：Falabella 每月（21 日–次月 20 日）开具的费用发票（Factura）与贷记单（Nota de Crédito）"
    "所含 IVA 能否全部作为进项抵扣。",
    "平台出资的促销补贴、CMR 卡“Oportunidad Única”补贴：本报表计入净销售收入（含 IVA 行另列），"
    "开票与入账方式请会计师确认。",
    "物流丢失/损坏赔付：本报表计入“丢失损坏赔付”（收入侧），是否需要开具单据请确认。",
]


# ======================================================================= helpers

def fmt(x):
    try:
        return "{:,.0f}".format(float(x))
    except (TypeError, ValueError):
        return "-"


def pct(x, dec=1):
    try:
        return "{:.{d}f}%".format(float(x) * 100, d=dec)
    except (TypeError, ValueError):
        return "-"


def money_col(s):
    """'13,990.00' / 13990 / NaN -> float."""
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False)
                         .str.replace("$", "", regex=False).str.strip()
                         .replace({"": None, "nan": None, "None": None}),
                         errors="coerce").fillna(0.0)


def _first(folder, pattern):
    hits = sorted(glob.glob(os.path.join(folder, pattern)))
    return hits[-1] if hits else None


def _all(folder, pattern):
    return sorted(glob.glob(os.path.join(folder, pattern)))


def read_meta(folder):
    try:
        with io.open(os.path.join(folder, "run_meta.json"), encoding="utf-8") as fh:
            return json.load(fh) or {}
    except Exception:
        return {}


# ======================================================================= load

def load_transactions(path):
    df = pd.read_excel(path, sheet_name="Reporte Transacciones", header=5, dtype=str)
    df = df[df[T_COLS["type"]].notna()].copy()
    for k in ("net", "iva", "gross"):
        df[T_COLS[k]] = money_col(df[T_COLS[k]])
    df["_date"] = pd.to_datetime(df[T_COLS["date"]].str[:10], errors="coerce")
    banner = pd.read_excel(path, sheet_name="Reporte Transacciones", header=None, nrows=1)
    updated = banner.iloc[0, 1] if banner.shape[1] > 1 else None
    return df, updated


def load_orders(paths):
    frames = []
    for p in paths:
        o = pd.read_excel(p, dtype=str)
        frames.append(o[[c for c in ORDER_COLS if c in o.columns]].copy())
    if not frames:
        return None
    o = pd.concat(frames, ignore_index=True).drop_duplicates(subset=["Item id"])
    for c in ("Precio pagado", "Precio unitario", "Costo de envio", "Cupon/Descuentos", "Impuesto"):
        if c in o.columns:
            o[c] = money_col(o[c])
    o["_created"] = pd.to_datetime(o["Fecha de creacion"], format="%b %d, %Y %H:%M",
                                   errors="coerce")
    return o


def load_invoice(paths):
    frames = []
    for p in paths:
        x = pd.read_excel(p, dtype=str)
        for c in ("Monto (Sin IVA)", "IVA", "Monto con IVA"):
            x[c] = money_col(x[c])
        x["_date"] = pd.to_datetime(x["Fecha de Transaccion"].str[:10], errors="coerce")
        m = re.search(r"(\d{4}-\d{2}-\d{2})-(\d{4}-\d{2}-\d{2})_", os.path.basename(p))
        x["_period"] = "%s_%s" % m.groups() if m else os.path.basename(p)
        frames.append(x)
    return pd.concat(frames, ignore_index=True) if frames else None


def load_liquidation(path):
    if not path:
        return None
    docs = pd.read_excel(path, sheet_name="Docs. liquidados", dtype=str)
    docs["Monto línea"] = money_col(docs["Monto línea"])
    return docs


def load_stock(path):
    if not path:
        return None
    s = pd.read_excel(path, dtype=str)
    for c in ("Precio", "Ventas ultimas 4 semanas", "Stock Disponible"):
        if c in s.columns:
            s[c] = money_col(s[c])
    return s


def load_store(store, folder, month):
    a, b = periods.month_bounds(month)
    S = {"store": store, "folder": folder, "month": month, "meta": read_meta(folder),
         "files": {}, "problems": []}
    tp = _first(folder, "NewReportTransaction_*.xlsx")
    S["files"]["billing"] = tp
    if tp:
        t, upd = load_transactions(tp)
        S["tx_all"] = t
        S["tx_updated"] = upd
        S["tx"] = t[(t["_date"] >= pd.Timestamp(a)) & (t["_date"] <= pd.Timestamp(b))].copy()
        S["tx_outside"] = int(len(t) - len(S["tx"]))
    else:
        S["tx"] = None
    ops = _all(folder, "Ordenes_*.xlsx")
    S["files"]["order"] = ops
    S["orders_all"] = load_orders(ops)
    if S["orders_all"] is not None:
        o = S["orders_all"]
        S["orders"] = o[(o["_created"] >= pd.Timestamp(a)) &
                        (o["_created"] < pd.Timestamp(b) + pd.Timedelta(days=1))].copy()
    else:
        S["orders"] = None
    ips = _all(folder, "InvoiceReport_*.xlsx")
    S["files"]["billing1"] = ips
    S["invoice"] = load_invoice(ips)
    lp = _first(folder, "*FSC_LIQ*.xlsx")
    S["files"]["billing2"] = lp
    S["liq"] = load_liquidation(lp)
    sp = _first(folder, "Product_details_*.xlsx")
    S["files"]["stock"] = sp
    S["stock"] = load_stock(sp)
    S["files"]["pdf"] = [os.path.basename(p) for p in _all(folder, "*.pdf")]
    return S


# ======================================================================= derive

def derive(S):
    a, b = periods.month_bounds(S["month"])
    m = {}
    t = S["tx"]
    if t is None or not len(t):
        t = pd.DataFrame(columns=list(T_COLS.values()) + ["_date"])
    looked = t[T_COLS["type"]].map(lambda n: FC.lookup(n) or ("⚠ 未登记：" + str(n), "unregistered"))
    t = t.assign(_zh=looked.map(lambda x: x[0]), _line=looked.map(lambda x: x[1]))
    S["tx"] = t
    m["unregistered_types"] = sorted(set(t.loc[t["_line"] == "unregistered", T_COLS["type"]]))
    by = t.groupby("_line")
    m["pl"] = {k: float(by[T_COLS["net"]].sum().get(k, 0.0)) for k in INCOME_LINES + FEE_LINES}
    m["pl_iva"] = {k: float(by[T_COLS["iva"]].sum().get(k, 0.0)) for k in INCOME_LINES + FEE_LINES}
    m["pl_gross"] = {k: float(by[T_COLS["gross"]].sum().get(k, 0.0)) for k in INCOME_LINES + FEE_LINES}
    m["income"] = sum(m["pl"][k] for k in INCOME_LINES)
    m["fees"] = sum(m["pl"][k] for k in FEE_LINES)
    m["revenue"] = m["pl"][FC.SALES] + m["pl"][FC.SHIP_INCOME]
    m["contrib"] = m["income"] + m["fees"]
    m["tx_net"] = float(t[T_COLS["net"]].sum())
    m["tx_iva"] = float(t[T_COLS["iva"]].sum())
    m["tx_gross"] = float(t[T_COLS["gross"]].sum())
    m["iva_out"] = sum(m["pl_iva"][k] for k in INCOME_LINES)
    m["iva_in"] = -sum(m["pl_iva"][k] for k in FEE_LINES)
    m["lines"] = int(len(t))
    sales = t[t["_line"] == FC.SALES]
    m["sales_items"] = int(len(sales))
    m["sales_orders"] = int(sales[T_COLS["order"]].nunique())
    m["sales_gross"] = float(sales[T_COLS["gross"]].sum())
    m["by_mode"] = sales.groupby(T_COLS["mode"])[T_COLS["gross"]].sum().to_dict() \
        if len(sales) else {}
    pay = t.groupby(T_COLS["pay"])[T_COLS["gross"]].agg(["sum", "count"]) if len(t) else None
    m["paid"] = {str(k): (float(v["sum"]), int(v["count"])) for k, v in pay.iterrows()} \
        if pay is not None else {}
    m["statements"] = int(t[T_COLS["stmt"]].nunique()) if len(t) else 0
    m["iva_bad_lines"] = int(((t[T_COLS["iva"]] - t[T_COLS["net"]] * IVA_RATE).abs() > 1.5).sum()) \
        if len(t) else 0
    m["gross_bad_lines"] = int(((t[T_COLS["gross"]] - t[T_COLS["net"]] - t[T_COLS["iva"]]).abs()
                                > 0.5).sum()) if len(t) else 0
    # ---- order-date KPIs
    o = S["orders"]
    if o is not None:
        live = o[o["Estado"] != "Canceladas"]
        m["gmv_orders"] = float(live["Precio pagado"].sum())
        m["orders_n"] = int(o["N° orden"].nunique())
        m["orders_live_n"] = int(live["N° orden"].nunique())
        m["items_n"] = int(len(live))
        m["cancel_n"] = int((o["Estado"] == "Canceladas").sum())
        m["cancel_amt"] = float(o.loc[o["Estado"] == "Canceladas", "Precio pagado"].sum())
        m["return_n"] = int((o["Estado"] == "Devoluciones").sum())
        m["coupon"] = float(o["Cupon/Descuentos"].sum())
    # ---- bridge A: orders of the month <-> sales settled in the month (item level)
    if o is not None:
        live = o[o["Estado"] != "Canceladas"]
        t_items = set(sales[T_COLS["item"]].astype(str))
        o_items_all = set(S["orders_all"]["Item id"].astype(str))
        both = live[live["Item id"].isin(t_items)]
        sx = sales.set_index(sales[T_COLS["item"]].astype(str))[T_COLS["gross"]]
        A = {"orders": float(live["Precio pagado"].sum()),
             "matched_orders": float(both["Precio pagado"].sum()),
             "matched_tx": float(sx.reindex(both["Item id"]).sum()),
             "not_settled": float(live.loc[~live["Item id"].isin(t_items), "Precio pagado"].sum()),
             "not_settled_n": int((~live["Item id"].isin(t_items)).sum()),
             "earlier": float(sales.loc[~sales[T_COLS["item"]].astype(str).isin(o_items_all),
                                        T_COLS["gross"]].sum()),
             "earlier_n": int((~sales[T_COLS["item"]].astype(str).isin(o_items_all)).sum()),
             "sales": m["sales_gross"]}
        A["price_diff"] = A["matched_tx"] - A["matched_orders"]
        A["resid"] = A["sales"] - (A["orders"] - A["not_settled"] + A["earlier"] + A["price_diff"])
        A["not_settled_states"] = live.loc[~live["Item id"].isin(t_items), "Estado"] \
            .value_counts().to_dict()
        A["earlier_months"] = sales.loc[~sales[T_COLS["item"]].astype(str).isin(o_items_all),
                                        T_COLS["created"]].astype(str).str[:7] \
            .value_counts().to_dict()
        m["bridge_a"] = A
    # ---- bridge B: fee invoices <-> invoiced transaction lines, on the days covered
    inv = S["invoice"]
    covered, pending = [], []
    for pa, pb in periods.billing_periods(S["month"]):
        pid = periods.period_id(pa, pb)
        lo, hi = max(pa, a), min(pb, b)
        if inv is not None and pid in set(inv["_period"]):
            covered.append((pid, lo, hi))
        else:
            pending.append((pid, lo, hi))
    Bs = []
    docs = t[t[T_COLS["doc"]].notna()]
    for pid, lo, hi in covered:
        iw = inv[(inv["_period"] == pid) & (inv["_date"] >= pd.Timestamp(lo)) &
                 (inv["_date"] <= pd.Timestamp(hi))]
        tw = docs[(docs["_date"] >= pd.Timestamp(lo)) & (docs["_date"] <= pd.Timestamp(hi))]
        Bs.append({"period": pid, "from": lo, "to": hi, "inv_lines": int(len(iw)),
                   "inv_net": float(iw["Monto (Sin IVA)"].sum()),
                   "inv_gross": float(iw["Monto con IVA"].sum()),
                   "tx_lines": int(len(tw)), "tx_net": float(tw[T_COLS["net"]].sum()),
                   "tx_gross": float(tw[T_COLS["gross"]].sum()),
                   "factura": float(inv.loc[(inv["_period"] == pid) &
                                            (inv["Tipo de documento"] == "Factura"),
                                            "Monto con IVA"].sum()),
                   "nc": float(inv.loc[(inv["_period"] == pid) &
                                       (inv["Tipo de documento"] == "Nota de Credito"),
                                       "Monto con IVA"].sum())})
    # Lines get their invoice number only when the period's invoice is issued, so
    # in a period still open nothing carries one yet: count the kinds of lines
    # that do get invoiced (learned from this month's invoiced lines).
    invoiced_kinds = set(docs["_line"]) or set(FEE_LINES)
    pend = []
    for pid, lo, hi in pending:
        tw = t[t["_line"].isin(invoiced_kinds) & (t["_date"] >= pd.Timestamp(lo)) &
               (t["_date"] <= pd.Timestamp(hi))]
        pend.append({"period": pid, "from": lo, "to": hi, "tx_lines": int(len(tw)),
                     "tx_net": float(tw[T_COLS["net"]].sum()),
                     "tx_gross": float(tw[T_COLS["gross"]].sum())})
    m["bridge_b"], m["bridge_b_pending"] = Bs, pend
    open_days = pd.Series(False, index=t.index)
    for _pid, lo, hi in pending:
        open_days |= (t["_date"] >= pd.Timestamp(lo)) & (t["_date"] <= pd.Timestamp(hi))
    m["not_invoiced_net"] = float(t.loc[t[T_COLS["doc"]].isna() & t["_line"].isin(FEE_LINES)
                                        & ~open_days, T_COLS["net"]].sum())
    # ---- bridge D: liquidation documents <-> FBF sales
    L = S["liq"]
    if L is not None:
        tot = L.groupby("Tipo de documento")["Monto línea"].sum().to_dict()
        m["liq"] = {"boleta": float(tot.get("BOLETA", 0.0)), "factura": float(tot.get("FACTURA", 0.0)),
                    "nc": float(tot.get("NOTA_DE_CREDITO", 0.0)),
                    "docs": int(L["Folio"].nunique()), "orders": int(L["N° de orden"].nunique())}
        m["liq"]["total"] = m["liq"]["boleta"] + m["liq"]["factura"] + m["liq"]["nc"]
    m["fbf_sales"] = float(m["by_mode"].get("FBF", 0.0))
    m["fbs_sales"] = float(m["by_mode"].get("FBS", 0.0))
    S["m"] = m
    S["sku"] = sku_table(S)
    S["returns"] = returns_table(S)
    return S


def sku_table(S):
    t = S["tx"]
    if t is None or not len(t):
        return pd.DataFrame()
    sk = t[T_COLS["sku"]].fillna("(无 SKU · 店铺级费用)")
    g = t.assign(_sku=sk).pivot_table(index="_sku", columns="_line", values=T_COLS["net"],
                                     aggfunc="sum", fill_value=0.0)
    for k in INCOME_LINES + FEE_LINES:
        if k not in g.columns:
            g[k] = 0.0
    sales = t[t["_line"] == FC.SALES].assign(_sku=sk)
    g["units"] = sales.groupby("_sku").size().reindex(g.index).fillna(0).astype(int)
    g["returns"] = t[t["_line"] == FC.REFUND].assign(_sku=sk).groupby("_sku").size() \
        .reindex(g.index).fillna(0).astype(int)
    names = t.assign(_sku=sk).groupby("_sku")[T_COLS["name"]].first()
    g["name"] = names.reindex(g.index)
    g["income"] = g[list(INCOME_LINES)].sum(axis=1)
    g["fees"] = g[list(FEE_LINES)].sum(axis=1)
    g["contrib"] = g["income"] + g["fees"]
    return g.sort_values(FC.SALES, ascending=False)


def returns_table(S):
    o = S["orders"]
    out = {}
    if o is not None:
        bad = o[o["Estado"].isin(["Canceladas", "Devoluciones"])]
        out["by_status"] = bad.groupby("Estado")["Precio pagado"].agg(["count", "sum"])
        out["by_reason"] = bad.assign(r=bad["Motivo de cancelacion"].fillna("(未写原因)")) \
            .groupby(["Estado", "r"])["Precio pagado"].agg(["count", "sum"]) \
            .sort_values("count", ascending=False)
        out["by_fulfilment"] = bad.groupby(["Estado", "Tipo de Fulfillment"]).size()
    t = S["tx"]
    ret_orders = set(t.loc[t["_line"] == FC.REFUND, T_COLS["order"]])
    r = t[t[T_COLS["order"]].isin(ret_orders)]
    out["money"] = r.groupby(["_line", "_zh"])[T_COLS["net"]].agg(["count", "sum"]) \
        if len(r) else pd.DataFrame()
    out["orders"] = len(ret_orders)
    return out


# ======================================================================= checks

def validate(stores):
    C = []
    for S in stores:
        st, m = S["store"], S["m"]
        f = S["files"]
        C.append(K.check("file_billing", st, "交易明细（Transacciones）在", bool(f.get("billing")),
                         "损益表的全部金额来自它" if f.get("billing") else
                         "没有 NewReportTransaction 文件：本店本月的损益表是空的，必须重新下载"))
        C.append(K.check("file_order", st, "订单列表在", bool(f.get("order")),
                         "%d 个文件" % len(f.get("order") or []) if f.get("order") else
                         "没有 Ordenes 文件：桥 A、退货分析和订单口径 GMV 无法计算",
                         severity="error"))
        C.append(K.check("file_stock", st, "库存快照在", bool(f.get("stock")),
                         os.path.basename(f["stock"]) if f.get("stock") else
                         "没有 Product_details：⑨ 库存页为空", severity="warn"))
        C.append(K.check("file_billing2", st, "Liquidación factura 在", bool(f.get("billing2")),
                         os.path.basename(f["billing2"]) if f.get("billing2") else
                         "没有 Liquidación：桥 D 无法核对平台代开的销售单据", severity="warn"))
        C.append(K.check("tx_in_month", st, "交易明细只取本月",
                         m["lines"] > 0,
                         "本月 %d 行；文件里另有 %d 行不在本月，已排除"
                         % (m["lines"], S.get("tx_outside") or 0) if m["lines"] else
                         "交易明细里没有任何一行在本月"))
        C.append(K.check("unregistered", st, "交易类型全部已登记", not m["unregistered_types"],
                         "全部 %d 种已登记" % S["tx"]["_line"].nunique() if not m["unregistered_types"]
                         else "未登记：%s（金额已计入“未登记科目”行）" % "；".join(m["unregistered_types"])))
        C.append(K.check("pl_tie", st, "损益表各行合计 ＝ 交易明细合计（不含 IVA）",
                         abs(m["contrib"] - m["tx_net"]) <= TOL,
                         "损益 %s，明细 %s，差 %s" % (fmt(m["contrib"]), fmt(m["tx_net"]),
                                                  fmt(m["contrib"] - m["tx_net"]))))
        C.append(K.check("iva", st, "每行 IVA ＝ 19% × 不含税金额", m["iva_bad_lines"] == 0
                         and m["gross_bad_lines"] == 0,
                         "全部 %d 行相符（容差 1.5）" % m["lines"] if not m["iva_bad_lines"]
                         else "%d 行 IVA 不是 19%%、%d 行含税≠不含税＋IVA"
                         % (m["iva_bad_lines"], m["gross_bad_lines"]), severity="warn"))
        A = m.get("bridge_a")
        if A:
            C.append(K.check("bridge_a", st, "桥 A：订单 → 本月结算商品价款", abs(A["resid"]) <= TOL,
                             "残差 %s；价格差 %s" % (fmt(A["resid"]), fmt(A["price_diff"])),
                             severity="warn"))
        else:
            C.append(K.check("bridge_a", st, "桥 A：订单 → 本月结算商品价款", True,
                             "没有订单列表", severity="warn", skipped=True))
        for B in m["bridge_b"]:
            d = B["inv_net"] - B["tx_net"]
            C.append(K.check("bridge_b", st, "桥 B：费用发票 %s ↔ 交易明细（%s–%s）"
                             % (B["period"], B["from"], B["to"]), abs(d) <= TOL,
                             "发票 %d 行 %s，明细 %d 行 %s，差 %s"
                             % (B["inv_lines"], fmt(B["inv_net"]), B["tx_lines"], fmt(B["tx_net"]),
                                fmt(d))))
        for P in m["bridge_b_pending"]:
            C.append(K.check("bridge_b_open", st, "桥 B：账期 %s 的费用发票还没出" % P["period"],
                             False, "%s–%s 的 %d 行应开票的费用 %s（不含 IVA）暂无发票可核对；账期结束后（次月 21 日起）"
                             "重新下载并重出本报表" % (P["from"], P["to"], P["tx_lines"],
                                                   fmt(P["tx_net"])), severity="warn"))
        unpaid = m["paid"].get("No Pagado")
        C.append(K.check("unpaid", st, "本月交易已全部付款", not unpaid,
                         "全部“Pagado”" if not unpaid else
                         "%d 行“No Pagado”合计 %s（含 IVA），下个结算日才到账"
                         % (unpaid[1], fmt(unpaid[0])), severity="warn"))
        if m.get("liq"):
            C.append(K.check("bridge_d", st, "桥 D：Liquidación 单据已下载", True,
                             "平台代开 %d 张单据，合计 %s（含 IVA）" % (m["liq"]["docs"],
                                                               fmt(m["liq"]["total"])),
                             severity="warn"))
    return C


def store_meta_missing(S):
    return [r for r in (S["meta"].get("missing") or []) if r.get("form")]


# ======================================================================= workbook

def build_workbook(stores, ctx, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    sheet_readme(wb, stores, ctx)
    sheet_pl(wb, stores, ctx)
    sheet_bridge(wb, stores, ctx)
    sheet_fees(wb, stores, ctx)
    sheet_returns(wb, stores, ctx)
    sheet_not_applicable(wb, SHEETS[5], "退货质检明细",
                         "不适用：Falabella 不提供逐件退货质检报表。退货的金额与原因见【%s】。" % SHEETS[4])
    sheet_sku(wb, stores, ctx)
    sheet_trend(wb, stores, ctx)
    sheet_inventory(wb, stores, ctx)
    sheet_gaps(wb, stores, ctx)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    return out_path


def _cols(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _section(ws, r, text, ncols):
    c = ws.cell(row=r, column=1, value=text)
    c.font = F(11, True, "FFFFFF")
    for j in range(1, ncols + 1):
        ws.cell(row=r, column=j).fill = fill(K.C_BAND)
    return r + 1


def _money(c, v):
    c.value = v
    c.number_format = MNY
    c.border = BOX
    return c


def _store_header(sh, r, stores, first, extra=("说明",), widths=None):
    cols = list(first) + [S["store"] for S in stores] + (["合计"] if len(stores) > 1 else []) + list(extra)
    sh.header(r, cols, widths or ([34, 30] + [16] * (len(stores) + (len(stores) > 1)) + [60]))
    return r + 1


def _values_row(ws, r, c0, vals, total=True, bold=False, color=None):
    for i, v in enumerate(vals):
        c = _money(ws.cell(row=r, column=c0 + i), v)
        if bold:
            c.font = F(10, True)
        if color:
            c.fill = fill(color)
    if total and len(vals) > 1:
        c = _money(ws.cell(row=r, column=c0 + len(vals)), sum(vals))
        c.font = F(10, True)
        if color:
            c.fill = fill(color)
    return c0 + len(vals) + (1 if total and len(vals) > 1 else 0)


# ----------------------------------------------------------------- ① readme

def sheet_readme(wb, stores, ctx):
    ws = wb.create_sheet(SHEETS[0])
    _cols(ws, [4, 30, 118])
    ws["B1"] = "Falabella 月度财务报表 · %s · %s" % (ctx["period_cn"], ctx["who"])
    ws["B1"].font = F(16, True, K.C_NAVY)
    ws["B2"] = "生成于 %s。所有金额单位：智利比索（CLP），无小数。" % ctx["built"]
    ws["B2"].font = F(9, False, K.C_GREY)
    rows = []
    rows.append(("S", "一、这份报表回答什么", ""))
    rows.append(("", "本月赚了多少", "【%s】从商品销售收入推到“经营贡献毛利（未扣商品成本）”，"
                 "填入商品成本后得到税前净利润。" % SHEETS[1]))
    rows.append(("", "钱对不对得上", "【%s】四座桥：A 订单 → 结算、B 平台费用发票 → 交易明细、"
                 "C 结算净额 → 到账、D 平台代开的销售单据。" % SHEETS[2]))
    rows.append(("", "还缺什么", "【%s】列出未通过的校验、缺的文件和要会计师确认的问题。" % SHEETS[9]))
    rows.append(("", "", ""))
    rows.append(("S", "二、口径（与美客多报表的不同之处）", ""))
    rows.append(("", "按交易日归集", "每一笔收入和费用按 Falabella“Transacciones”报表的交易日"
                 "（Fecha de transacción，商品价款≈送达日）计入当月，和平台的结算一致。"
                 "按下单日的 GMV 作为参考指标列在②下方，两者的差在③桥 A 逐单解释。"))
    rows.append(("", "不含 IVA", "损益表金额不含 IVA（Monto Sin IVA）：卖家是 IVA 纳税人，销项 IVA"
                 "（销售）和进项 IVA（平台费用发票）不是收入也不是成本，在②备查项单列。"
                 "美客多智利报表按含税口径，是因为美客多的文件不拆分 IVA；Falabella 的文件逐行拆分了，"
                 "这里采用会计上正确的不含税口径。现金桥（③桥 C）用含税金额。"))
    rows.append(("", "费用科目", "Falabella 的 147 种交易类型全部登记在字典里（西语原名＋中文＋归类），"
                 "见【%s】。出现字典里没有的类型时报 error，金额进“未登记科目 ⚠”行。" % SHEETS[3]))
    rows.append(("", "", ""))
    rows.append(("S", "三、数据来源", ""))
    for S in stores:
        f = S["files"]
        meta = S["meta"]
        rows.append(("", S["store"], "下载于 %s；卖家 %s %s。交易明细：%s（平台生成于 %s）；订单：%d 个文件；"
                     "费用发票明细：%s；Liquidación：%s；库存：%s；PDF %d 份。"
                     % (meta.get("downloaded_at") or "?", (meta.get("flags") or {}).get("seller") or "",
                        (meta.get("flags") or {}).get("seller_name") or "",
                        os.path.basename(f["billing"]) if f.get("billing") else "缺",
                        S.get("tx_updated") or "?", len(f.get("order") or []),
                        "、".join(os.path.basename(p) for p in f.get("billing1") or []) or "缺",
                        os.path.basename(f["billing2"]) if f.get("billing2") else "缺",
                        os.path.basename(f["stock"]) if f.get("stock") else "缺",
                        len(f.get("pdf") or []))))
    rows.append(("", "", ""))
    rows.append(("S", "四、本期要点", ""))
    tot = lambda k: sum(S["m"].get(k, 0.0) or 0.0 for S in stores)
    rev = sum(S["m"]["pl"][FC.SALES] for S in stores)
    rows.append(("", "规模", "本月结算商品价款 %s（不含 IVA），%d 件；按下单日的 GMV（含 IVA、不含已取消）%s，"
                 "%d 张订单。" % (fmt(rev), int(tot("sales_items")), fmt(tot("gmv_orders")),
                               int(tot("orders_live_n")))))
    if rev:
        rows.append(("", "费用率", "平台费用合计 %s，占商品销售收入 %s；其中佣金 %s、物流 %s、Fulfillment %s、广告 %s。"
                     % (fmt(-tot("fees")), pct(-tot("fees") / rev),
                        pct(-sum(S["m"]["pl"][FC.COMMISSION] for S in stores) / rev),
                        pct(-sum(S["m"]["pl"][k] for S in stores for k in
                                 (FC.SHIP_COFIN, FC.SHIP_BUYER, FC.SHIP_PROMO, FC.SHIP_REVERSE)) / rev),
                        pct(-sum(S["m"]["pl"][FC.FULFILMENT] for S in stores) / rev),
                        pct(-sum(S["m"]["pl"][FC.ADS] for S in stores) / rev))))
        rows.append(("", "贡献毛利", "经营贡献毛利（未扣商品成本）%s，占商品销售收入 %s。"
                     % (fmt(tot("contrib")), pct(tot("contrib") / rev))))
    rows.append(("", "", ""))
    rows.append(("S", "五、各页说明", ""))
    for nm, desc in [
        (SHEETS[1], "各店分列 + 合计。行位置每期固定，可直接跨月对比；灰底=公式，浅蓝=小计，黄底=需你录入。"),
        (SHEETS[2], "桥 A 订单 → 结算、桥 B 费用发票 → 交易明细、桥 C 结算净额 → 到账、桥 D 平台代开销售单据。"),
        (SHEETS[3], "本月出现的每一种交易类型：西语原名、中文、归入哪一行、笔数、不含税、IVA、含税。"),
        (SHEETS[4], "取消与退货：按状态、按原因、按配送方式，以及退货订单上的全部资金流。"),
        (SHEETS[5], "不适用：Falabella 不提供逐件退货质检报表。"),
        (SHEETS[6], "按 SKU 的收入、费用与贡献；黄色列填单位采购成本即得 SKU 毛利。"),
        (SHEETS[7], "本月按周的销售与费用；有往月下载时追加各月对比。"),
        (SHEETS[8], "Fulfillment 库存快照：库存、近 4 周销量、可售周数、滞销与断货风险。"),
        (SHEETS[9], "校验结果、缺失文件及原因、要会计师确认的问题。空表示没有待办。"),
    ]:
        rows.append(("", nm, desc))
    rows.append(("", "", ""))
    rows.append(("S", "六、到账净额 ≠ 利润", ""))
    rows.append(("", "结算净额是现金指标", "交易明细含 IVA 的合计（本月 %s）是 Falabella 本月结算给你的钱"
                 "（含应缴的销项 IVA、已扣的进项 IVA）。利润看②的经营贡献毛利（不含 IVA）。"
                 % fmt(tot("tx_gross"))))
    rows.append(("", "", ""))
    rows.append(("S", "七、要请会计师确认的事", ""))
    for i, q in enumerate(ACCOUNTANT_QUESTIONS, 1):
        rows.append(("", "问题 %d" % i, q))
    r = 4
    for tag, b, c in rows:
        if tag == "S":
            cell = ws.cell(row=r, column=2, value=b)
            cell.font = F(12, True, K.C_NAVY)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            cell.fill = fill(K.C_SUB)
        else:
            ws.cell(row=r, column=2, value=b).font = F(10, True)
            ws.cell(row=r, column=2).alignment = Alignment(vertical="top")
            c2 = ws.cell(row=r, column=3, value=c)
            c2.font = F(10)
            c2.alignment = wrap("top")
            if c and len(c) > 90:
                ws.row_dimensions[r].height = K.est_height(c, 78)
        r += 1
    bad = [c for c in ctx["checks"] if not c["ok"] and not c["skipped"] and c["severity"] == "error"]
    if bad:
        ws.cell(row=r + 1, column=2, value="⚠ 注意").font = F(10, True, K.C_RED)
        c2 = ws.cell(row=r + 1, column=3, value="本期有 %d 项 error 级校验未通过，数字可能不完整，"
                     "请先看【%s】顶部的校验结果。" % (len(bad), SHEETS[9]))
        c2.font = F(10, True, K.C_RED)


# ----------------------------------------------------------------- ② P&L

def sheet_pl(wb, stores, ctx):
    n = len(stores)
    sh = K.Sheet(wb, SHEETS[1], first_store_col=3, nstores=n)
    ws = sh.ws
    sh.title("合并损益表（%s）" % ctx["period_cn"],
             "口径：按【交易日】归集（Falabella 结算口径），金额不含 IVA。单位：CLP。"
             "灰底=公式行，浅蓝=小计，黄底=需你录入。行位置每期固定，可直接跨月对比。")
    r = _store_header(sh, 5, stores, ["项目", "平台科目（原文）"])
    ws.freeze_panes = "C6"
    last = sh.c0 + n - 1 + (1 if n > 1 else 0)
    note_col = last + 1
    rows_at = {}
    first_data = r
    for key, zh, es, kind, note in PL_ROWS:
        ws.cell(row=r, column=1, value=zh).font = F(10, kind in ("sum", "f"))
        ws.cell(row=r, column=2, value=es).font = F(9, False, K.C_GREY)
        if kind == "d":
            vals = [S["m"]["pl"][key] if key != "unregistered" else S["m"]["pl"].get("unregistered", 0.0)
                    for S in stores]
            _values_row(ws, r, sh.c0, vals)
        elif kind == "in":
            for i in range(n):
                c = _money(ws.cell(row=r, column=sh.c0 + i), 0)
                c.fill = fill(K.C_INPUT)
            if n > 1:
                c = ws.cell(row=r, column=sh.c0 + n,
                            value="=SUM(%s%d:%s%d)" % (sh.scol(0), r, sh.scol(n - 1), r))
                c.number_format = MNY
                c.border = BOX
        else:
            for i in range(n + (1 if n > 1 else 0)):
                col = get_column_letter(sh.c0 + i)
                if zh == "营业收入合计":
                    f_ = "=%s%d+%s%d" % (col, rows_at[FC.SALES], col, rows_at[FC.SHIP_INCOME])
                elif zh == "净销售收入":
                    f_ = "=%s%d+%s%d+%s%d+%s%d" % (col, rows_at["营业收入合计"], col, rows_at[FC.REFUND],
                                                 col, rows_at[FC.PROMO], col, rows_at[FC.OTHER_INCOME])
                elif zh == "平台费用小计":
                    f_ = "=SUM(%s%d:%s%d)" % (col, rows_at[FC.COMMISSION], col, rows_at["unregistered"])
                elif zh.startswith("经营贡献毛利"):
                    f_ = "=%s%d+%s%d" % (col, rows_at["净销售收入"], col, rows_at["平台费用小计"])
                else:   # 税前净利润
                    f_ = "=%s%d-%s%d" % (col, rows_at["经营贡献毛利"], col, rows_at["cogs"])
                c = ws.cell(row=r, column=sh.c0 + i, value=f_)
                c.number_format = MNY
                c.border = BOX
                c.font = F(10, True)
                c.fill = fill(K.C_SUB if kind == "sum" else K.C_CALC)
            ws.cell(row=r, column=1).fill = fill(K.C_SUB if kind == "sum" else K.C_CALC)
        ws.cell(row=r, column=note_col, value=note).font = F(9, False, K.C_GREY)
        ws.cell(row=r, column=note_col).alignment = wrap()
        rows_at[key or zh] = r
        if zh.startswith("经营贡献毛利"):
            rows_at["经营贡献毛利"] = r
        r += 1
    # ratios
    r += 1
    r = _section(ws, r, "费用率（占商品销售收入）", note_col)
    for label, keys in [("佣金", (FC.COMMISSION,)),
                        ("物流（配送共担＋运费转付＋免运促销＋逆向物流 − 运费收入）",
                         (FC.SHIP_COFIN, FC.SHIP_BUYER, FC.SHIP_PROMO, FC.SHIP_REVERSE, FC.SHIP_INCOME)),
                        ("Fulfillment 仓储与操作", (FC.FULFILMENT,)), ("广告推广", (FC.ADS,)),
                        ("平台费用合计", FEE_LINES), ("经营贡献毛利率", "contrib")]:
        ws.cell(row=r, column=1, value=label).font = F(10)
        vals = []
        for S in stores:
            base = S["m"]["pl"][FC.SALES]
            v = S["m"]["contrib"] if keys == "contrib" else -sum(S["m"]["pl"][k] for k in keys)
            vals.append(v / base if base else 0.0)
        for i, v in enumerate(vals):
            c = ws.cell(row=r, column=sh.c0 + i, value=v)
            c.number_format = K.PCT
            c.border = BOX
        if n > 1:
            base = sum(S["m"]["pl"][FC.SALES] for S in stores)
            v = (sum(S["m"]["contrib"] for S in stores) if keys == "contrib"
                 else -sum(S["m"]["pl"][k] for S in stores for k in keys))
            c = ws.cell(row=r, column=sh.c0 + n, value=v / base if base else 0.0)
            c.number_format = K.PCT
            c.border = BOX
        r += 1
    # memo
    r += 1
    r = _section(ws, r, "备查项（不参与上表计算）", note_col)
    memo = [
        ("订单口径 GMV（含 IVA，不含已取消）", "Lista de Órdenes · Precio pagado",
         lambda S: S["m"].get("gmv_orders", 0.0),
         "本月创建的订单买家实付金额，含 IVA；与结算口径的差见③桥 A"),
        ("订单数（不含已取消）", "", lambda S: S["m"].get("orders_live_n", 0), ""),
        ("已取消订单金额", "Canceladas", lambda S: S["m"].get("cancel_amt", 0.0), ""),
        ("销项 IVA（收入各行所含，19%）", "IVA débito fiscal", lambda S: S["m"]["iva_out"],
         "卖家应向 SII 申报；平台不代扣"),
        ("进项 IVA（平台费用各行所含）", "IVA crédito fiscal", lambda S: S["m"]["iva_in"],
         "Falabella 费用发票所含 IVA，能否全额抵扣见①会计问题"),
        ("本月结算净额（含 IVA）", "Σ Monto con IVA", lambda S: S["m"]["tx_gross"],
         "＝ 经营贡献毛利 ＋ 销项 IVA − 进项 IVA；对到账见③桥 C"),
    ]
    for label, es, fn, note in memo:
        ws.cell(row=r, column=1, value=label).font = F(10)
        ws.cell(row=r, column=2, value=es).font = F(9, False, K.C_GREY)
        _values_row(ws, r, sh.c0, [fn(S) for S in stores])
        ws.cell(row=r, column=note_col, value=note).font = F(9, False, K.C_GREY)
        r += 1


# ----------------------------------------------------------------- ③ bridges

def sheet_bridge(wb, stores, ctx):
    n = len(stores)
    sh = K.Sheet(wb, SHEETS[2], first_store_col=2, nstores=n)
    ws = sh.ws
    sh.title("平账勾稽表（%s）" % ctx["period_cn"],
             "每座桥最后一行是残差，应为 0。单位：CLP。桥 A、C、D 用含 IVA 金额，桥 B 用不含 IVA 金额（与发票明细一致）。")
    note_col = sh.c0 + n + (1 if n > 1 else 0)
    sh.header(5, ["项目"] + [S["store"] for S in stores] + (["合计"] if n > 1 else []) + ["说明"],
              [52] + [16] * (n + (n > 1)) + [70])
    r = 6

    def line(label, fn, note="", bold=False, color=None):
        nonlocal r
        ws.cell(row=r, column=1, value=label).font = F(10, bold)
        _values_row(ws, r, sh.c0, [fn(S) for S in stores], bold=bold, color=color)
        ws.cell(row=r, column=note_col, value=note).font = F(9, False, K.C_GREY)
        ws.cell(row=r, column=note_col).alignment = wrap()
        r += 1

    A = lambda S, k: (S["m"].get("bridge_a") or {}).get(k, 0.0)
    r = _section(ws, r, "桥 A　订单（下单日）→ 本月结算的商品价款（交易日）", note_col)
    line("本月创建的订单，买家实付（不含已取消）", lambda S: A(S, "orders"), "Lista de Órdenes · Precio pagado")
    line("减：其中本月还没结算的", lambda S: -A(S, "not_settled"),
         "多为月底下单、次月送达的订单（状态：%s）" % per_store(stores, lambda S: "、".join(
             "%s %d" % (k, v) for k, v in ((S["m"].get("bridge_a") or {}).get("not_settled_states") or {}).items())))
    line("加：以前月份的订单，本月结算", lambda S: A(S, "earlier"),
         "按订单创建月：%s" % per_store(stores, lambda S: "、".join(
             "%s %d 件" % (k, v) for k, v in ((S["m"].get("bridge_a") or {}).get("earlier_months") or {}).items())))
    line("加：结算价与订单价的差", lambda S: A(S, "price_diff"), "逐件比较，正常为 0")
    line("＝ 推算的本月结算商品价款", lambda S: A(S, "orders") - A(S, "not_settled") + A(S, "earlier")
         + A(S, "price_diff"), bold=True, color=K.C_SUB)
    line("交易明细：本月“Pago por precio del producto”（含 IVA）", lambda S: A(S, "sales"))
    line("残差（应为 0）", lambda S: A(S, "resid"), bold=True, color=K.C_TOTAL)
    r += 1
    r = _section(ws, r, "桥 B　平台费用发票（Factura falabella.com）→ 交易明细里开了票的行", note_col)
    pids = sorted(set(B["period"] for S in stores for B in S["m"]["bridge_b"]) |
                  set(P["period"] for S in stores for P in S["m"]["bridge_b_pending"]))
    for pid in pids:
        def B(S, k, pid=pid):
            for x in S["m"]["bridge_b"]:
                if x["period"] == pid:
                    return x[k]
            return 0.0

        def P(S, k, pid=pid):
            for x in S["m"]["bridge_b_pending"]:
                if x["period"] == pid:
                    return x[k]
            return 0.0
        a_, b_ = pid.split("_")
        if any(x["period"] == pid for S in stores for x in S["m"]["bridge_b"]):
            line("账期 %s 至 %s：发票明细中交易日在本月的行（不含 IVA）" % (a_, b_), lambda S: B(S, "inv_net"),
                 "InvoiceReport 文件")
            line("　　交易明细中同期、标了发票/贷记单号的行", lambda S: B(S, "tx_net"))
            line("　　残差（应为 0）", lambda S: B(S, "inv_net") - B(S, "tx_net"), bold=True, color=K.C_TOTAL)
            line("　　参考：本账期发票总额（含 IVA）", lambda S: B(S, "factura"), "与卖家后台账期列表一致")
            line("　　参考：本账期贷记单总额（含 IVA）", lambda S: B(S, "nc"))
        else:
            line("账期 %s 至 %s：发票还没出 ⏳" % (a_, b_), lambda S: P(S, "tx_net"),
                 "账期结束后（%s 起）重新下载并重出报表即可核对；本月这几天的费用已按交易明细计入②"
                 % (pd.Timestamp(b_) + pd.Timedelta(days=1)).date(), color=K.C_WARN)
    line("交易明细中不开票的费用（不含 IVA）", lambda S: S["m"]["not_invoiced_net"],
         "没有发票号的费用行（一般为 0）")
    r += 1
    r = _section(ws, r, "桥 C　结算净额（含 IVA）→ 到账", note_col)
    line("经营贡献毛利（不含 IVA，见②）", lambda S: S["m"]["contrib"])
    line("加：销项 IVA", lambda S: S["m"]["iva_out"])
    line("减：进项 IVA", lambda S: -S["m"]["iva_in"])
    line("＝ 本月结算净额（含 IVA）", lambda S: S["m"]["contrib"] + S["m"]["iva_out"] - S["m"]["iva_in"],
         bold=True, color=K.C_SUB)
    line("交易明细含 IVA 合计", lambda S: S["m"]["tx_gross"])
    line("残差（应为 0）", lambda S: S["m"]["contrib"] + S["m"]["iva_out"] - S["m"]["iva_in"]
         - S["m"]["tx_gross"], bold=True, color=K.C_TOTAL)
    line("其中：已付款（Pagado）", lambda S: (S["m"]["paid"].get("Pagado") or (0.0, 0))[0],
         "已进入 Fpay 结算单（Estado de cuenta）；本月涉及 %s 张结算单"
         % per_store(stores, lambda S: str(S["m"]["statements"])))
    line("其中：未付款（No Pagado）", lambda S: (S["m"]["paid"].get("No Pagado") or (0.0, 0))[0],
         "下个结算日才到账")
    r += 1
    r = _section(ws, r, "桥 D　Falabella 以卖家名义开具的销售单据（Liquidación factura）", note_col)
    L = lambda S, k: (S["m"].get("liq") or {}).get(k, 0.0)
    line("Boletas（含 IVA）", lambda S: L(S, "boleta"))
    line("Facturas（含 IVA）", lambda S: L(S, "factura"))
    line("Notas de crédito（含 IVA）", lambda S: L(S, "nc"))
    line("＝ Liquidación 合计", lambda S: L(S, "total"), bold=True, color=K.C_SUB)
    line("参考：本月结算的 FBF 订单商品价款（含 IVA）", lambda S: S["m"]["fbf_sales"],
         "单据在出库/送达时开，结算在送达后，两者有时间差；比较的是量级")
    line("参考：本月结算的卖家自配送（FBS）商品价款（含 IVA）", lambda S: S["m"]["fbs_sales"],
         "这部分的销售单据不在 Liquidación 里 —— 由谁开具请会计师确认（见①问题 1）")


def per_store(stores, fn, sep="；"):
    if len(stores) == 1:
        return fn(stores[0])
    return sep.join("%s %s" % (S["store"], fn(S)) for S in stores)


# ----------------------------------------------------------------- ④ fees

def sheet_fees(wb, stores, ctx):
    ws = wb.create_sheet(SHEETS[3])
    ws["A1"] = "平台费用明细（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, K.C_NAVY)
    ws["A2"] = "本月交易明细里出现的每一种交易类型。归入②的哪一行见“损益表行”。金额 CLP。"
    ws["A2"].font = F(9, False, K.C_GREY)
    cols = ["店铺", "交易类型（Falabella 原文）", "中文", "损益表行", "归类", "笔数", "不含 IVA", "IVA", "含 IVA"]
    hdr = K.Sheet.__new__(K.Sheet)
    hdr.ws = ws
    K.Sheet.header(hdr, 4, cols, [22, 58, 36, 22, 18, 8, 16, 14, 16])
    label = dict((k, zh) for k, zh, _es, _kind, _n in PL_ROWS if k)
    r = 5
    for S in stores:
        t = S["tx"]
        if not len(t):
            continue
        g = t.groupby([T_COLS["type"], "_zh", "_line"]).agg(
            n=(T_COLS["net"], "size"), net=(T_COLS["net"], "sum"), iva=(T_COLS["iva"], "sum"),
            gross=(T_COLS["gross"], "sum")).reset_index()
        g["_order"] = g["_line"].map(lambda k: [x[0] for x in PL_ROWS].index(k)
                                     if k in [x[0] for x in PL_ROWS] else 99)
        for _, x in g.sort_values(["_order", "net"]).iterrows():
            vals = [S["store"], x[T_COLS["type"]], x["_zh"], label.get(x["_line"], x["_line"]),
                    FC.LINE_CATEGORY.get(x["_line"], "⚠ 未登记"), int(x["n"])]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.border = BOX
                c.font = F(10, False, K.C_RED if x["_line"] == "unregistered" else "000000")
            for j, v in enumerate((x["net"], x["iva"], x["gross"]), 7):
                _money(ws.cell(row=r, column=j), float(v))
            r += 1
        ws.cell(row=r, column=1, value="%s 合计" % S["store"]).font = F(10, True)
        for j, v in enumerate((S["m"]["tx_net"], S["m"]["tx_iva"], S["m"]["tx_gross"]), 7):
            c = _money(ws.cell(row=r, column=j), v)
            c.font = F(10, True)
            c.fill = fill(K.C_TOTAL)
        ws.cell(row=r, column=3, value="＝ ② 收入各行 ＋ 费用各行（不含 IVA）").font = F(9, False, K.C_GREY)
        r += 2
    ws.freeze_panes = "A5"


# ----------------------------------------------------------------- ⑤ returns

def sheet_returns(wb, stores, ctx):
    ws = wb.create_sheet(SHEETS[4])
    ws["A1"] = "退货退款分析（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, K.C_NAVY)
    ws["A2"] = ("上半：本月创建的订单里被取消/退货的（订单列表，含 IVA）；下半：本月交易明细里有退货冲回的订单，"
                "它们的全部资金流（不含 IVA）。")
    ws["A2"].font = F(9, False, K.C_GREY)
    _cols(ws, [22, 30, 60, 12, 16, 50])
    r = 4
    for S in stores:
        R = S["returns"]
        r = _section(ws, r, S["store"], 6)
        if "by_status" in R:
            live = S["m"].get("items_n", 0) + S["m"].get("cancel_n", 0)
            for st_, row in R["by_status"].iterrows():
                ws.cell(row=r, column=1, value="订单状态").font = F(10, True)
                ws.cell(row=r, column=2, value=st_)
                ws.cell(row=r, column=3, value="占本月订单件数 %s" % pct(row["count"] / live if live else 0))
                ws.cell(row=r, column=4, value=int(row["count"])).border = BOX
                _money(ws.cell(row=r, column=5), float(row["sum"]))
                r += 1
            r += 1
            for (st_, reason), row in R["by_reason"].head(25).iterrows():
                ws.cell(row=r, column=1, value="原因")
                ws.cell(row=r, column=2, value=st_)
                ws.cell(row=r, column=3, value=reason).alignment = wrap()
                ws.cell(row=r, column=4, value=int(row["count"])).border = BOX
                _money(ws.cell(row=r, column=5), float(row["sum"]))
                r += 1
            r += 1
            for (st_, ful), cnt in R["by_fulfilment"].items():
                ws.cell(row=r, column=1, value="配送方式")
                ws.cell(row=r, column=2, value=st_)
                ws.cell(row=r, column=3, value=ful)
                ws.cell(row=r, column=4, value=int(cnt)).border = BOX
                r += 1
        else:
            ws.cell(row=r, column=1, value="没有订单列表，上半部分无法计算").font = F(10, False, K.C_RED)
            r += 1
        r += 1
        ws.cell(row=r, column=1, value="退货订单的资金流").font = F(10, True)
        ws.cell(row=r, column=3, value="本月有退货冲回的订单 %d 张" % R["orders"])
        r += 1
        if len(R["money"]):
            for (line_, zh), row in R["money"].iterrows():
                ws.cell(row=r, column=2, value=dict((k, z) for k, z, *_ in PL_ROWS if k).get(line_, line_))
                ws.cell(row=r, column=3, value=zh)
                ws.cell(row=r, column=4, value=int(row["count"])).border = BOX
                _money(ws.cell(row=r, column=5), float(row["sum"]))
                r += 1
            ws.cell(row=r, column=3, value="退货订单本月净影响（不含 IVA）").font = F(10, True)
            c = _money(ws.cell(row=r, column=5), float(R["money"]["sum"].sum()))
            c.font = F(10, True)
            c.fill = fill(K.C_TOTAL)
            ws.cell(row=r, column=6, value="正数说明这些订单的原始货款也在本月结算；负数是退货净损失").font = F(9, False, K.C_GREY)
            r += 1
        r += 2


def sheet_not_applicable(wb, name, title, text):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = F(14, True, K.C_NAVY)
    ws["A3"] = text
    ws["A3"].font = F(11, False, K.C_GREY)
    ws.column_dimensions["A"].width = 120


# ----------------------------------------------------------------- ⑦ SKU

def sheet_sku(wb, stores, ctx):
    ws = wb.create_sheet(SHEETS[6])
    ws["A1"] = "SKU 盈利分析（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, K.C_NAVY)
    ws["A2"] = ("按交易明细的 SKU 汇总（不含 IVA）。广告、仓储等没有 SKU 的店铺级费用单列一行。"
                "黄色列填单位采购成本（不含 IVA），右边两列自动算出 SKU 毛利。")
    ws["A2"].font = F(9, False, K.C_GREY)
    cols = ["店铺", "SKU", "商品", "结算件数", "退货件数", "商品销售收入", "运费收入", "退货退款",
            "平台促销", "佣金", "物流（净）", "Fulfillment", "其他费用", "贡献（未扣成本）",
            "单位成本（录入）", "商品成本", "SKU 毛利"]
    hdr = K.Sheet.__new__(K.Sheet)
    hdr.ws = ws
    K.Sheet.header(hdr, 4, cols, [18, 18, 44, 9, 9, 14, 12, 12, 12, 12, 12, 12, 12, 14, 12, 14, 14])
    r = 5
    ship = [FC.SHIP_COFIN, FC.SHIP_BUYER, FC.SHIP_PROMO, FC.SHIP_REVERSE]
    other = [FC.ADS, FC.SUBSCRIPTION, FC.PENALTY, FC.OTHER, "unregistered", FC.OTHER_INCOME]
    for S in stores:
        g = S["sku"]
        for sku, x in g.iterrows():
            vals = [S["store"], sku, x["name"], int(x["units"]), int(x["returns"])]
            for j, v in enumerate(vals, 1):
                ws.cell(row=r, column=j, value=v).border = BOX
            money = [x[FC.SALES], x[FC.SHIP_INCOME], x[FC.REFUND], x[FC.PROMO], x[FC.COMMISSION],
                     sum(x[k] for k in ship), x[FC.FULFILMENT], sum(x[k] for k in other), x["contrib"]]
            for j, v in enumerate(money, 6):
                _money(ws.cell(row=r, column=j), float(v))
            c = _money(ws.cell(row=r, column=15), 0)
            c.fill = fill(K.C_INPUT)
            c = ws.cell(row=r, column=16, value="=-O%d*D%d" % (r, r))
            c.number_format = MNY
            c.fill = fill(K.C_CALC)
            c = ws.cell(row=r, column=17, value="=N%d+P%d" % (r, r))
            c.number_format = MNY
            c.fill = fill(K.C_CALC)
            r += 1
        r += 1
    ws.freeze_panes = "D5"
    ws.auto_filter.ref = "A4:Q%d" % max(4, r - 1)


# ----------------------------------------------------------------- ⑧ trend

def sheet_trend(wb, stores, ctx):
    ws = wb.create_sheet(SHEETS[7])
    ws["A1"] = "月度趋势（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, K.C_NAVY)
    ws["A2"] = ("本月按周（周一开始，按交易日，不含 IVA）。往月趋势需要往月的下载：同一台机器上有更早月份的"
                "Falabella 下载目录时，下方自动追加各月合计。")
    ws["A2"].font = F(9, False, K.C_GREY)
    cols = ["店铺", "周（起始日）", "结算件数", "商品销售收入", "退货退款", "平台费用", "贡献（未扣成本）", "费用率"]
    hdr = K.Sheet.__new__(K.Sheet)
    hdr.ws = ws
    K.Sheet.header(hdr, 4, cols, [22, 16, 10, 16, 14, 16, 16, 10])
    r = 5
    for S in stores:
        t = S["tx"]
        if not len(t):
            continue
        wk = t["_date"] - pd.to_timedelta(t["_date"].dt.weekday, unit="D")
        t = t.assign(_wk=wk.dt.date)
        for w, x in t.groupby("_wk"):
            sales = x.loc[x["_line"] == FC.SALES, T_COLS["net"]].sum()
            fees = x.loc[x["_line"].isin(FEE_LINES), T_COLS["net"]].sum()
            row = [S["store"], str(w), int((x["_line"] == FC.SALES).sum())]
            for j, v in enumerate(row, 1):
                ws.cell(row=r, column=j, value=v).border = BOX
            for j, v in enumerate((sales, x.loc[x["_line"] == FC.REFUND, T_COLS["net"]].sum(), fees,
                                   x[T_COLS["net"]].sum()), 4):
                _money(ws.cell(row=r, column=j), float(v))
            c = ws.cell(row=r, column=8, value=(-fees / sales) if sales else 0.0)
            c.number_format = K.PCT
            r += 1
        r += 1
    hist = ctx.get("history") or {}
    if hist:
        r = _section(ws, r + 1, "往月（来自本机更早的 Falabella 下载）", 8)
        for store, months in hist.items():
            for mon, v in sorted(months.items()):
                ws.cell(row=r, column=1, value=store)
                ws.cell(row=r, column=2, value=mon)
                ws.cell(row=r, column=3, value=v["items"])
                _money(ws.cell(row=r, column=4), v["sales"])
                _money(ws.cell(row=r, column=5), v["refund"])
                _money(ws.cell(row=r, column=6), v["fees"])
                _money(ws.cell(row=r, column=7), v["contrib"])
                c = ws.cell(row=r, column=8, value=(-v["fees"] / v["sales"]) if v["sales"] else 0.0)
                c.number_format = K.PCT
                r += 1


def history(stores, base):
    """Monthly totals of earlier months, from other Falabella folders of the same store."""
    out = {}
    for S in stores:
        root = os.path.dirname(S["folder"])
        seen = {}
        for d in sorted(glob.glob(os.path.join(root, "*_m??????"))):
            mon = d[-6:-2] + "-" + d[-2:]
            if mon >= S["month"] or mon in seen:
                continue
            p = _first(d, "NewReportTransaction_*.xlsx")
            if not p:
                continue
            try:
                t, _ = load_transactions(p)
                a, b = periods.month_bounds(mon)
                t = t[(t["_date"] >= pd.Timestamp(a)) & (t["_date"] <= pd.Timestamp(b))]
                ln = t[T_COLS["type"]].map(lambda n: (FC.lookup(n) or ("", "unregistered"))[1])
                seen[mon] = {"items": int((ln == FC.SALES).sum()),
                             "sales": float(t.loc[ln == FC.SALES, T_COLS["net"]].sum()),
                             "refund": float(t.loc[ln == FC.REFUND, T_COLS["net"]].sum()),
                             "fees": float(t.loc[ln.isin(FEE_LINES), T_COLS["net"]].sum()),
                             "contrib": float(t[T_COLS["net"]].sum())}
            except Exception:
                continue
        if seen:
            out[S["store"]] = dict(sorted(seen.items())[-6:])
    return out


# ----------------------------------------------------------------- ⑨ inventory

def sheet_inventory(wb, stores, ctx):
    ws = wb.create_sheet(SHEETS[8])
    ws["A1"] = "库存与动销分析（Fulfillment by Falabella）"
    ws["A1"].font = F(14, True, K.C_NAVY)
    ws["A2"] = ("库存快照取自下载当时（不是月末），近 4 周销量是平台给的。可售周数 = 库存 ÷ (近 4 周销量 ÷ 4)。"
                "滞销：有库存、近 4 周没卖；断货风险：近 4 周有销量、可售不到 2 周。")
    ws["A2"].font = F(9, False, K.C_GREY)
    cols = ["店铺", "SKU 卖家", "SKU Falabella", "商品", "售价（含 IVA）", "可售库存", "近 4 周销量",
            "可售周数", "本月结算件数", "库存货值（售价）", "判断"]
    hdr = K.Sheet.__new__(K.Sheet)
    hdr.ws = ws
    K.Sheet.header(hdr, 4, cols, [20, 18, 14, 50, 14, 10, 10, 10, 12, 16, 16])
    r = 5
    for S in stores:
        s = S["stock"]
        if s is None:
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=4, value="本月没有下载到库存快照").font = F(10, False, K.C_RED)
            r += 2
            continue
        sold = S["sku"]["units"] if len(S["sku"]) else pd.Series(dtype=int)
        for _, x in s.iterrows():
            stock, w4 = float(x.get("Stock Disponible") or 0), float(x.get("Ventas ultimas 4 semanas") or 0)
            weeks = stock / (w4 / 4.0) if w4 else None
            verdict = ("滞销" if stock > 0 and not w4 else "断货风险" if w4 and (weeks or 0) < 2
                       else "无库存" if not stock else "")
            vals = [S["store"], x.get("SKU Vendedor"), x.get("SKU GSC"), x.get("Nombre"),
                    float(x.get("Precio") or 0), int(stock), int(w4),
                    round(weeks, 1) if weeks is not None else None,
                    int(sold.get(x.get("SKU Vendedor"), 0)) if len(sold) else 0,
                    stock * float(x.get("Precio") or 0), verdict]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.border = BOX
                if j in (5, 10):
                    c.number_format = MNY
                if verdict in ("滞销", "断货风险") and j == 11:
                    c.font = F(10, True, K.C_AMBER)
            r += 1
        r += 1
    ws.freeze_panes = "E5"
    ws.auto_filter.ref = "A4:K%d" % max(4, r - 1)


# ----------------------------------------------------------------- ⑩ gaps

def sheet_gaps(wb, stores, ctx):
    ws = wb.create_sheet(SHEETS[9])
    ws["A1"] = "数据缺口与待办（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, K.C_NAVY)
    ws["A2"] = "error = 数字在修好之前不能信；warn = 请看一眼，数字仍成立；skipped = 不适用。空表示没有待办。"
    ws["A2"].font = F(9, False, K.C_GREY)
    hdr = K.Sheet.__new__(K.Sheet)
    hdr.ws = ws
    K.Sheet.header(hdr, 4, ["店铺", "校验", "级别", "结果", "说明"], [22, 46, 9, 9, 100])
    r = 5
    order = sorted(ctx["checks"], key=lambda c: (c["ok"] or c["skipped"], c["severity"] != "error"))
    for c in order:
        res = "跳过" if c["skipped"] else ("通过" if c["ok"] else "未通过")
        color = K.C_OK if (c["ok"] or c["skipped"]) else (K.C_BAD if c["severity"] == "error" else K.C_WARN)
        for j, v in enumerate((c["store"], c["name"], c["severity"], res, c["detail"]), 1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = BOX
            cell.fill = fill(color)
            cell.alignment = wrap()
        r += 1
    r += 1
    r = _section(ws, r, "没有下载到的文件及原因（来自下载目录的 run_meta.json）", 5)
    any_missing = False
    for S in stores:
        for rec in store_meta_missing(S):
            any_missing = True
            vals = (S["store"], rec.get("form"), rec.get("cause"),
                    "影响报表" if rec.get("in_report", True) else "报表不读",
                    "%s　下一步：%s" % (rec.get("explain") or "", rec.get("action") or ""))
            for j, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=j, value=v)
                cell.border = BOX
                cell.alignment = wrap()
            r += 1
    if not any_missing:
        ws.cell(row=r, column=2, value="没有。应下的文件全部到齐。")
        r += 1
    r += 1
    r = _section(ws, r, "要请会计师确认的问题", 5)
    for i, q in enumerate(ACCOUNTANT_QUESTIONS, 1):
        ws.cell(row=r, column=2, value="问题 %d" % i).font = F(10, True)
        c = ws.cell(row=r, column=5, value=q)
        c.alignment = wrap()
        ws.row_dimensions[r].height = K.est_height(q, 95)
        r += 1


# ======================================================================= entry

def build_context(stores, month):
    y, mo = periods.parse_month(month)
    return {"month": month, "period_cn": "%d 年 %d 月" % (y, mo),
            "who": stores[0]["store"] if len(stores) == 1 else "%d 家店铺汇总" % len(stores),
            "built": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}


def build(pairs, month, out_path, strict=False):
    """pairs: [(store, folder)]. Returns a result dict (outcome, checks, out)."""
    stores = [derive(load_store(s, f, month)) for s, f in pairs]
    ctx = build_context(stores, month)
    ctx["checks"] = validate(stores)
    ctx["history"] = history(stores, None)
    bad = [c for c in ctx["checks"] if not c["ok"] and not c["skipped"] and c["severity"] == "error"]
    if strict and bad:
        return {"outcome": "strict_aborted", "checks": ctx["checks"], "out": out_path}
    build_workbook(stores, ctx, out_path)
    return {"outcome": "checks_failed" if bad else "ok", "checks": ctx["checks"], "out": out_path,
            "metrics": {S["store"]: S["m"] for S in stores}}
