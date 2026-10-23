# -*- coding: utf-8 -*-
"""
Falabella monthly financial workbook — the sister of MercadoLibre's.

Same ten sheets, same look, same check levels: the layout is written by
core/report_family.py on the frame of core/report_kit.py (docs/REPORT_FAMILY.md);
the content — lines, bridges, columns, rules — is Falabella's. Accounting rules are in
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
  cost column in ⑦ — filled from the shared SKU cost table (core/sku_costs.py) when
  config.json names one.

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

import warnings

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from core import report_family as RF
from core import report_kit as K
from core import sku_costs as SC
from core.report_kit import F, fill, wrap, BOX
from platforms.falabella import fee_catalog as FC
from platforms.falabella import periods

# Falabella's exports carry no default cell style; openpyxl says so for every file it
# reads, and the runner would show each of those lines as a warning in the run log.
warnings.filterwarnings("ignore", message="Workbook contains no default style",
                        category=UserWarning)

SHEETS = K.SHEET_NAMES
PLATFORM, COUNTRY = "FA", "CL"  # core/markets.py codes (and the SKU cost table's 平台 / 国家)
CUR = "CLP"
NO_SKU = "(无 SKU · 店铺级费用)"
# which units carry a cost of goods (core/sku_costs, README "Accounting rules")
COST_UNITS_BASIS = ("本月结算件数（“Pago por precio del producto”的行数）− 本月退货件数"
                    "（“Descuento por devolución de producto”的行数，退回的商品回到库存）")
DECIMALS = 0                    # CLP has no cents
MNY = K.money_format(DECIMALS)
IVA_RATE = 0.19
# The cost table holds costs net of IVA; this report is with IVA, so each unit cost is
# multiplied by 1 + IVA (core/sku_costs.apply factor).
COST_FACTOR = 1 + IVA_RATE
COST_IVA_NOTE = "（成本表的成本不含 IVA，本报表 × 1.19 换算为含 IVA）"
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
# Every amount of the report is Monto con IVA (the accountant's choice, 2026-10-21).
# Net and IVA are still read: for the IVA checks and ②'s "其中含 IVA" memo line.
AMT = T_COLS["gross"]

# ② fixed rows (core/report_family.PLRow): key, 中文, 平台原文, kind, 说明, formula.
# kind: d 取数 / sum 小计 / f 计算 / in 录入. A formula lists the rows it adds
# ("-key" subtracts); Span(a, b) sums every row from a to b.
_R = RF.PLRow
PL_ROWS = [
    _R(FC.SALES, "商品销售收入", "Pago por precio del producto", "d",
     "本月结算的商品价款（含 IVA），按交易日（≈ 送达日）归集；含 CMR 积分支付的部分"),
    _R("ship_net", "运费收入（净额）", "Pago de envío comprador − Reversa / Pago por envío directo", "d",
     "卖家自己挣的运费：自配送订单买家付的运费，以及平台付给卖家的自配送运费。Falabella 配送的订单，"
     "买家付的运费原样转付 Falabella —— 卖家只是代收代付（代理人），按 IFRS 15 不计收入，两边金额见下方备查项"),
    _R("revenue", "营业收入合计", "", "sum", "＝ 商品销售收入 ＋ 运费收入（净额）",
       (FC.SALES, "ship_net")),
    _R(FC.REFUND, "减：退货退款", "Descuento por devolución de producto", "d",
     "退货冲回的商品价款（含申诉裁定）"),
    _R(FC.PROMO, "加：平台出资促销与激励", "Aportes promocionales / Oportunidad Única", "d",
     "Falabella 出资的优惠券、CMR 卡活动补贴、销售目标奖励"),
    _R("net_sales", "净销售收入", "", "f", "＝ 营业收入合计 − 退货退款 ＋ 平台出资促销",
       ("revenue", FC.REFUND, FC.PROMO)),
    _R(FC.COMMISSION, "减：销售佣金", "Cobro por comisión por venta", "d",
     "按品类费率；退货订单的佣金退还已抵减"),
    _R(FC.SHIP_COFIN, "减：配送共担费", "Cobro por cofinanciamiento logístico", "d",
     "卖家承担的配送费（FBF 与自配送订单）"),
    _R(FC.SHIP_PROMO, "减：免运费促销", "Promo envío falabella.com / Promo Envio Seller", "d",
     "平台免运促销先扣后还（净额通常为 0）；卖家出资的免运促销"),
    _R(FC.SHIP_REVERSE, "减：退货逆向物流", "Cobro por logística inversa", "d", "退货的取回运费"),
    _R(FC.FULFILMENT, "减：Fulfillment 仓储与操作", "Fulfillment by Falabella", "d",
     "FBF 仓储、超期仓储、拣货打包、揽收、取回库存等"),
    _R(FC.ADS, "减：广告推广", "Productos patrocinados / Banners …", "d", "店铺级费用，无法分摊到订单"),
    _R(FC.SUBSCRIPTION, "减：订阅与增值服务", "Plan marketplace 360 / Mi tienda …", "d", ""),
    _R(FC.PENALTY, "减：罚款与违规扣费", "Comisión por cancelación / Penalty …", "d",
     "卖家原因取消的佣金罚款、尺寸申报不实等"),
    _R(FC.OTHER, "减：其他调整", "Corrección de pagos / cobros", "d", ""),
    _R("unregistered", "减：未登记科目 ⚠", "(fee_catalog 里没有)", "d",
     "字典里没有的交易类型。金额照常计入以保证勾稽，但需人工确认科目归属，见 ⑩ 页"),
    _R("fee_total", "平台费用小计", "", "sum", "从结算款中直接扣除", RF.Span(FC.COMMISSION, "unregistered")),
    _R(FC.OTHER_INCOME, "加：其他收入（物流丢失/损坏赔付）", "Compensación por extravío o producto dañado", "d",
     "平台对丢失/损坏商品的赔付，是存货损失的补偿，不是销售收入，单列在费用之后"),
    _R("contrib", "经营贡献毛利（未扣商品成本）", "", "f", "＝ 净销售收入 − 平台费用小计 ＋ 其他收入",
       ("net_sales", "fee_total", FC.OTHER_INCOME)),
    _R("cogs", "减：商品采购成本", "", "in",
     "⚠ 平台数据不含此项：在【⑦ SKU盈利分析】填单位成本，或直接在此填总额"),
    _R("profit", "税前净利润", "", "f", "＝ 经营贡献毛利 − 商品采购成本", ("contrib", "-cogs")),
]
PL_LABELS = dict((x.key, x.label) for x in PL_ROWS)
INCOME_LINES = (FC.SALES, FC.SHIP_INCOME, FC.REFUND, FC.PROMO, FC.OTHER_INCOME)
# FEE_ROWS: the fee rows of ②. FEE_LINES: every charge line, incl. the buyer-shipping
# pass-through that ② nets against shipping income.
FEE_ROWS = (FC.COMMISSION, FC.SHIP_COFIN, FC.SHIP_PROMO, FC.SHIP_REVERSE, FC.FULFILMENT, FC.ADS,
            FC.SUBSCRIPTION, FC.PENALTY, FC.OTHER, "unregistered")
LOGISTICS_ROWS = (FC.SHIP_COFIN, FC.SHIP_PROMO, FC.SHIP_REVERSE)
FEE_LINES = (FC.COMMISSION, FC.SHIP_COFIN, FC.SHIP_BUYER, FC.SHIP_PROMO, FC.SHIP_REVERSE,
             FC.FULFILMENT, FC.ADS, FC.SUBSCRIPTION, FC.PENALTY, FC.OTHER, "unregistered")

# The accounting policies the report applies (decided, not open questions), and what
# the tax filing has to take into account. Shown on ① and ⑩; README.md explains them.
ACCOUNTING_POLICIES = [
    ("收入确认", "商品价款在 Falabella 记“Pago por precio del producto”时确认 —— 即送达、控制权转给买家时"
     "（IFRS 15），按交易日归月。已下单、未送达的订单不确认收入（金额见③桥 A）。"),
    ("含 IVA", "按会计师的要求，全部金额用含 IVA 金额（Monto con IVA），与 Falabella 后台、实际到账一致；"
     "每一行的 IVA 都是 19%，所以各项费用率、毛利率与不含税口径相同。损益表的金额里含有应缴 SII 的 IVA"
     "（销项 − 进项），金额见②备查项“其中含 IVA”；商品成本同样按含 IVA 计（成本表的不含税成本 × 1.19）。"),
    ("运费", "Falabella 配送的订单，买家付的运费原样转付 Falabella，卖家只是代收代付（代理人），按净额处理、"
     "不计收入；卖家自配送订单的运费、以及平台付给卖家的自配送运费才是卖家的收入。"),
    ("促销", "Falabella 出资的优惠券、CMR 卡“Oportunidad Única”补贴和销售奖励，是这笔销售的对价的一部分，"
     "计入净销售收入；卖家自己出资的免运费促销是销售费用。"),
    ("退货", "退货冲减收入；平台退还的佣金冲减佣金；退货的逆向物流计入物流费用。"),
    ("丢失/损坏赔付", "是对存货损失的补偿，不是销售：单列为“其他收入”，放在平台费用之后。"),
    ("平台费用", "按 Falabella 的扣费日计入当月（与平台费用发票的口径相同）；佣金、物流、仓储、广告、订阅、"
     "罚款分行列示，同一科目的退还/更正冲减同一行。"),
    ("商品成本与存货", "平台数据不含采购成本：在②或⑦录入。⑨的“库存货值”按售价计算，仅用于判断库存规模，"
     "不是会计上的存货价值。"),
]
TAX_NOTES = [
    "销项 IVA 的单据：Fulfillment by Falabella 订单的 boleta/factura 由 Falabella 以卖家名义开具（Liquidación "
    "factura，见③桥 D），申报以这些单据为准，不要重复开具；卖家自配送（FBS）订单的单据须由卖家自己开具"
    "（本月金额见③桥 D 最后一行）。",
    "进项 IVA 的期间：Falabella 的费用发票按账期（21 日–次月 20 日）在账期结束后开出，进项 IVA 按发票日期"
    "申报 —— 本月 21–30 日费用所含的进项 IVA 进入下个月的申报。这只影响 IVA 申报的月份，不影响本报表的损益。",
    "促销补贴与赔付的单据由 Falabella 开具（Nota de crédito / Nota de débito promociones），以平台单据为准。",
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


def _same_month(folder):
    """This folder first, then the store's other folders of the same month, newest
    first. A later `run_batch --only stock` (or a rerun of one route) writes a new
    folder with just those files; the report should still find them."""
    root, name = os.path.split(os.path.abspath(folder))
    suffix = name[-8:] if re.search(r"_m\d{6}$", name) else None
    sib = sorted(glob.glob(os.path.join(root, "*" + suffix)), reverse=True) if suffix else []
    return [folder] + [d for d in sib if os.path.abspath(d) != os.path.abspath(folder)]


def _first(folder, pattern):
    for d in _same_month(folder):
        hits = sorted(glob.glob(os.path.join(d, pattern)))
        if hits:
            return hits[-1]
    return None


def _all(folder, pattern):
    """All files of one kind from ONE folder (the first that has any): order
    pieces or invoice periods must not be mixed across runs."""
    for d in _same_month(folder):
        hits = sorted(glob.glob(os.path.join(d, pattern)))
        if hits:
            return hits
    return []


def rel(path, folder):
    """File name, plus its folder when it comes from another run of the same month."""
    if not path:
        return ""
    d = os.path.dirname(os.path.abspath(path))
    return os.path.basename(path) if d == os.path.abspath(folder) else \
        "%s（取自 %s）" % (os.path.basename(path), os.path.basename(d))


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


def empty_orders():
    o = pd.DataFrame({c: pd.Series(dtype="object") for c in ORDER_COLS})
    for c in ("Precio pagado", "Precio unitario", "Costo de envio", "Cupon/Descuentos", "Impuesto"):
        o[c] = pd.Series(dtype="float64")
    o["_created"] = pd.Series(dtype="datetime64[ns]")
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
    # What the downloader found to be empty ("No tienes órdenes", no Liquidación row):
    # an empty month, not a missing file.
    no_data = set(r.get("form") for r in (S["meta"].get("missing") or [])
                  if r.get("cause") == "NO_DATA")
    # The month's orders come in pieces (Ordenes_<from>_<to>.xlsx). Each piece is taken
    # from the newest run of the month that has it: a rerun in which one piece failed
    # must not lose it when an earlier run has it. A piece neither downloaded nor
    # reported empty ("No tienes órdenes") is missing — an error, never silently fewer
    # orders.
    ops, S["orders_missing"] = [], []
    for x, y in periods.order_chunks(month):
        pid = periods.period_id(x, y)
        p = _first(folder, "Ordenes_%s.xlsx" % pid)
        if p:
            ops.append(p)
        elif "order/%s" % pid not in no_data:
            S["orders_missing"].append(pid)
    S["files"]["order"] = ops
    S["orders_none"] = not ops and not S["orders_missing"]
    S["orders_all"] = load_orders(ops) if ops else (empty_orders() if S["orders_none"] else None)
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
    S["liq_none"] = not lp and "billing2" in no_data
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
    m["pl"] = {k: float(by[AMT].sum().get(k, 0.0)) for k in INCOME_LINES + FEE_LINES}
    m["pl"]["ship_net"] = m["pl"][FC.SHIP_INCOME] + m["pl"][FC.SHIP_BUYER]
    m["income"] = sum(m["pl"][k] for k in INCOME_LINES)
    m["revenue"] = m["pl"][FC.SALES] + m["pl"]["ship_net"]
    m["net_sales"] = m["revenue"] + m["pl"][FC.REFUND] + m["pl"][FC.PROMO]
    m["fees"] = sum(m["pl"][k] for k in FEE_ROWS)
    m["contrib"] = m["net_sales"] + m["fees"] + m["pl"][FC.OTHER_INCOME]
    m["tx_net"] = float(t[T_COLS["net"]].sum())
    m["tx_iva"] = float(t[T_COLS["iva"]].sum())
    m["tx_gross"] = float(t[T_COLS["gross"]].sum())
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
                     "tx_gross": float(tw[T_COLS["gross"]].sum())})
    m["bridge_b"], m["bridge_b_pending"] = Bs, pend
    open_days = pd.Series(False, index=t.index)
    for _pid, lo, hi in pending:
        open_days |= (t["_date"] >= pd.Timestamp(lo)) & (t["_date"] <= pd.Timestamp(hi))
    m["not_invoiced"] = float(t.loc[t[T_COLS["doc"]].isna() & t["_line"].isin(FEE_LINES)
                                    & ~open_days, AMT].sum())
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
    sk = t[T_COLS["sku"]].fillna(NO_SKU)
    g = t.assign(_sku=sk).pivot_table(index="_sku", columns="_line", values=AMT,
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
    out["money"] = r.groupby(["_line", "_zh"])[AMT].agg(["count", "sum"]) \
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
        miss = S.get("orders_missing") or []
        C.append(K.check("file_order", st, "订单列表齐全",
                         (bool(f.get("order")) or S.get("orders_none")) and not miss,
                         ("缺 %s 这几段的订单（没下到，也不是“没有订单”）：订单口径 GMV、桥 A、退货分析"
                          "少了这几天，重跑这家店的下载" % "、".join(m.replace("_", "–") for m in miss))
                         if miss else
                         "%d 段（%s）" % (len(f["order"]), "、".join(
                             os.path.basename(p)[8:-5].replace("_", "–") for p in f["order"]))
                         if f.get("order") else
                         "本月没有订单（下载时平台显示 No tienes órdenes）",
                         severity="error"))
        C.append(K.check("file_stock", st, "库存快照在", bool(f.get("stock")),
                         os.path.basename(f["stock"]) if f.get("stock") else
                         "没有 Product_details：⑨ 库存页为空", severity="warn"))
        C.append(K.check("file_billing2", st, "Liquidación factura 在",
                         bool(f.get("billing2")) or S.get("liq_none"),
                         os.path.basename(f["billing2"]) if f.get("billing2") else
                         "本月没有平台代开的销售单据（列表里没有这个月）" if S.get("liq_none") else
                         "没有 Liquidación：桥 D 无法核对平台代开的销售单据", severity="warn"))
        C.append(K.check("tx_in_month", st, "交易明细只取本月",
                         m["lines"] > 0,
                         "本月 %d 行；文件里另有 %d 行不在本月，已排除"
                         % (m["lines"], S.get("tx_outside") or 0) if m["lines"] else
                         "交易明细里没有任何一行在本月"))
        C.append(K.check("unregistered", st, "交易类型全部已登记", not m["unregistered_types"],
                         "全部 %d 种已登记" % S["tx"]["_line"].nunique() if not m["unregistered_types"]
                         else "未登记：%s（金额已计入“未登记科目”行）" % "；".join(m["unregistered_types"])))
        C.append(K.check("pl_tie", st, "损益表各行合计 ＝ 交易明细合计（含 IVA）",
                         abs(m["contrib"] - m["tx_gross"]) <= TOL,
                         "损益 %s，明细 %s，差 %s" % (fmt(m["contrib"]), fmt(m["tx_gross"]),
                                                  fmt(m["contrib"] - m["tx_gross"]))))
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
            d = B["inv_gross"] - B["tx_gross"]
            C.append(K.check("bridge_b", st, "桥 B：费用发票 %s ↔ 交易明细（%s–%s）"
                             % (B["period"], B["from"], B["to"]), abs(d) <= TOL,
                             "发票 %d 行 %s，明细 %d 行 %s，差 %s"
                             % (B["inv_lines"], fmt(B["inv_gross"]), B["tx_lines"], fmt(B["tx_gross"]),
                                fmt(d))))
        for P in m["bridge_b_pending"]:
            if not P["tx_lines"]:
                C.append(K.check("bridge_b_open", st, "桥 B：账期 %s 的费用发票还没出" % P["period"],
                                 True, "%s–%s 没有要开票的费用，不用等这张发票" % (P["from"], P["to"]),
                                 severity="warn", skipped=True))
                continue
            C.append(K.check("bridge_b_open", st, "桥 B：账期 %s 的费用发票还没出" % P["period"],
                             False, "%s–%s 的 %d 行应开票的费用 %s（含 IVA）暂无发票可核对。"
                             "这些费用已按交易明细计入②，数字不受影响；发票在账期结束后（次月 21 日起）开出，"
                             "下个月的下载会取到它，无需人工处理" % (P["from"], P["to"], P["tx_lines"],
                                                   fmt(P["tx_gross"])), severity="warn"))
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
#
# The layout comes from core/report_family.py (the family's sheet writers); what
# follows is Falabella's content: which lines, which bridges, which columns.

def build_workbook(stores, ctx, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    fam = RF.Family(stores, money=MNY)
    sheet_readme(wb, fam, ctx)
    sheet_pl(wb, fam, ctx)
    sheet_bridge(wb, fam, ctx)
    sheet_fees(wb, fam, ctx)
    sheet_returns(wb, fam, ctx)
    RF.not_applicable_sheet(wb, SHEETS[5], "退货质检明细",
                            "不适用：Falabella 不提供逐件退货质检报表。退货的金额与原因见【%s】。"
                            % SHEETS[4])
    sheet_sku(wb, fam, ctx)
    sheet_trend(wb, fam, ctx)
    sheet_inventory(wb, fam, ctx)
    sheet_gaps(wb, fam, ctx)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    return out_path


def pl_label(key):
    """② label of a P&L line key; the key itself when it is not a row of ②."""
    return PL_LABELS.get(key, key)


# ----------------------------------------------------------------- ① readme

def sheet_readme(wb, fam, ctx):
    stores = fam.stores
    sources = []
    for S in stores:
        f = S["files"]
        meta = S["meta"]
        sources.append((S["store"], "下载于 %s；卖家 %s %s。交易明细：%s（平台生成于 %s）；订单：%d 个文件；"
                        "费用发票明细：%s；Liquidación：%s；库存：%s；PDF %d 份。"
                        % (meta.get("downloaded_at") or "?", (meta.get("flags") or {}).get("seller") or "",
                           (meta.get("flags") or {}).get("seller_name") or "",
                           os.path.basename(f["billing"]) if f.get("billing") else "缺",
                           S.get("tx_updated") or "?", len(f.get("order") or []),
                           "、".join(os.path.basename(p) for p in f.get("billing1") or []) or "缺",
                           os.path.basename(f["billing2"]) if f.get("billing2") else "缺",
                           rel(f.get("stock"), S["folder"]) if f.get("stock") else "缺",
                           len(f.get("pdf") or []))))
    tot = lambda k: sum(S["m"].get(k, 0.0) or 0.0 for S in stores)
    rev = sum(S["m"]["pl"][FC.SALES] for S in stores)
    points = [("规模", "本月结算商品价款 %s（含 IVA），%d 件；按下单日的 GMV（含 IVA、不含已取消）%s，"
               "%d 张订单。" % (fmt(rev), int(tot("sales_items")), fmt(tot("gmv_orders")),
                             int(tot("orders_live_n"))))]
    if rev:
        points.append(("费用率", "平台费用合计 %s，占商品销售收入 %s；其中佣金 %s、物流 %s、Fulfillment %s、广告 %s。"
                       % (fmt(-tot("fees")), pct(-tot("fees") / rev),
                          pct(-sum(S["m"]["pl"][FC.COMMISSION] for S in stores) / rev),
                          pct(-sum(S["m"]["pl"][k] for S in stores for k in LOGISTICS_ROWS) / rev),
                          pct(-sum(S["m"]["pl"][FC.FULFILMENT] for S in stores) / rev),
                          pct(-sum(S["m"]["pl"][FC.ADS] for S in stores) / rev))))
        points.append(("贡献毛利", "经营贡献毛利（未扣商品成本）%s，占商品销售收入 %s。"
                       % (fmt(tot("contrib")), pct(tot("contrib") / rev))))
    points.append(("其中含 IVA", "以上金额含 IVA。经营贡献毛利里含应缴 SII 的 IVA 净额 %s（销项 − 进项，按交易日）。"
                   % fmt(tot("tx_iva"))))
    costs = ctx.get("sku_costs")
    policies = list(ACCOUNTING_POLICIES)
    if costs is not None:
        sources.append(SC.source_line(costs))
        results = [S.get("cost") for S in stores]
        points.append(SC.point_line(results, tot("contrib") - sum(r.cogs for r in results if r)))
        policies.append(SC.policy_line(COST_UNITS_BASIS, COST_IVA_NOTE.strip("（）") + "。" + SC.STOCK_NOT_VALUED))
    sections = [
        ("一、这份报表回答什么", [
            ("本月赚了多少", "【%s】从商品销售收入推到“经营贡献毛利（未扣商品成本）”，"
             "填入商品成本后得到税前净利润。" % SHEETS[1]),
            ("钱对不对得上", "【%s】四座桥：A 订单 → 结算、B 平台费用发票 → 交易明细、"
             "C 结算净额 → 到账、D 平台代开的销售单据。" % SHEETS[2]),
            ("还缺什么", "【%s】列出未通过的校验、缺的文件，以及报税时要注意的事。" % SHEETS[9])]),
        ("二、口径（与美客多报表的不同之处）", [
            ("按交易日归集", "每一笔收入和费用按 Falabella“Transacciones”报表的交易日"
             "（Fecha de transacción，商品价款≈送达日）计入当月，和平台的结算一致。"
             "按下单日的 GMV 作为参考指标列在②下方，两者的差在③桥 A 逐单解释。"),
            ("含 IVA", "全部金额含 IVA（Monto con IVA），按会计师的要求，与 Falabella 后台和到账金额一致，"
             "也与美客多智利报表的口径相同。经营贡献毛利 ＝ 本月结算净额；其中应缴 SII 的 IVA 见②备查项。"),
            ("费用科目", "Falabella 的 147 种交易类型全部登记在字典里（西语原名＋中文＋归类），"
             "见【%s】。出现字典里没有的类型时报 error，金额进“未登记科目 ⚠”行。" % SHEETS[3])]),
        ("三、数据来源", sources),
        ("四、本期要点", points),
        ("五、各页说明", [
            (SHEETS[1], "各店分列 + 合计。行位置每期固定，可直接跨月对比；灰底=公式，浅蓝=小计，黄底=需你录入。"),
            (SHEETS[2], "桥 A 订单 → 结算、桥 B 费用发票 → 交易明细、桥 C 结算净额 → 到账、桥 D 平台代开销售单据。"),
            (SHEETS[3], "本月出现的每一种交易类型：西语原名、中文、归入哪一行、笔数、金额（含 IVA）。"),
            (SHEETS[4], "取消与退货：按状态、按原因、按配送方式，以及退货订单上的全部资金流。"),
            (SHEETS[5], "不适用：Falabella 不提供逐件退货质检报表。"),
            (SHEETS[6], "按 SKU 的收入、费用与贡献；黄色列填单位采购成本即得 SKU 毛利。"),
            (SHEETS[7], "本月按周的销售与费用；有往月下载时追加各月对比。"),
            (SHEETS[8], "Fulfillment 库存快照：库存、近 4 周销量、可售周数、滞销与断货风险。"),
            (SHEETS[9], "校验结果、缺失文件及原因、报税提示。校验和缺失为空表示没有待办。")]),
        ("六、利润里含 IVA", [
            ("贡献毛利 ＝ 结算净额", "经营贡献毛利（本月 %s）就是 Falabella 本月结算给你的钱（含 IVA）。"
             "其中 %s 是应缴 SII 的 IVA 净额（销项 − 进项），不是你的利润；报税以 SII 的单据为准。"
             % (fmt(tot("tx_gross")), fmt(tot("tx_iva"))))]),
        ("七、会计处理", policies),
        ("八、报税时要注意", [("提示 %d" % i, q) for i, q in enumerate(TAX_NOTES, 1)]),
    ]
    bad = RF.failed_errors(ctx["checks"])
    warning = ("本期有 %d 项 error 级校验未通过，数字可能不完整，请先看【%s】顶部的校验结果。"
               % (len(bad), SHEETS[9])) if bad else None
    fam.readme_sheet(wb, SHEETS[0], "Falabella 月度财务报表 · %s · %s" % (ctx["period_cn"], ctx["who"]),
                     "生成于 %s。所有金额单位：智利比索（CLP），无小数。" % ctx["built"],
                     sections, warning=warning)


# ----------------------------------------------------------------- ② P&L

def _fee_share(keys):
    return lambda S: -sum(S["m"]["pl"][k] for k in keys)


def sheet_pl(wb, fam, ctx):
    sh, r, rows_at, note_col = fam.pl_sheet(
        wb, SHEETS[1], PL_ROWS, value=lambda S, key: S["m"]["pl"][key],
        title="合并损益表（%s）" % ctx["period_cn"],
        subtitle="口径：按【交易日】归集（Falabella 结算口径），金额含 IVA（Monto con IVA）。单位：CLP。"
                 "灰底=公式行，浅蓝=小计，黄底=需你录入。行位置每期固定，可直接跨月对比。")
    fill_cogs(fam, sh, rows_at, note_col, ctx)
    sales = lambda S: S["m"]["pl"][FC.SALES]
    r = fam.ratio_block(sh, r + 1, "费用率（占商品销售收入）", [
        ("佣金", _fee_share((FC.COMMISSION,)), sales),
        ("物流（配送共担＋免运促销＋逆向物流）", _fee_share(LOGISTICS_ROWS), sales),
        ("Fulfillment 仓储与操作", _fee_share((FC.FULFILMENT,)), sales),
        ("广告推广", _fee_share((FC.ADS,)), sales),
        ("平台费用合计", _fee_share(FEE_ROWS), sales),
        ("经营贡献毛利率", lambda S: S["m"]["contrib"], sales)], note_col)
    fam.memo_block(sh, r + 1, "备查项（不参与上表计算）", [row for _k, row in PL_MEMO], note_col)


# ② 备查项：(key, (行名, 原文, 取值, 说明))。key 为 None 的行不是金额（结果文件不收）。
# ② 页和结果文件（results_figures）都读这一份。
PL_MEMO = [
    ("gmv_orders", ("订单口径 GMV（含 IVA，不含已取消）", "Lista de Órdenes · Precio pagado",
                    lambda S: S["m"].get("gmv_orders", 0.0),
                    "本月创建的订单买家实付金额，含 IVA；与结算口径的差见③桥 A")),
    (None, ("订单数（不含已取消）", "", lambda S: S["m"].get("orders_live_n", 0), "")),
    ("cancel_amt", ("已取消订单金额", "Canceladas", lambda S: S["m"].get("cancel_amt", 0.0), "")),
    ("ship_buyer_paid", ("买家支付的运费（代收）", "Pago de envío comprador", lambda S: S["m"]["pl"][FC.SHIP_INCOME],
                         "其中 Falabella 配送的部分原样转付平台（下一行），剩余计入“运费收入（净额）”")),
    ("ship_passed_on", ("其中转付 Falabella", "Reversa de pago de envío comprador",
                        lambda S: S["m"]["pl"][FC.SHIP_BUYER], "代收代付，不是收入也不是费用")),
    ("iva_net", ("其中含 IVA（应缴 SII 的 IVA 净额，参考）", "IVA débito − crédito fiscal",
                 lambda S: S["m"]["tx_iva"],
                 "上表金额含 IVA：经营贡献毛利里有这么多是销项 IVA − 进项 IVA（按交易日），要向 SII 缴纳，不是利润")),
]


def fill_cogs(fam, sh, rows_at, note_col, ctx):
    """With a SKU cost table (core/sku_costs): the yellow 商品采购成本 cells get the
    covered SKUs' cost — still yellow, still overridable; the note says the coverage."""
    if ctx.get("sku_costs") is None:
        return
    r = rows_at["cogs"]
    for i, S in enumerate(fam.stores):
        sh.ws.cell(row=r, column=sh.c0 + i).value = S["cost"].cogs
    sh.ws.cell(row=r, column=note_col).value = fam.per_store(lambda S: SC.pl_note(S["cost"])) + COST_IVA_NOTE


# ----------------------------------------------------------------- ③ bridges

def sheet_bridge(wb, fam, ctx):
    stores = fam.stores
    Ln = RF.Line
    A = lambda S, k: (S["m"].get("bridge_a") or {}).get(k, 0.0)
    bridge_a = [
        Ln("本月创建的订单，买家实付（不含已取消）", lambda S: A(S, "orders"), "Lista de Órdenes · Precio pagado"),
        Ln("减：其中本月还没结算的", lambda S: -A(S, "not_settled"),
           "多为月底下单、次月送达的订单（状态：%s）" % fam.per_store(lambda S: "、".join(
               "%s %d" % (k, v) for k, v in ((S["m"].get("bridge_a") or {}).get("not_settled_states")
                                             or {}).items()))),
        Ln("加：以前月份的订单，本月结算", lambda S: A(S, "earlier"),
           "按订单创建月：%s" % fam.per_store(lambda S: "、".join(
               "%s %d 件" % (k, v) for k, v in ((S["m"].get("bridge_a") or {}).get("earlier_months")
                                               or {}).items()))),
        Ln("加：结算价与订单价的差", lambda S: A(S, "price_diff"), "逐件比较，正常为 0"),
        Ln("＝ 推算的本月结算商品价款", lambda S: A(S, "orders") - A(S, "not_settled") + A(S, "earlier")
           + A(S, "price_diff"), style="subtotal"),
        Ln("交易明细：本月“Pago por precio del producto”（含 IVA）", lambda S: A(S, "sales")),
        Ln("残差（应为 0）", lambda S: A(S, "resid"), style="residual"),
    ]
    bridge_b = []
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
            bridge_b += [
                Ln("账期 %s 至 %s：发票明细中交易日在本月的行" % (a_, b_),
                   lambda S, B=B: B(S, "inv_gross"), "InvoiceReport 文件"),
                Ln("　　交易明细中同期、标了发票/贷记单号的行", lambda S, B=B: B(S, "tx_gross")),
                Ln("　　残差（应为 0）", lambda S, B=B: B(S, "inv_gross") - B(S, "tx_gross"), style="residual"),
                Ln("　　参考：本账期发票总额（含 IVA）", lambda S, B=B: B(S, "factura"), "与卖家后台账期列表一致"),
                Ln("　　参考：本账期贷记单总额（含 IVA）", lambda S, B=B: B(S, "nc")),
            ]
        else:
            bridge_b.append(Ln("账期 %s 至 %s：发票还没出 ⏳" % (a_, b_), lambda S, P=P: P(S, "tx_gross"),
                               "本月这几天的费用已按交易明细计入②，数字不受影响；发票 %s 起开出，下个月的下载会取到它"
                               % (pd.Timestamp(b_) + pd.Timedelta(days=1)).date(), style="pending"))
    bridge_b.append(Ln("交易明细中不开票的费用", lambda S: S["m"]["not_invoiced"],
                       "没有发票号的费用行（一般为 0）"))
    m = lambda S, k: S["m"][k]
    bridge_c = [
        Ln("经营贡献毛利（见②）", lambda S: m(S, "contrib")),
        Ln("交易明细合计（Σ Monto con IVA）", lambda S: m(S, "tx_gross")),
        Ln("残差（应为 0）", lambda S: m(S, "contrib") - m(S, "tx_gross"), style="residual"),
        Ln("其中：已付款（Pagado）", lambda S: (S["m"]["paid"].get("Pagado") or (0.0, 0))[0],
           "已进入 Fpay 结算单（Estado de cuenta）；本月涉及 %s 张结算单"
           % fam.per_store(lambda S: str(S["m"]["statements"]))),
        Ln("其中：未付款（No Pagado）", lambda S: (S["m"]["paid"].get("No Pagado") or (0.0, 0))[0],
           "下个结算日才到账"),
    ]
    L = lambda S, k: (S["m"].get("liq") or {}).get(k, 0.0)
    bridge_d = [
        Ln("Boletas（含 IVA）", lambda S: L(S, "boleta")),
        Ln("Facturas（含 IVA）", lambda S: L(S, "factura")),
        Ln("Notas de crédito（含 IVA）", lambda S: L(S, "nc")),
        Ln("＝ Liquidación 合计", lambda S: L(S, "total"), style="subtotal"),
        Ln("参考：本月结算的 FBF 订单商品价款（含 IVA）", lambda S: S["m"]["fbf_sales"],
           "单据在出库/送达时开，结算在送达后，两者有时间差；比较的是量级"),
        Ln("参考：本月结算的卖家自配送（FBS）商品价款（含 IVA）", lambda S: S["m"]["fbs_sales"],
           "这部分的销售单据不在 Liquidación 里，须由卖家自己开具（见①报税提示 1）"),
    ]
    fam.bridge_sheet(wb, SHEETS[2], "平账勾稽表（%s）" % ctx["period_cn"],
                     "每座桥最后一行是残差，应为 0。单位：CLP，金额含 IVA。", [
                         ("桥 A　订单（下单日）→ 本月结算的商品价款（交易日）", bridge_a),
                         ("桥 B　平台费用发票（Factura falabella.com）→ 交易明细里开了票的行", bridge_b),
                         ("桥 C　经营贡献毛利 ＝ 结算净额 → 到账", bridge_c),
                         ("桥 D　Falabella 以卖家名义开具的销售单据（Liquidación factura）", bridge_d)])


# ----------------------------------------------------------------- ④ fees

TEXT10 = F(10)
FEE_COLS = [RF.Col("店铺", 22, font=TEXT10), RF.Col("交易类型（Falabella 原文）", 58, font=TEXT10),
            RF.Col("中文", 36, font=TEXT10), RF.Col("损益表行", 22, font=TEXT10),
            RF.Col("归类", 18, font=TEXT10), RF.Col("笔数", 8, font=TEXT10),
            RF.Col("金额（含 IVA）", 16, RF.MONEY)]


def fee_groups(S):
    """④'s rows of one store: transaction type × 中文 × ② line, n and amount (含 IVA),
    in ②'s row order then by amount; None when the month has no transactions.
    ④ and the results file (results_figures) both read this."""
    t = S["tx"]
    if not len(t):
        return None
    keys = [x.key for x in PL_ROWS]
    g = t.groupby([T_COLS["type"], "_zh", "_line"]).agg(n=(AMT, "size"), amt=(AMT, "sum")).reset_index()
    g["_order"] = g["_line"].map(lambda k: keys.index(k) if k in keys else 99)
    return g.sort_values(["_order", "amt"])


def sheet_fees(wb, fam, ctx):
    ws = fam.table_sheet(wb, SHEETS[3], "平台费用明细（%s）" % ctx["period_cn"],
                         "本月交易明细里出现的每一种交易类型。归入②的哪一行见“损益表行”。金额 CLP，含 IVA。", FEE_COLS)
    r = 5
    for S in fam.stores:
        g = fee_groups(S)
        if g is None:
            continue
        for _, x in g.iterrows():
            cells = fam.write_row(ws, r, FEE_COLS, [
                S["store"], x[T_COLS["type"]], x["_zh"], pl_label(x["_line"]),
                FC.LINE_CATEGORY.get(x["_line"], "⚠ 未登记"), int(x["n"]),
                float(x["amt"])])
            if x["_line"] == "unregistered":
                for c in cells[:6]:
                    c.font = F(10, False, K.C_RED)
            r += 1
        ws.cell(row=r, column=1, value="%s 合计" % S["store"]).font = F(10, True)
        c = fam.money_cell(ws.cell(row=r, column=7), S["m"]["tx_gross"])
        c.font = F(10, True)
        c.fill = fill(K.C_TOTAL)
        ws.cell(row=r, column=3, value="＝ ② 收入各行 ＋ 费用各行").font = F(9, False, K.C_GREY)
        r += 2
    ws.freeze_panes = "A5"


# ----------------------------------------------------------------- ⑤ returns

def sheet_returns(wb, fam, ctx):
    ws = wb.create_sheet(SHEETS[4])
    RF.sheet_title(ws, "退货退款分析（%s）" % ctx["period_cn"],
                   "上半：本月创建的订单里被取消/退货的（订单列表）；下半：本月交易明细里有退货冲回的订单，"
                   "它们的全部资金流。金额含 IVA。")
    RF.set_widths(ws, [22, 30, 60, 12, 16, 50])
    money = fam.money_cell
    r = 4
    for S in fam.stores:
        R = S["returns"]
        r = RF.section(ws, r, S["store"], 6)
        if "by_status" in R:
            live = S["m"].get("items_n", 0) + S["m"].get("cancel_n", 0)
            for st_, row in R["by_status"].iterrows():
                ws.cell(row=r, column=1, value="订单状态").font = F(10, True)
                ws.cell(row=r, column=2, value=st_)
                ws.cell(row=r, column=3, value="占本月订单件数 %s" % pct(row["count"] / live if live else 0))
                ws.cell(row=r, column=4, value=int(row["count"])).border = BOX
                money(ws.cell(row=r, column=5), float(row["sum"]))
                r += 1
            r += 1
            for (st_, reason), row in R["by_reason"].head(25).iterrows():
                ws.cell(row=r, column=1, value="原因")
                ws.cell(row=r, column=2, value=st_)
                ws.cell(row=r, column=3, value=reason).alignment = wrap()
                ws.cell(row=r, column=4, value=int(row["count"])).border = BOX
                money(ws.cell(row=r, column=5), float(row["sum"]))
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
                ws.cell(row=r, column=2, value=pl_label(line_))
                ws.cell(row=r, column=3, value=zh)
                ws.cell(row=r, column=4, value=int(row["count"])).border = BOX
                money(ws.cell(row=r, column=5), float(row["sum"]))
                r += 1
            ws.cell(row=r, column=3, value="退货订单本月净影响").font = F(10, True)
            c = money(ws.cell(row=r, column=5), float(R["money"]["sum"].sum()))
            c.font = F(10, True)
            c.fill = fill(K.C_TOTAL)
            ws.cell(row=r, column=6, value="正数说明这些订单的原始货款也在本月结算；负数是退货净损失").font = F(9, False, K.C_GREY)
            r += 1
        r += 2


# ----------------------------------------------------------------- ⑦ SKU

_SKU_FORMULA = dict(fmt=RF.MONEY, border=False, fill=K.C_CALC)
SKU_COLS = ([RF.Col("店铺", 18), RF.Col("SKU", 18), RF.Col("商品", 44), RF.Col("结算件数", 9),
             RF.Col("退货件数", 9)]
            + [RF.Col(h, w, RF.MONEY) for h, w in (
                ("商品销售收入", 14), ("运费收入（净额）", 12), ("退货退款", 12), ("平台促销", 12),
                ("佣金", 12), ("物流（净）", 12), ("Fulfillment", 12), ("其他费用", 12),
                ("贡献（未扣成本）", 14))]
            + [RF.Col("单位成本（录入，含 IVA）", 12, RF.MONEY, fill=K.C_INPUT),
               RF.Col("商品成本", 14, formula="=-O{r}*D{r}", **_SKU_FORMULA),
               RF.Col("SKU 毛利", 14, formula="=N{r}+P{r}", **_SKU_FORMULA)])
# With a SKU cost table (core/sku_costs): the cost is on units net of returns, the unit
# cost comes from the table (yellow only where it is missing), and a column says where
# each cost comes from.
SKU_COLS_COST = SKU_COLS[:14] + [
    RF.Col("单位成本（含 IVA）", 12, RF.MONEY),
    RF.Col("商品成本", 14, formula="=-O{r}*(D{r}-E{r})", **_SKU_FORMULA),
    RF.Col("SKU 毛利", 14, formula="=N{r}+P{r}", **_SKU_FORMULA),
    RF.Col("成本来源", 40, font=F(9, False, K.C_GREY))]


def sku_items(S):
    """⑦'s SKUs for the cost table (core/sku_costs.Item): units, returned units, sales."""
    g = S["sku"]
    return [SC.Item(sku, x["name"] if isinstance(x["name"], str) else "", int(x["units"]),
                    int(x["returns"]), float(x[FC.SALES]))
            for sku, x in g.iterrows() if sku != NO_SKU] if len(g) else []


def sku_list(store, folder, month):
    """(platform, country, currency, [Item]) of one store's month — for
    tools/sku_cost_template. Reads the download folder only."""
    return PLATFORM, COUNTRY, CUR, sku_items(derive(load_store(store, folder, month)))


def _cost_cells(cells, S, sku):
    """⑦ with a cost table: unit cost, its source, yellow where it is missing."""
    if sku == NO_SKU:
        cells[17].value = "店铺级费用，不计商品成本"
        return
    ln = S["cost"].line(sku)
    if ln is not None and ln.status == SC.OK:
        cells[14].value = ln.unit_cost
        cells[17].value = SC.status_text(ln)
    else:
        cells[14].fill = fill(K.C_INPUT)
        cells[17].value = SC.status_text(ln)
        cells[17].font = F(9, True, K.C_AMBER)


def sheet_sku(wb, fam, ctx):
    costs = ctx.get("sku_costs") is not None
    cols = SKU_COLS_COST if costs else SKU_COLS
    ws = fam.table_sheet(wb, SHEETS[6], "SKU 盈利分析（%s）" % ctx["period_cn"],
                         "按交易明细的 SKU 汇总（含 IVA）。广告、仓储等没有 SKU 的店铺级费用单列一行。" + (
                             "单位成本 ＝ 商品成本表的成本（不含 IVA）× 1.19，商品成本 ＝ 单位成本 ×（结算件数 − 退货件数）；"
                             "“未录入成本”的 SKU 黄色格可临时填写（含 IVA），长期请补录到成本表。" if costs else
                             "黄色列填单位采购成本（含 IVA），右边两列自动算出 SKU 毛利。"), cols)
    r = 5
    ship = list(LOGISTICS_ROWS)
    other = [FC.ADS, FC.SUBSCRIPTION, FC.PENALTY, FC.OTHER, "unregistered", FC.OTHER_INCOME]
    for S in fam.stores:
        g = S["sku"]
        for sku, x in g.iterrows():
            money = [x[FC.SALES], x[FC.SHIP_INCOME] + x[FC.SHIP_BUYER], x[FC.REFUND], x[FC.PROMO],
                     x[FC.COMMISSION],
                     sum(x[k] for k in ship), x[FC.FULFILMENT], sum(x[k] for k in other), x["contrib"]]
            cells = fam.write_row(ws, r, cols, [S["store"], sku, x["name"], int(x["units"]), int(x["returns"])]
                                  + [float(v) for v in money] + ([None, None, None, None] if costs else
                                                                 [0, None, None]))
            if costs:
                _cost_cells(cells, S, sku)
            r += 1
        r += 1
    ws.freeze_panes = "D5"
    ws.auto_filter.ref = "A4:%s%d" % (get_column_letter(len(cols)), max(4, r - 1))


# ----------------------------------------------------------------- ⑧ trend

TREND_COLS = [RF.Col("店铺", 22), RF.Col("周（起始日）", 16), RF.Col("结算件数", 10),
              RF.Col("商品销售收入", 16, RF.MONEY), RF.Col("退货退款", 14, RF.MONEY),
              RF.Col("平台费用", 16, RF.MONEY), RF.Col("贡献（未扣成本）", 16, RF.MONEY),
              RF.Col("费用率", 10, K.PCT, border=False)]
# the earlier-months block below the weeks: same columns, no borders on the text
HIST_COLS = [RF.Col("店铺", 22, border=False), RF.Col("月份", 16, border=False),
             RF.Col("结算件数", 10, border=False)] + TREND_COLS[3:]


def sheet_trend(wb, fam, ctx):
    ws = fam.table_sheet(wb, SHEETS[7], "月度趋势（%s）" % ctx["period_cn"],
                         "本月按周（周一开始，按交易日，含 IVA）。往月趋势需要往月的下载：同一台机器上有更早月份的"
                         "Falabella 下载目录时，下方自动追加各月合计。",
                         TREND_COLS)
    r = 5
    for S in fam.stores:
        t = S["tx"]
        if not len(t):
            continue
        wk = t["_date"] - pd.to_timedelta(t["_date"].dt.weekday, unit="D")
        t = t.assign(_wk=wk.dt.date)
        for w, x in t.groupby("_wk"):
            sales = x.loc[x["_line"] == FC.SALES, AMT].sum()
            fees = x.loc[x["_line"].isin(FEE_ROWS), AMT].sum()
            fam.write_row(ws, r, TREND_COLS, [
                S["store"], str(w), int((x["_line"] == FC.SALES).sum()), float(sales),
                float(x.loc[x["_line"] == FC.REFUND, AMT].sum()), float(fees),
                float(x[AMT].sum()), (-fees / sales) if sales else 0.0])
            r += 1
        r += 1
    hist = ctx.get("history") or {}
    if hist:
        r = RF.section(ws, r + 1, "往月（来自本机更早的 Falabella 下载）", len(TREND_COLS))
        for store, months in hist.items():
            for mon, v in sorted(months.items()):
                fam.write_row(ws, r, HIST_COLS, [
                    store, mon, v["items"], v["sales"], v["refund"], v["fees"], v["contrib"],
                    (-v["fees"] / v["sales"]) if v["sales"] else 0.0])
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
                             "sales": float(t.loc[ln == FC.SALES, AMT].sum()),
                             "refund": float(t.loc[ln == FC.REFUND, AMT].sum()),
                             "fees": float(t.loc[ln.isin(FEE_ROWS), AMT].sum()),
                             "contrib": float(t[AMT].sum())}
            except Exception:
                continue
        if seen:
            out[S["store"]] = dict(sorted(seen.items())[-6:])
    return out


# ----------------------------------------------------------------- ⑨ inventory

INV_COLS = [RF.Col("店铺", 20), RF.Col("SKU 卖家", 18), RF.Col("SKU Falabella", 14), RF.Col("商品", 50),
            RF.Col("售价（含 IVA）", 14, RF.MONEY), RF.Col("可售库存", 10), RF.Col("近 4 周销量", 10),
            RF.Col("可售周数", 10), RF.Col("本月结算件数", 12), RF.Col("库存货值（售价）", 16, RF.MONEY),
            RF.Col("判断", 16)]


def sheet_inventory(wb, fam, ctx):
    ws = fam.table_sheet(wb, SHEETS[8], "库存与动销分析（Fulfillment by Falabella）",
                         "库存快照取自下载当时（不是月末），近 4 周销量是平台给的。可售周数 = 库存 ÷ (近 4 周销量 ÷ 4)。"
                         "滞销：有库存、近 4 周没卖；断货风险：近 4 周有销量、可售不到 2 周。", INV_COLS)
    r = 5
    for S in fam.stores:
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
            cells = fam.write_row(ws, r, INV_COLS, [
                S["store"], x.get("SKU Vendedor"), x.get("SKU GSC"), x.get("Nombre"),
                float(x.get("Precio") or 0), int(stock), int(w4),
                round(weeks, 1) if weeks is not None else None,
                int(sold.get(x.get("SKU Vendedor"), 0)) if len(sold) else 0,
                stock * float(x.get("Precio") or 0), verdict])
            if verdict in ("滞销", "断货风险"):
                cells[10].font = F(10, True, K.C_AMBER)
            r += 1
        r += 1
    ws.freeze_panes = "E5"
    ws.auto_filter.ref = "A4:K%d" % max(4, r - 1)


# ----------------------------------------------------------------- ⑩ gaps

def store_meta_missing(S):
    return RF.meta_missing(S["meta"])


def sheet_gaps(wb, fam, ctx):
    fam.gaps_sheet(wb, SHEETS[9], "数据缺口与待办（%s）" % ctx["period_cn"], ctx["checks"],
                   notes=[("报税时要注意（不影响本报表的数字）",
                           [("提示 %d" % i, q) for i, q in enumerate(TAX_NOTES, 1)])])


# ======================================================================= entry

def build_context(stores, month, build_date=None):
    y, mo = periods.parse_month(month)
    return {"month": month, "period_cn": "%d 年 %d 月" % (y, mo),
            "who": stores[0]["store"] if len(stores) == 1 else "%d 家店铺汇总" % len(stores),
            "built": build_date or datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}


def apply_costs(stores, month, checks, table=None):
    """The SKU cost table (core/sku_costs), when config.json names one: each store
    gets S["cost"] and S["m"]["cogs"]; its checks are appended. Returns the table."""
    table = table or SC.load_configured()
    if table is None:
        return None
    for S in stores:
        S["cost"] = SC.apply(table, PLATFORM, COUNTRY, CUR, month, sku_items(S), COST_FACTOR)
        S["m"]["cogs"] = S["cost"].cogs
        S["m"]["profit"] = S["m"]["contrib"] - S["cost"].cogs
    checks.extend(SC.checks(table, [(S["store"], S["cost"]) for S in stores], PLATFORM, COUNTRY))
    return table


# ======================================================================= results file

# ② row of a transaction line that is not a ② row itself
PL_ROW_OF_LINE = {FC.SHIP_INCOME: "ship_net", FC.SHIP_BUYER: "ship_net"}


def results_figures(S):
    """This store's numbers for the monthly results file (core/results), from the
    same S and row specs that fill the workbook: ② (PL_ROWS, S["m"]["pl"], the
    cost table's cogs as fill_cogs writes it), PL_MEMO, ④ by transaction type
    (fee lines only), ⑤'s money of the returned orders by line, ⑦ S["sku"], ⑨
    S["stock"] (Falabella gives sales of the last 4 weeks: sold_window_days 28)."""
    from core import results
    cost = S.get("cost")
    spec = [{"key": x.key, "label": x.label, "original": x.original, "kind": x.kind, "formula": x.formula,
             "value": (S["m"]["pl"][x.key] if x.kind == "d" else
                       (cost.cogs if cost is not None else 0.0) if x.kind == "in" else None)}
            for x in PL_ROWS]
    memo = [{"key": k, "label": label, "value": float(fn(S))} for k, (label, _o, fn, _note) in PL_MEMO if k]
    fees = []
    g = fee_groups(S)
    if g is not None:
        for _, x in g.iterrows():
            if x["_line"] in FEE_LINES:
                fees.append({"name": x[T_COLS["type"]], "label": x["_zh"],
                             "pl_key": PL_ROW_OF_LINE.get(x["_line"], x["_line"]), "amount": float(x["amt"])})
    rets = []
    money = S["returns"].get("money")
    if money is not None and len(money):
        for (line_, zh), row in money.iterrows():
            rets.append({"key": line_, "label": zh, "count": int(row["count"]), "amount": float(row["sum"])})
    skus = []
    for sku, x in (S["sku"].iterrows() if len(S["sku"]) else ()):
        skus.append({"sku": sku, "title": x["name"] if isinstance(x["name"], str) else "",
                     "units": float(x["units"]), "units_returned": float(x["returns"]),
                     "net_sales": float(x[FC.SALES] + x[FC.REFUND]), "fees": float(x["fees"]),
                     "contribution": float(x["contrib"]),
                     "cost": results.sku_cost(cost, sku) if sku != NO_SKU else None})
    inv = None
    s = S.get("stock")
    if s is not None:
        rows = [{"sku": x.get("SKU Vendedor"), "title": x.get("Nombre") if isinstance(x.get("Nombre"), str) else "",
                 "stock": float(x.get("Stock Disponible") or 0), "sold_30d": float(x.get("Ventas ultimas 4 semanas")
                                                                                  or 0)}
                for _, x in s.iterrows()]
        inv = {"as_of": stock_date(S), "sold_window_days": 28, "rows": rows}
    return {"pl": results.pl_rows(spec), "memo": memo, "fees": fees, "returns": rets, "sku": skus,
            "inventory": inv}


def stock_date(S):
    """The snapshot's day: the Product_details file's modification day (the
    download), else run_meta.json's downloaded_at; None when neither is known."""
    p = S["files"].get("stock")
    if p and os.path.isfile(p):
        return datetime.date.fromtimestamp(os.path.getmtime(p)).isoformat()
    at = (S.get("meta") or {}).get("downloaded_at")
    return str(at)[:10] if at else None


def build(pairs, month, out_path, strict=False, build_date=None):
    """pairs: [(store, folder)]. Returns a result dict (outcome, checks, out, and
    "figures": {store: results_figures(S)} when the workbook was written)."""
    stores = [derive(load_store(s, f, month)) for s, f in pairs]
    ctx = build_context(stores, month, build_date)
    ctx["checks"] = validate(stores)
    ctx["sku_costs"] = apply_costs(stores, month, ctx["checks"])
    ctx["history"] = history(stores, None)
    bad = [c for c in ctx["checks"] if not c["ok"] and not c["skipped"] and c["severity"] == "error"]
    if strict and bad:
        return {"outcome": "strict_aborted", "checks": ctx["checks"], "out": out_path}
    build_workbook(stores, ctx, out_path)
    return {"outcome": "checks_failed" if bad else "ok", "checks": ctx["checks"], "out": out_path,
            "metrics": {S["store"]: S["m"] for S in stores},
            "results": RF.results_inputs(stores, results_figures)}
