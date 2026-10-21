# -*- coding: utf-8 -*-
"""MercadoLibre report — fixed facts about the platform's files, and small helpers.

Raw-file structure (header rows, column names), the fee catalogue FEE_CATALOG
(the only place a Detalle is mapped to a Chinese name and bucket), order status
and return dictionaries, the nominal tax rates, and the number / date helpers
(money, fmt, pct, parse_spanish_date, month_bounds …) used everywhere.

Per-country values (currency-suffixed column names, tolerances …) are not here:
they are in site.py."""

import re

import pandas as pd

from core import report_kit

from platforms.mercadolibre.report import site


# ════════════════════════════════════════════════════════════════════════
# 常量：原始报表的结构
# ════════════════════════════════════════════════════════════════════════

VENTAS_HEADER_ROW = 5        # 0-indexed，实际表头在第 6 行
BILLING_HEADER_ROW = 7       # Facturación / Notas de Crédito，表头在第 8 行
CARGOS_HEADER_ROW = 5        # Reporte_Cargos_Full_*，表头在第 6 行
PAGOS_HEADER_ROW = 9         # Reporte_Pagos_Facturas_*，表头在第 10 行
STORAGE_HEADER_ROW = 4       # Costos_por_servicio_almacenamiento，表头在第 5 行
RETURNS_HEADER_ROW = 2       # Returns_*.xlsx / Triages，表头在第 3 行

SPANISH_MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}
MONTH_FULL_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo",
                 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre",
                 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
# 平台文件名里用的月份缩写。注意 9 月是 "Sept" 不是 "Sep"，10 月两种都见过。
FILE_MONTH_TOKENS = {
    "ene": 1, "enero": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
    # 秘鲁把九月叫 Setiembre。IDG_PE 2026-09 的账单文件名实测仍是 "Sept2026"，
    # 但平台若哪天按本地拼法给出 "Set2026"，也要认得。
    "set": 9,
}

# Ventas 列名。各店的导出列数不一致（例如 EWTTO 多一列"换货运费"），所以一律按列名取，
# 绝不按列号。缺列时按 0 处理。
V_ORDER_ID = "# de venta"
V_PACK_ID = "Orden de compra"
V_DATE = "Fecha de venta"
V_STATUS = "Estado"
V_UNITS = "Unidades"
V_SKU = "SKU"
V_TITLE = "Título de la publicación"
V_PACK_FLAG = "Paquete de varios productos"
V_KIT = "Pertenece a un kit"
V_AD_SALE = "Venta por publicidad"
V_DISCOUNT = "Descuentos y bonificaciones"   # 平台导出为空列，本报表由 Total 倒推

# Facturación 列名
F_INVOICE = "N° de factura fiscal"
F_CHARGE_DATE = "Fecha del cargo"
F_CHARGE_NO = "Número del cargo"
F_DETALLE = "Detalle"
F_CHANNEL = "Descontado de la operación"      # Si / No / No aplica
F_STATE = "Estado del cargo"
F_REVERSES = "Cargo que bonifica"             # 指向被冲销的 Número del cargo
# 智利的账单把同一列叫 "Cargo que anula"（DAB_CL 2026-09 实测），含义相同。
F_REVERSES_ALT = ("Cargo que anula",)
F_AMOUNT = "Valor del cargo"
F_SALE_NO = "Número de venta"
F_PACK_NO = "Número de paquete"
F_SALE_DATE = "Fecha de venta"

CH_SI = "Si"
CH_NO = "No"
CH_NA = "No aplica"
CHANNEL_ZH = {CH_SI: "订单内直扣", CH_NO: "月底账单支付", CH_NA: "取消订单冲销"}
CHANNEL_MEANING = {
    CH_SI: "下单打款时平台已直接扣除，不会再收第二次",
    CH_NO: "月底汇总，从 MercadoPago 余额自动扣款",
    CH_NA: "取消/退货订单的原计费与冲销",
}


# ════════════════════════════════════════════════════════════════════════
# 费用科目字典 —— 唯一的真相来源
# ════════════════════════════════════════════════════════════════════════
# (西班牙语 Detalle 原文, 中文, 大类 bucket, 是否订单级)
#
# "是否订单级" 只用于校验：订单级科目理应带 Fecha de venta，账单级理应不带。
# 实际归属仍以数据里有没有 Fecha de venta 为准（数据 > 字典），不一致时记为异常。
#
# 新增科目请加在这里，并同时在 PL_LINES 里给它一行，否则它会落进"未登记科目"。
FEE_CATALOG = [
    # ---- 订单级（打款时直扣，Descontado = Si） ----
    ("Cargo por venta",                                "平台销售佣金",              "commission",   True),
    ("Cargo por venta con afiliados",                  "联盟推广佣金",              "commission",   True),
    ("Cargo por envíos de Mercado Libre",              "平台配送费 (Mercado Envíos)", "shipping",   True),
    ("Cargo por devolución",                           "退货处理费",                "returns",      True),
    # ---- 账单级（月底自动扣款，Descontado = No） ----
    ("Cargo por campaña de publicidad de Product Ads", "商品广告费 Product Ads",     "ads",          False),
    ("Cargo por campaña de publicidad de Display Ads", "展示广告费 Display Ads",     "ads",          False),
    ("Cargo por servicio de almacenamiento Full",      "Full 仓储费",               "storage",      False),
    ("Cargo por servicio de colecta Full",             "Full 揽收/入仓费",           "fulfillment",  False),
    ("Cargo por comprar más espacio en Full",          "Full 购买额外仓位",          "fulfillment",  False),
    ("Cargo por sobrepasar espacio Full",              "Full 超仓位罚金",            "penalty",      False),
    ("Cargo por retiro de stock Full",                 "Full 退仓/取货费",           "fulfillment",  False),
    ("Cargo por stock antiguo en Full",                "Full 长龄库存费 (库龄4个月+)", "fulfillment", False),
    ("Cargo por incumplimiento en Envíos Full",        "Full 发货违规罚金 (库存差异)", "penalty",     False),
    ("Cargo por mantenimiento de Mi página",           "官方店页面维护费",           "subscription", False),
    # 2026-10 登记。注意平台这里的 Publicidad 是大写 P，匹配是逐字的。
    ("Cargo por campaña de Publicidad de Seguidores",  "粉丝推广广告费 (Seguidores)", "ads",         False),
    # 自成一类 "misc"：不是 Full 的费用（放进 penalty 会被算进 Full 合计、把
    # "Full 费用报表 vs 账单"打破），也不是广告。归属待会计确认。
    ("Cargo debitado del dinero en garantía del Programa Beneficio de Reputación",
                                                       "信誉保障计划保证金扣款",      "misc",         False),
]
FEE_BY_ES = {es: (zh, bucket, order_level) for es, zh, bucket, order_level in FEE_CATALOG}

# 冲销行的名字通常是原名把 "Cargo por" 换成 "Anulación del cargo por"，
# 但平台自己有拼写不一致（Product → Products），所以这里显式列出例外。
REVERSAL_ALIASES = {
    "Anulación del cargo por campaña de publicidad de Products Ads":
        "Cargo por campaña de publicidad de Product Ads",
}


def _auto_reversal_name(es):
    return es.replace("Cargo por", "Anulación del cargo por", 1)


# 冲销名 → 原科目名
REVERSAL_TO_PARENT = dict(REVERSAL_ALIASES)
for _es, _zh, _b, _o in FEE_CATALOG:
    REVERSAL_TO_PARENT.setdefault(_auto_reversal_name(_es), _es)

UNCLASSIFIED_ZH = "⚠ 未登记科目（需人工确认）"


def is_reversal(detalle):
    return str(detalle or "").startswith("Anulación del cargo")


def classify(detalle):
    """Detalle → (中文, bucket, 是否订单级, 是否冲销, 是否已登记)。

    未登记的科目返回 bucket='unclassified'。调用方必须把它当成"看得见的异常"，
    而不是丢进 other 了事。
    """
    s = str(detalle or "").strip()
    if is_reversal(s):
        parent = REVERSAL_TO_PARENT.get(s)
        if parent and parent in FEE_BY_ES:
            zh, bucket, order_level = FEE_BY_ES[parent]
            return "－ " + zh + " 冲销", bucket, order_level, True, True
        return UNCLASSIFIED_ZH + "（冲销）", "unclassified", None, True, False
    if s in FEE_BY_ES:
        zh, bucket, order_level = FEE_BY_ES[s]
        return zh, bucket, order_level, False, True
    return UNCLASSIFIED_ZH, "unclassified", None, False, False


# ════════════════════════════════════════════════════════════════════════
# 订单状态 / 退货原因 词典
# ════════════════════════════════════════════════════════════════════════
# (中文, 资金归属, 库存处置)。未收录的状态照常统计，标"待确认"。
STATUS_ZH = {
    "Cancelada por el comprador": ("买家下单后取消", "买家", "未发出，库存无损"),
    "Cancelaste la venta": ("卖家主动取消", "卖家", "未发出，库存无损"),
    "Cancelada por seguridad": ("平台风控取消", "平台", "未发出，库存无损"),
    "Paquete cancelado por Mercado Libre": ("平台取消整个包裹", "平台", "未发出，库存无损"),
    "Devolución finalizada con reembolso al comprador": ("退货完成，已退款买家", "买家", "待质检结果"),
    "Devolución finalizada. Pusimos el producto de nuevo a la venta": ("退货完成，商品已重新上架", "买家", "✔ 可回收，重新销售"),
    "Devolución finalizada. Pusimos los productos de nuevo a la venta": ("退货完成，多件商品已重新上架", "买家", "✔ 可回收，重新销售"),
    "Devolución revisada. Pusimos el producto de nuevo a la venta": ("退货已质检，商品重新上架", "买家", "✔ 可回收，重新销售"),
    "Devolución revisada. Solicita el retiro del producto": ("退货已质检，需卖家自提", "买家", "✘ 不可再售，须付费取回"),
    "Devolución finalizada. Descartamos el producto": ("退货完成，商品已销毁", "买家", "✘ 彻底损失"),
    "Devolución rechazada. Regresamos el producto al comprador": ("退货被驳回，商品退回买家", "卖家(保留货款)", "— 无库存回收"),
    "Devolución en camino": ("退货在途", "待定", "待质检"),
    "Devolución en preparación": ("退货准备中", "待定", "待质检"),
    "Devolución en revisión": ("退货质检中", "待定", "待质检"),
    "Devolución con fecha actualizada": ("退货日期已更新", "待定", "待质检"),
    "Devolución finalizada": ("退货已完成", "买家", "待质检结果"),
    "Mediación finalizada con reembolso al comprador": ("平台仲裁，判退款给买家", "买家", "通常不回收"),
    "Mediación finalizada. Te dimos el dinero.": ("平台仲裁，判货款归卖家", "卖家(保留货款)", "— 无库存回收"),
    "Mediación con devolución habilitada": ("仲裁中，已开放退货", "待定", "待质检"),
    "Cambio no entregado": ("换货未送达", "待定", "待处理"),
    "Venta con solicitud de cambio": ("买家申请换货", "卖家(换货)", "换新，旧品待回收"),
    "Cambio entregado. Devolución finalizada.": ("换货已送达，退货完成", "卖家(换货)", "换新，旧品待回收"),
    "Paquete no entregado": ("包裹未送达", "待定", "待处理"),
    "Reclamo cerrado": ("投诉已关闭", "待定", "—"),
    "Entregado": ("已送达", "—", "—"),
    "Venta entregada": ("销售已送达", "—", "—"),
    "En camino": ("运输中", "—", "—"),
    "Procesando en la bodega": ("仓库处理中", "—", "—"),
    # ↓ 实际数据里出现、平台文案的变体或细分说法
    "Cancelada": ("订单已取消", "待定", "未发出，库存无损"),
    # 秘鲁 IDG_PE 2026-09 见到的两种取消说法
    "Venta cancelada": ("销售已取消", "待定", "未发出，库存无损"),
    "Cancelada. Tu transportista debe devolverte el paquete": ("已取消，承运商需把包裹退回给你", "买家", "包裹退回卖家，待验货"),
    "Devolución en camino sin costo de envío": ("退货在途（免运费）", "待定", "待质检"),
    "Devolución en preparación sin costo de envío": ("退货准备中（免运费）", "待定", "待质检"),
    "Devolución en camino. Revisaremos el producto": ("退货在途，将进行质检", "待定", "待质检"),
    "Devolución en camino. Revisaremos los productos": ("退货在途，多件将进行质检", "待定", "待质检"),
    "Devolución finalizada. Revisa los resultados.": ("退货完成，请查看质检结果", "买家", "待质检结果"),
    "Devolución no entregada": ("退货未送达", "待定", "待处理"),
    "Mediación finalizada. Te dimos el dinero": ("平台仲裁，判货款归卖家", "卖家(保留货款)", "— 无库存回收"),
    "Mediación en espera de respuesta de Mercado Libre": ("仲裁中，等待平台答复", "待定", "待处理"),
    "Mediación para responder el lunes": ("仲裁中，需在周一前答复", "待定", "★ 待卖家答复"),
    "Reclamo abierto para resolver hasta mañana a las 14 hs": ("投诉待处理，明日 14 时截止", "待定", "★ 待卖家答复"),
    "Reclamo en espera de respuesta del comprador": ("投诉中，等待买家答复", "待定", "待处理"),
    "Reclamo cerrado con reembolso al comprador": ("投诉结案，全额退款买家", "买家", "通常不回收"),
    "Reclamo cerrado con reembolso parcial": ("投诉结案，部分退款买家", "买家(部分)", "通常不回收"),
}
# 识别"售后/异常订单"的正则。用于售后率与退货分析。
ABNORMAL_RE = r"Devoluci|Mediaci|Cancel|Reclamo|Cambio|no entregado|rechaz"
# 库存去向分组
STATES_BACK_ON_SALE = [
    "Devolución finalizada. Pusimos el producto de nuevo a la venta",
    "Devolución finalizada. Pusimos los productos de nuevo a la venta",
    "Devolución revisada. Pusimos el producto de nuevo a la venta",
]
STATES_STOCK_LOST = [
    "Devolución revisada. Solicita el retiro del producto",
    "Devolución finalizada. Descartamos el producto",
]
STATES_SELLER_KEEPS = [
    "Mediación finalizada. Te dimos el dinero.",
    "Devolución rechazada. Regresamos el producto al comprador",
]

# 退货质检：商品状态 → (中文, 库存后果)
RETURN_REASON_ZH = {
    "Sello dañado": ("包装封条破损", "✘ 多数判定不可再售"),
    "Producto dañado": ("商品损坏", "✘ 不可再售"),
    "Producto con marcas de mal uso": ("商品有使用/磨损痕迹", "✘ 多数不可再售"),
    "No funciona correctamente": ("功能故障，无法正常工作", "✘ 不可再售"),
    "Producto sin accesorios, manuales o etiquetas": ("缺配件/说明书/标签", "△ 补齐后可能可售"),
    "Producto ausente, destruido o ultrajado": ("商品缺失、损毁或被拆封破坏", "✘ 彻底损失"),
    "-": ("平台未记录具体原因", "△ 需人工核实"),
}
# 退货质检：处置结果 → (中文, 库存价值, 财务影响)
RETURN_RESULT_ZH = {
    "Producto quedó nuevamente para la venta": ("商品重新上架销售", "100% 保留", "退款损失可由再次销售弥补"),
    "Producto para retirar en centro de distribución": ("需卖家到配送中心自提", "需付退仓费才能取回", "★ 主要损失来源：退款 + 取货费 + 商品贬值"),
    "Producto devuelto al comprador": ("退货被驳回，商品退回买家", "商品已给买家", "货款保留，无库存回收"),
    "Producto archivado en centro de distribución": ("商品在配送中心封存", "待处理", "需尽快决定自提或放弃"),
}
MONEY_REFUNDED = "Reembolsamos el dinero al comprador"
MONEY_KEPT = "Te dimos el dinero"

# Returns（Triages）导出现在是英文的：表头和取值都换成了英文，而且不随账号界面
# 语言变（2026-10 全部 11 家店都是英文，其余报表都是西语）。下面所有分析都按
# 西语写的，所以读进来先翻回西语 —— 只翻我们见过、核对过的取值，没见过的原样
# 保留（宁可显示英文，也不猜一个错的西语）。
RETURNS_EN_COLUMNS = {
    "Order": "Número de orden",
    "Review date": "Fecha de revisión",
    "Product status": "Estado del producto",
    "Review result": "Resultado de la revisión",
    "Money status": "Estado del dinero",
}
RETURNS_EN_VALUES = {
    "Estado del producto": {
        "Damaged seal": "Sello dañado",
        "Damaged product": "Producto dañado",
        "Product with signs of misuse": "Producto con marcas de mal uso",
        "It does not work correctly": "No funciona correctamente",
        "Product missing, destroyed or damaged": "Producto ausente, destruido o ultrajado",
    },
    "Resultado de la revisión": {
        "Product is once again available for sale": "Producto quedó nuevamente para la venta",
        "Product to be picked up at distribution center": "Producto para retirar en centro de distribución",
        "Product returned to the buyer": "Producto devuelto al comprador",
        "Product archived in distribution center": "Producto archivado en centro de distribución",
    },
    "Estado del dinero": {
        "We refunded the buyer": MONEY_REFUNDED,
        "We gave you the money": MONEY_KEPT,
    },
}

# 代扣代缴税金的预期费率区间（占含税成交额）。落在区间外说明税制或账号状态变了。
TAX_RATE_MIN, TAX_RATE_MAX = 0.085, 0.095
TAX_RATE_NOMINAL_GROSS = 0.0905     # 占含税价
TAX_RATE_NOMINAL_NET = 0.1050       # 占不含税价（÷1.16）

# SHEET_NAMES：报表家族统一的十页（core/report_kit.py）
SHEET_NAMES = report_kit.SHEET_NAMES


# ════════════════════════════════════════════════════════════════════════
# 通用小工具
# ════════════════════════════════════════════════════════════════════════

def money(x, dec=2):
    try:
        return round(float(x), dec)
    except (TypeError, ValueError):
        return 0.0


def fmt(x, dec=None):
    """1427367.66 -> '1,427,367.66'。不给 dec 时按本站点的币种精度（智利 0 位）。"""
    if dec is None:
        dec = site.MONEY_DEC
    try:
        return "{:,.{d}f}".format(float(x), d=dec)
    except (TypeError, ValueError):
        return "-"


def fmt0(x):
    return fmt(x, 0)


def pct(x, dec=2):
    try:
        return "{:.{d}f}%".format(float(x) * 100.0, d=dec)
    except (TypeError, ValueError):
        return "-"


def parse_spanish_date(value):
    """'31 de agosto de 2026 15:10 hs.' -> Timestamp；解析不了返回 NaT。"""
    if isinstance(value, pd.Timestamp):
        return value
    if not isinstance(value, str):
        return pd.NaT
    m = re.match(r"\s*(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)\s+de\s+(\d{4})"
                 r"(?:\s+(\d{1,2}):(\d{2}))?", value.strip(), re.IGNORECASE)
    if not m:
        return pd.NaT
    day, month_name, year, hh, mm = m.groups()
    month = SPANISH_MONTHS.get(month_name.lower())
    if not month:
        return pd.NaT
    try:
        return pd.Timestamp(int(year), month, int(day), int(hh or 0), int(mm or 0))
    except ValueError:
        return pd.NaT


def num(df, col):
    """按列名取数值列；列不存在时返回全 0（各店导出列数不一致）。"""
    if col in df.columns:
        return pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return pd.Series(0.0, index=df.index)


def idkey(series):
    """订单号在不同报表里一边是 str 一边是 float，统一成无小数点的字符串。"""
    return series.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()


def month_bounds(month_str):
    y, m = int(month_str[:4]), int(month_str[5:7])
    start = pd.Timestamp(y, m, 1)
    end = (start + pd.offsets.MonthEnd(1)).normalize()
    return start, end


def file_month_token(name):
    """'Reporte_Pagos_Facturas_Ago2026.xlsx' -> (2026, 8)；识别不了返回 None。"""
    m = re.search(r"_([A-Za-z]{3,5})(\d{4})\.xlsx$", name)
    if not m:
        return None
    tok = m.group(1).lower()
    mon = FILE_MONTH_TOKENS.get(tok)
    if not mon:
        return None
    return int(m.group(2)), mon


def storage_file_range(name):
    """'01-08-26_31-08-26_Costos...' -> (start, end)；识别不了返回 None。"""
    m = re.match(r"(\d{2})-(\d{2})-(\d{2})_(\d{2})-(\d{2})-(\d{2})_", name)
    if not m:
        return None
    d1, m1, y1, d2, m2, y2 = [int(x) for x in m.groups()]
    try:
        return (pd.Timestamp(2000 + y1, m1, d1), pd.Timestamp(2000 + y2, m2, d2))
    except ValueError:
        return None
