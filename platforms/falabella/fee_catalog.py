# -*- coding: utf-8 -*-
"""
Every Falabella transaction type ("Tipo de transacción" in the Transacciones
report, "Tipo de Transaccion" in the Factura falabella.com report) -> where it
goes in the monthly report.

    TYPE -> (Chinese name, P&L line, family category)

P&L lines are the fixed rows of ② (report_build.PL_ROWS). Family categories are
the fixed list of docs/REPORT_FAMILY.md §5 (commission, shipping, ads,
fulfilment_storage, penalty, subscription, tax_withheld, other) plus the
revenue side (sales, shipping_income, refund, promotion, other_income).

Built from the 147-entry glossary sheet Falabella ships inside every
Transacciones file (2026-10). A type that is not here is an **error** check in
the report — never silently "other": a new fee must be looked at once and given
its row.

Sign convention: amounts are kept as Falabella writes them (income positive,
charges negative). A row's "Reembolso / Corrección / Ajuste" variants go to the
same line as the charge they correct, so the line shows the net.
"""
import re
import unicodedata

# P&L lines (keys of report_build.PL_ROWS)
SALES = "sales"                 # 商品销售收入
SHIP_INCOME = "ship_income"     # 运费收入（买家支付 / 平台付卖家自配送）
REFUND = "refund"               # 退货退款
PROMO = "promo"                 # 平台出资促销与激励
OTHER_INCOME = "other_income"   # 丢失损坏赔付等
COMMISSION = "commission"       # 销售佣金
SHIP_COFIN = "ship_cofin"       # 配送共担费
SHIP_BUYER = "ship_buyer"       # 买家运费转付平台
SHIP_PROMO = "ship_promo"       # 免运费促销
SHIP_REVERSE = "ship_reverse"   # 逆向物流（退货）
FULFILMENT = "fulfilment"       # Fulfillment 仓储与操作
ADS = "ads"                     # 广告推广
SUBSCRIPTION = "subscription"   # 订阅与增值服务
PENALTY = "penalty"             # 罚款与违规扣费
OTHER = "other"                 # 其他调整

# P&L line -> family category
LINE_CATEGORY = {
    SALES: "sales", SHIP_INCOME: "shipping_income", REFUND: "refund", PROMO: "promotion",
    OTHER_INCOME: "other_income", COMMISSION: "commission", SHIP_COFIN: "shipping",
    SHIP_BUYER: "shipping", SHIP_PROMO: "shipping", SHIP_REVERSE: "shipping",
    FULFILMENT: "fulfilment_storage", ADS: "ads", SUBSCRIPTION: "subscription",
    PENALTY: "penalty", OTHER: "other",
}

_C = {}


def _add(line, zh, *names):
    for n in names:
        _C[n] = (zh, line)


# ---------------------------------------------------------------- income side
_add(SALES, "商品价款",
     "Pago por precio del producto", "Corrección de pago por precio del producto")
_add(SALES, "商品价款（CMR 积分支付）",
     "Pago de precio de producto con puntos CMR",
     "Corrección de pago de precio de producto con puntos CMR")
_add(SALES, "门店自助机售出的商品价款", "ITEM_PRICE_KIOSK_CREDIT")
_add(SALES, "纯积分订单的 $10 占位款（平台规则，收付相抵）",
     "Pago por $10 ficticios por venta 100% puntos",
     "Cobro por $10 ficticios por venta 100% puntos",
     "Ajuste del pago por $10 ficticios por venta 100% puntos",
     "Ajuste del cobro por $10 ficticios porventa 100% puntos")
_add(SHIP_INCOME, "买家支付的运费",
     "Pago de envío comprador", "Ajuste Pago de envío comprador",
     "Corrección de pago de envío comprador")
_add(SHIP_INCOME, "平台支付的卖家自配送运费（Falabella directo）",
     "Pago por envío directo", "Corrección de pago envio directo",
     "Cobro por envío directo", "Corrección de cobro por envío directo")
_add(REFUND, "退货冲回商品价款",
     "Descuento por devolución de producto", "Corrección de Cobro sobre precio producto",
     "Reversa del pago de precio de producto con puntos CMR",
     "Corrección de reversa del pago de precio de producto con puntos CMR")
_add(REFUND, "申诉裁定的商品价款调整",
     "Precio del producto por apelación", "Reembolso  precio de producto por apelación",
     "Correccion  precio del producto por apelación",
     "Corrección reembolso precio del producto por apelacion", "Appealed Return Credit")
_add(PROMO, "平台出资的买家优惠券补贴",
     "Pago de aporte promocionales a cliente (Promo)",
     "Descuento por aportes promocionales a clientes (Promo)")
_add(PROMO, "falabella.com 促销活动补贴",
     "Pago de aportes promocionales falabella.com",
     "Corrección de pago aportes promocionales falabella.com")
_add(PROMO, "CMR 卡“Oportunidad Única”活动补贴",
     "Pago por oportunidad única", "Cobro por oportunidad única",
     "Ajuste de pago por oportunidad única", "Ajuste de cobro por oportunidad única")
_add(PROMO, "销售目标达成奖励",
     "Pago por cumplimiento de ventas", "Corrección de Pago por cumplimiento de ventas")
_add(OTHER_INCOME, "物流丢失/损坏赔付",
     "Compensación por extravío o producto dañado",
     "Descuento por extravío o producto dañado",
     "Ajuste por Descuento por extravío o producto dañado",
     "Ajuste por extravío o producto dañado",
     "Pago por extravío o producto dañado bajo el servicio de Fulfillment",
     "Corrección de pago por extravío o producto dañado bajo el servicio de Fulfillment",
     "Reversa de pago por extravío o producto dañado bajo el servicio de Fulfillment",
     "Corrección de reversa de pago por extravío o producto dañado bajo el servicio de Fulfillment")

# ---------------------------------------------------------------- charges
_add(COMMISSION, "销售佣金",
     "Cobro por comisión por venta", "Reembolso por comisión por venta",
     "Corrección de cobro por comisión por venta",
     "Corrección de reembolso por comisión por venta")
_add(COMMISSION, "Celebremos 活动佣金",
     "Comisión Celebremos", "Ajuste comisión Celebremos", "Reembolso comisión Celebremos",
     "Reembolso ajuste comisión Celebremos")
_add(COMMISSION, "申诉后的佣金调整",
     "Cobro Comisión por Apelación", "Reembolso Comisión por Apelación",
     "Corrección comisión por apelación", "Corrección Reembolso  Comisión por Apelación")
_add(SHIP_COFIN, "配送共担费（卖家自配送订单）",
     "Cobro por cofinanciamiento logístico", "Reembolso por cofinanciamiento logístico",
     "Corrección de cobro por cofinanciamiento logístico",
     "Corrección de reembolso por cofinanciamiento logístico")
_add(SHIP_COFIN, "配送共担费（Fulfillment by Falabella 订单）",
     "Cobro por cofinanciamiento logístico, Fulfillment by Falabella",
     "Reembolso por cofinanciamiento logístico, Fulfillment by Falabella",
     "Ajuste por cofinanciamiento logístico, Fulfillment by Falabella",
     "Ajuste de cobro por cofinanciamiento logístico, Fulfillment by Falabella")
_add(SHIP_BUYER, "买家运费转付平台",
     "Reversa de pago de envío comprador", "Ajuste Reversa de pago de envío comprador",
     "Corrección de reversa de pago de envío comprador")
_add(SHIP_PROMO, "falabella.com 免运费促销",
     "Cobro Promo envío falabella.com", "Reembolso por Promo envío falabella.com")
_add(SHIP_PROMO, "卖家出资的免运费促销",
     "Cobro por Promo Envio Seller", "Reembolso por descuento Promo Envio Seller")
_add(SHIP_REVERSE, "退货逆向物流费",
     "Cobro por logística inversa", "Reembolso por logística inversa",
     "Corrección de cobro por logística inversa", "Corrección de reembolso por logística inversa",
     "Cobro por logística inversa, Fulfillment by Falabella",
     "Reembolso por logística inversa, Fulfillment by Falabella",
     "Corrección de cobro por logística inversa, Fulfillment by Falabella",
     "Corrección de reembolso por logística inversa, Fulfillment by Falabella",
     "Return Delivery Fee")
_add(FULFILMENT, "FBF 仓储费",
     "Cobro por almacenamiento de productos, Fulfillment by Falabella",
     "Reembolso por almacenamiento de productos, Fulfillment by Falabella",
     "Corrección de cobro por almacenamiento de productos, Fulfillment by Falabella",
     "Corrección de reembolso por almacenamiento de productos, Fulfillment by Falabella")
_add(FULFILMENT, "FBF 超期仓储费（按 SKU）",
     "Cobro por almacenamiento prolongado por SKU, Fulfillment by Falabella",
     "Reembolso por almacenamiento prolongado por SKU, Fulfillment by Falabella")
_add(FULFILMENT, "FBF 拣货打包费",
     "Cobro por manejo de productos, Fulfillment by Falabella",
     "Reembolso por manejo de productos, Fulfillment by Falabella",
     "Corrección de cobro por manejo de productos, Fulfillment by Falabella",
     "Corrección de reembolso por manejo de productos, Fulfillment by Falabella")
_add(FULFILMENT, "FBF 上门揽收费（卖家仓库取货）",
     "Cobro por recolecciones de productos Fby en bodega del seller nivel seller",
     "Reembolso por cobro por recolecciones de productos Fby en bodega del seller nivel seller")
_add(FULFILMENT, "FBF 取回库存费",
     "Cobro por retiro de inventario, Fulfillment by Falabella",
     "Reembolso por retiro de inventario, Fulfillment by Falabella",
     "Corrección de cobro por retiro de inventario, Fulfillment by Falabella",
     "Corrección de reembolso por retiro de inventario, Fulfillment by Falabella")
_add(FULFILMENT, "FBF 贴标/装袋费",
     "Cobro por etiquetado y embolsado por SKU, Fulfillment by Falabella",
     "Reembolso por etiquetado y embolsado por SKU, Fulfillment by Falabella")
_add(FULFILMENT, "FBF 入仓收货费",
     "Cobro por recepción de productos, Fulfillment by Falabella",
     "Reembolso por recepción de productos, Fulfillment by Falabella")
_add(FULFILMENT, "包装材料费", "Costo de empaques f.com", "Reembolso por costo de empaques f.com")
_add(ADS, "商品推广（Productos patrocinados）",
     "Cobro por Productos patrocinados", "Reembolso por Productos patrocinados")
_add(ADS, "横幅广告",
     "Cobro por Banners ON", "Reembolso por Banners ON",
     "Cobro por Banners home", "Reembolso por Banners home")
_add(ADS, "邮件营销", "Cobro por Mensajería", "Reembolso por Mensajería")
_add(ADS, "站外媒体广告",
     "Cobro por Media Off", "Reembolso por Media Off", "Cobro por Media On",
     "Reembolso por Media On")
_add(ADS, "推荐位", "Cobro por recomendaciones", "Reembolso por recomendaciones")
_add(ADS, "实体店广告", "Cobro por publicidad en tienda", "Reembolso por publicidad en tienda")
_add(ADS, "营销增值服务", "vas de marketing", "Marketing VAS Credit")
_add(ADS, "重点商品/品类推广",
     "Cobro por productos relevantes ", "Reembolso por productos relevantes",
     "Cobro por categorías relevantes ", "Reembolso por categorías relevantes ")
_add(SUBSCRIPTION, "平台订阅月费",
     "Monthly Subscription", "Monthly Subscription Credit",
     "Cobro por Plan marketplace 360", "Reembolso por Plan marketplace 360",
     "Cobro por Mi tienda", "Reembolso por Mi tienda")
_add(SUBSCRIPTION, "卖家增值服务（发展顾问、竞品分析、价格监控、拍照）",
     "Cobro por servicio desarrollo estratégico de sellers",
     "Reembolso de cobro por servicio desarrollo estratégico de sellers",
     "Corrección de cobro por servicio desarrollo estratégico de sellers",
     "Corrección de reembolso de cobro por servicio desarrollo estratégico de sellers",
     "Cobro por servicios de análisis de competitividad",
     "Reembolso por servicios de análisis de competitividad",
     "Cobro por monitor de precios", "Reembolso por monitor de precios",
     "Cobro por servicios de fotografía", "Reembolso por servicios de fotografía")
_add(PENALTY, "卖家原因取消订单的佣金罚款",
     "Cobro por comisión por cancelación", "Reembolso por comisión por cancelación")
_add(PENALTY, "商品尺寸申报不实的物流罚款",
     "Cobro por medidas incorrectas en cofinanciamiento logístico",
     "Reembolso por medidas incorrectas en cofinanciamiento logístico")
_add(PENALTY, "违规罚款", "Penalty", "Penalty (Order Item Level) Credit")
_add(OTHER, "平台付款/扣款更正", "Corrección de pagos ", "Corrección de cobros ")


def _key(name):
    """Compare names the way people would: accents, case and spacing ignored.
    The glossary itself has trailing spaces and double spaces in some names."""
    t = unicodedata.normalize("NFKD", str(name or ""))
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", t).strip().lower()


CATALOG = dict((_key(k), (k.strip(), zh, line)) for k, (zh, line) in _C.items())


def lookup(name):
    """(Chinese name, P&L line) or None when the type is not registered."""
    hit = CATALOG.get(_key(name))
    return (hit[1], hit[2]) if hit else None


def known(name):
    return _key(name) in CATALOG
