# -*- coding: utf-8 -*-
"""Every invoice `Detalle` this report knows, and what it does with it.

An invoice row's Detalle that is not here is an **error-level check**; its amount
is still counted (in "其他平台扣款（未登记）") so the ledger ties, but the
workbook says the numbers cannot be trusted until it is registered here.

Fields of an entry:

  cn        Chinese name shown in the workbook
  line      the P&L (②) line it belongs to (PL_LINES keys)
  attrib    how it reaches a SKU:
              "sale"     tied to one sale (Número de venta) → that sale's SKU
              "full_sku" a Full charge whose SKU the Full charges report /
                         the row's listing id identifies
              "storage"  Full storage: by the platform's daily per-SKU storage
                         detail (Costos_por_servicio_almacenamiento)
              "alloc"    store-level: allocated by net GMV share
  reversal  True for "Anulación …" rows (negative amounts that undo a charge)
  kind      "commission" for Cargo por venta (the row that carries the sale
            itself — GMV comes from it); None otherwise
"""
import collections

Entry = collections.namedtuple("Entry", "cn line attrib reversal kind")
Entry.__new__.__defaults__ = (False, None)

SALE, FULL_SKU, STORAGE, ALLOC = "sale", "full_sku", "storage", "alloc"

ATTRIB_CN = {
    SALE: "直接：随订单（按订单号归 SKU）",
    FULL_SKU: "直接：Full 费用明细 / 刊登号归 SKU",
    STORAGE: "直接：按平台仓储费 SKU 日明细归属",
    ALLOC: "按销售额分摊（店铺级费用）",
}

COMMISSION = "Cargo por venta"

CATALOG = collections.OrderedDict([
    # ---- tied to a sale
    ("Cargo por venta", Entry("销售佣金", "fee_commission", SALE, False, "commission")),
    ("Anulación del cargo por venta", Entry("销售佣金冲销（取消/退款）", "fee_commission", SALE, True)),
    ("Cargo por venta con afiliados", Entry("联盟推广佣金", "fee_affiliate", SALE)),
    ("Anulación del cargo por venta con afiliados", Entry("联盟推广佣金冲销", "fee_affiliate", SALE, True)),
    ("Cargo por envíos de Mercado Libre", Entry("平台运费（Mercado Envíos）", "fee_shipping", SALE)),
    ("Anulación del cargo por envíos de Mercado Libre", Entry("平台运费冲销", "fee_shipping", SALE, True)),
    ("Cargo por devolución", Entry("退货处理费", "fee_return", SALE)),
    ("Anulación del cargo por devolución", Entry("退货处理费冲销", "fee_return", SALE, True)),
    # ---- advertising (store level: the invoice row names no listing)
    ("Cargo por campaña de publicidad de Product Ads", Entry("广告：Product Ads", "fee_ads", ALLOC)),
    ("Anulación del cargo por campaña de publicidad de Products Ads",
     Entry("广告冲销：Product Ads", "fee_ads", ALLOC, True)),
    ("Anulación del cargo por campaña de publicidad de Product Ads",
     Entry("广告冲销：Product Ads", "fee_ads", ALLOC, True)),
    ("Cargo por campaña de Publicidad de Seguidores", Entry("广告：粉丝推广（Seguidores）", "fee_ads", ALLOC)),
    ("Anulación del cargo por campaña de Publicidad de Seguidores",
     Entry("广告冲销：粉丝推广", "fee_ads", ALLOC, True)),
    ("Cargo por campaña de publicidad de Display Ads", Entry("广告：Display Ads", "fee_ads", ALLOC)),
    ("Anulación del cargo por campaña de publicidad de Display Ads",
     Entry("广告冲销：Display Ads", "fee_ads", ALLOC, True)),
    # ---- Full warehouse: storage family
    ("Cargo por servicio de almacenamiento Full", Entry("Full 仓储费（日租）", "fee_full_storage", STORAGE)),
    ("Cargo por stock antiguo en Full", Entry("Full 长龄库存费", "fee_full_storage", FULL_SKU)),
    ("Cargo por sobrepasar espacio Full", Entry("Full 超出分配空间费", "fee_full_storage", ALLOC)),
    ("Cargo por comprar más espacio en Full", Entry("Full 购买额外空间费", "fee_full_storage", ALLOC)),
    # ---- Full warehouse: operations / penalties
    ("Cargo por servicio de colecta Full", Entry("Full 上门揽收入仓费", "fee_full_ops", FULL_SKU)),
    ("Cargo por retiro de stock Full", Entry("Full 取回库存费", "fee_full_ops", FULL_SKU)),
    ("Cargo por incumplimiento en Envíos Full", Entry("Full 入仓违约罚款", "fee_full_ops", ALLOC)),
    # ---- store services
    ("Cargo por mantenimiento de Mi página", Entry("店铺页面（Mi página）月费", "fee_store", ALLOC)),
    ("Anulación del cargo por mantenimiento de Mi página",
     Entry("店铺页面月费冲销", "fee_store", ALLOC, True)),
    ("Cargo debitado del dinero en garantía del Programa Beneficio de Reputación",
     Entry("信誉保障计划：从保证金扣款", "fee_other", ALLOC)),
])

UNKNOWN = Entry("未登记的账单科目", "fee_other", ALLOC)

# P&L fee lines in their fixed order: key, Chinese label, original term, note
PL_FEE_LINES = [
    ("fee_commission", "减：销售佣金（净）", "Cargo por venta − Anulación",
     "按计费日；已减去本期冲销（取消/退款退还的佣金）"),
    ("fee_affiliate", "减：联盟推广佣金", "Cargo por venta con afiliados", ""),
    ("fee_shipping", "减：平台运费（净）", "Cargo por envíos de Mercado Libre − Anulación",
     "平台向卖家收取的全额运费；买家付的部分在收入里"),
    ("fee_return", "减：退货处理费（净）", "Cargo por devolución − Anulación", ""),
    ("fee_ads", "减：广告费", "Product Ads / Seguidores / Display Ads", "店铺级，按净销售额分摊到 SKU"),
    ("fee_full_storage", "减：Full 仓储类费用", "almacenamiento / stock antiguo / espacio",
     "日租按平台仓储明细归 SKU；长龄按 SKU；空间费按销售额分摊"),
    ("fee_full_ops", "减：Full 揽收/取回/违约", "colecta / retiro de stock / incumplimiento",
     "揽收、取回按 SKU；违约罚款按销售额分摊"),
    ("fee_store", "减：店铺服务费", "Mantenimiento de Mi página", "店铺级，按净销售额分摊"),
    ("fee_other", "减：其他平台扣款", "Programa Beneficio de Reputación / 未登记科目",
     "店铺级，按净销售额分摊；未登记科目会在⑩报错"),
]
PL_FEE_KEYS = [k for k, _, _, _ in PL_FEE_LINES]
FEE_CN = dict((k, lab.replace("减：", "")) for k, lab, _, _ in PL_FEE_LINES)


def entry(detalle):
    """(Entry, known?) for one Detalle text."""
    d = str(detalle or "").strip()
    if d in CATALOG:
        return CATALOG[d], True
    return UNKNOWN, False
