# -*- coding: utf-8 -*-
"""Invoice sale -> seller SKU.

The invoice has no SKU. Each billed sale (Número de venta, Número de paquete,
Pago, Número de publicación, Título) is mapped to the seller's SKU in tiers; the
first tier that answers wins and is recorded with its confidence:

  T1   sales export: the invoice's Número de venta is a row's "# de venta"
       (single sale or an item of a multi-product pack), or its Número de
       paquete is a single-item row's "# de venta", or a pack head whose items
       include this listing                                      → 高
  T1b  MercadoPago collection: the same sale (order_id) or payment (Pago) with
       its "SKU Producto" — covers sales the export left out    → 高
  T2   listing id -> SKU from every file that pairs them (sales export, Full
       charges, storage detail, stock, conciliation, MercadoPago); unique
       → 中, several SKUs on the listing (variants) → the most frequent, 低
  T3   title: the same listing title in the sales export        → 低

A sale none of them answers is "未匹配" and listed in ⑩.
"""
import collections
import difflib
import re

from platforms.mercadolibre.bill_report.readers import listing_digits

T1, T1B, T2U, T2M, T3, NONE = "T1", "T1b", "T2", "T2*", "T3", "—"
TIER_CN = {
    T1: "T1 订单号 → 销售报表",
    T1B: "T1b 订单/付款号 → MercadoPago 收款记录",
    T2U: "T2 刊登号 → 唯一 SKU",
    T2M: "T2* 刊登号 → 多个 SKU 取最常见",
    T3: "T3 标题匹配",
    NONE: "未匹配",
}
CONFIDENCE = {T1: "高", T1B: "高", T2U: "中", T2M: "低", T3: "低", NONE: "—"}
TIER_ORDER = [T1, T1B, T2U, T2M, T3, NONE]

Mapping = collections.namedtuple("Mapping", "sku tier source note")


def norm_title(t):
    t = str(t or "").lower()
    t = re.sub(r"\(kit[^)]*\)", "", t)
    return re.sub(r"[^0-9a-záéíóúñü]+", " ", t).strip()


class Mapper(object):
    """Built once per store from whatever files the folder has."""

    def __init__(self, ventas=None, collection=None, cargos_full=None, storage=None, stock=None,
                 conciliation=None):
        self.by_sale = {}           # export "# de venta" (item rows) -> row
        self.single_by_id = {}      # single-item rows by their "# de venta" (may be a pack id)
        self.children = collections.defaultdict(list)     # pack head id -> item rows
        self.by_listing = collections.defaultdict(collections.Counter)   # digits -> Counter(sku)
        self.listing_src = collections.defaultdict(set)
        self.by_title = collections.defaultdict(collections.Counter)
        self.mp_by_order = {}
        self.mp_by_op = {}
        self.titles = {}            # sku -> a title (for display)
        self.prices = collections.defaultdict(lambda: collections.defaultdict(set))  # listing -> sku -> prices
        if ventas is not None and len(ventas):
            for r in ventas.itertuples(index=False):
                if r.kind == "parent":
                    continue
                if r.kind == "child" and r.parent:
                    self.children[r.parent].append(r)
                self.by_sale.setdefault(r.sale_no, r)
                if r.kind == "single":
                    self.single_by_id.setdefault(r.sale_no, r)
                if r.sku:
                    self._listing(r.listing, r.sku, "销售报表", max(1, int(r.units or 1)))
                    if r.unit_price:
                        self.prices[listing_digits(r.listing)][r.sku].add(round(float(r.unit_price), 2))
                    self.by_title[norm_title(r.title)][r.sku] += 1
                    self.titles.setdefault(r.sku, r.title)
        if collection is not None and len(collection):
            for r in collection.itertuples(index=False):
                if not r.sku:
                    continue
                if r.order_id:
                    self.mp_by_order.setdefault(r.order_id, r.sku)
                self.mp_by_op.setdefault(r.op_id, r.sku)
                self._listing(r.item_id, r.sku, "MercadoPago 收款")
        for df, col, src in ((cargos_full, "listing", "Full 费用明细"), (storage, "listing", "仓储费明细")):
            if df is not None and len(df):
                for r in df[["sku", col]].drop_duplicates().itertuples(index=False):
                    if r[0]:
                        self._listing(r[1], r[0], src)
        for df, src in ((stock, "库存报表"), (conciliation, "库存对账")):
            if df is not None and len(df) and "listings" in df.columns:
                for r in df[["sku", "listings", "title"]].drop_duplicates().itertuples(index=False):
                    if not r[0]:
                        continue
                    for part in str(r[1]).split("|"):
                        self._listing(part, r[0], src)
                    self.titles.setdefault(r[0], r[2])

    def _listing(self, listing, sku, src, n=1):
        d = listing_digits(listing)
        if d and sku:
            self.by_listing[d][sku] += n
            self.listing_src[d].add(src)

    def title_of(self, sku):
        return self.titles.get(sku, "")

    # ---- the tiers

    def _t1(self, sale_no, pack_no, listing):
        r = self.by_sale.get(sale_no)
        if r is not None and r.sku:
            return Mapping(r.sku, T1, "销售报表 # de venta %s" % sale_no, "")
        r = self.single_by_id.get(pack_no) if pack_no else None
        if r is not None and r.sku:
            return Mapping(r.sku, T1, "销售报表 # de venta %s（包裹号）" % pack_no, "")
        kids = self.children.get(pack_no) if pack_no else None
        if kids:
            d = listing_digits(listing)
            skus = sorted(set(k.sku for k in kids if k.sku and listing_digits(k.listing) == d))
            if len(skus) == 1:
                return Mapping(skus[0], T1, "销售报表 包裹 %s 内同刊登的商品行" % pack_no, "")
        if r is not None or (sale_no in self.by_sale):
            return Mapping("", None, "", "销售报表有此单但 SKU 为空（卖家未设置 SKU）")
        return None

    def _t1b(self, sale_no, pago):
        sku = self.mp_by_order.get(sale_no) or (self.mp_by_op.get(pago) if pago else None)
        if sku:
            return Mapping(sku, T1B, "MercadoPago 收款记录（订单 %s）" % sale_no, "")
        return None

    def _t2(self, listing, unit_price=0.0):
        d = listing_digits(listing)
        c = self.by_listing.get(d)
        if not c:
            return None
        src = "、".join(sorted(self.listing_src.get(d, ())))
        if len(c) == 1:
            return Mapping(next(iter(c)), T2U, "刊登 %s（%s）" % (d, src), "")
        if unit_price:
            # variants of one listing: the SKU whose sold unit price was this one
            hit = sorted(k for k, ps in self.prices.get(d, {}).items()
                         if any(abs(x - unit_price) <= max(0.01, 0.005 * unit_price) for x in ps))
            if len(hit) == 1:
                return Mapping(hit[0], T2U, "刊登 %s ＋ 单价 %s 只对应这个变体（销售报表）" % (d, unit_price), "")
        top = c.most_common()
        return Mapping(top[0][0], T2M, "刊登 %s（%s）" % (d, src),
                       "该刊登对应多个 SKU：%s" % "、".join("%s×%d" % (k, v) for k, v in top[:5]))

    def _t3(self, title):
        t = norm_title(title)
        if not t:
            return None
        c = self.by_title.get(t)
        if c:
            return Mapping(c.most_common(1)[0][0], T3, "标题完全一致", "")
        best = difflib.get_close_matches(t, list(self.by_title.keys()), n=1, cutoff=0.92)
        if best:
            return Mapping(self.by_title[best[0]].most_common(1)[0][0], T3, "标题近似（≥92%）",
                           "近似标题：%s" % best[0][:60])
        return None

    def sale(self, sale_no, pack_no="", pago="", listing="", title="", unit_price=0.0):
        note = ""
        m = self._t1(sale_no, pack_no, listing)
        if m is not None and m.tier:
            return m
        if m is not None:
            note = m.note
        for m in (self._t1b(sale_no, pago), self._t2(listing, unit_price), self._t3(title)):
            if m is not None:
                return m._replace(note="；".join(x for x in (note, m.note) if x))
        return Mapping("", NONE, "", note or "销售报表、MercadoPago、Full/库存文件都找不到此单或此刊登")

    def listing(self, listing, title=""):
        """A store-level Full charge that names a listing: T2, then T3."""
        for m in (self._t2(listing), self._t3(title)):
            if m is not None:
                return m
        return Mapping("", NONE, "", "刊登号在任何文件里都对不上 SKU")


UNMAPPED_PREFIX = "未匹配·"


def sku_label(m, listing):
    """The SKU key a mapped row is aggregated under in ⑦."""
    if m.sku:
        return m.sku
    d = listing_digits(listing)
    return UNMAPPED_PREFIX + (d or "无刊登号")
