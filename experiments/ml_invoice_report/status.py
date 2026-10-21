# -*- coding: utf-8 -*-
"""What happened to a billed sale — the status classes and their evidence.

Amounts never come from here: a sale's reversal (how much of it was undone) is
read from the invoices (commission reversal rows). This module only decides
*which kind* of reversal it was, from the evidence files:

  sales export   Estado, ship / delivery dates, ML's return review fields
  MercadoPago    collection status / status_detail / shipment status,
                 settlement REFUND rows (money actually sent back)
  Returns_*.xlsx warehouse return triage
  the invoices   a "Cargo por devolución" (return shipping) on the sale

Rules (in this order; the first that applies decides):

  reversal share < 99 %                         → PARTIAL 部分退款
  return evidence (Devolución / Reclamo / Mediación / Cambio in the status,
  ML review done, return fee charged, warehouse triage, buyer-protection
  refund in MercadoPago)                        → RETURN 退货退款
  shipped evidence (ship / delivery date, MercadoPago shipment shipped /
  delivered / not_delivered)                     → RETURN 退货退款 (发货后取消/退回)
  cancel evidence (status says cancelled, MercadoPago cancelled, shipment
  cancelled, or refunded by_admin / by_payer with no shipment)
                                                 → CANCEL 发货前取消
  no export row, no MercadoPago row, reversal ≤ 3 days after the sale and
  no return fee                                  → CANCEL, marked 推定（时滞）
  otherwise                                      → UNPROVEN 已冲销·状态未证实

A sale without any reversal: COMPLETED when delivery is proven (or ML paid the
seller on a claim), otherwise OPEN 未完结/未知 (in transit, open return or claim,
or refund evidence that the invoices have not reversed yet — flagged).
"""
import collections

COMPLETED, CANCEL, RETURN, PARTIAL, UNPROVEN, OPEN = (
    "completed", "cancel", "return", "partial", "unproven", "open")
CLASS_CN = collections.OrderedDict([
    (COMPLETED, "正常完成"),
    (CANCEL, "发货前取消"),
    (RETURN, "退货退款"),
    (PARTIAL, "部分退款"),
    (UNPROVEN, "已冲销·状态未证实"),
    (OPEN, "未完结/未知"),
])
REVERSAL_CLASSES = (CANCEL, RETURN, PARTIAL, UNPROVEN)
FULL = 0.99                 # a reversal share at or above this is a full reversal
CANCEL_LAG_DAYS = 3

RETURN_WORDS = ("devoluci", "reclamo", "mediaci", "cambio", "reembolsamos al comprador")
DELIVERED = ("entregado", "venta entregada")
SHIPPED_MP = ("shipped", "delivered", "not_delivered")

Evidence = collections.namedtuple(
    "Evidence", "ventas mp_status mp_detail mp_shipment mp_refund triage return_fee")
Evidence.__new__.__defaults__ = (None, "", "", "", 0.0, None, 0.0)


def _has(s, words):
    s = (s or "").lower()
    return any(w in s for w in words)


def flags(ev):
    """(cancel, shipped, returned, delivered, seller_paid, refund_seen, texts)."""
    v = ev.ventas
    est = (v.estado if v is not None else "") or ""
    texts = []
    if v is not None:
        texts.append("销售报表：%s" % est)
        if v.shipped_at:
            texts.append("在途 %s" % v.shipped_at)
        if v.delivered_at:
            texts.append("送达 %s" % v.delivered_at)
        if v.reviewed or v.result:
            texts.append("ML 审核：%s %s %s" % (v.result, v.destination, v.result_reason))
        if v.money_to:
            texts.append("款项归：%s" % v.money_to)
    else:
        texts.append("销售报表无此单")
    if ev.mp_status:
        texts.append("MercadoPago：%s/%s，运单 %s" % (ev.mp_status, ev.mp_detail, ev.mp_shipment or "无"))
    if ev.mp_refund:
        texts.append("MercadoPago 退款 %s" % "{:,.2f}".format(ev.mp_refund))
    if ev.triage is not None:
        texts.append("仓库质检：%s / %s" % (ev.triage.review_result, ev.triage.money_status))
    if ev.return_fee:
        texts.append("账单退货处理费 %s" % "{:,.2f}".format(ev.return_fee))
    low = est.lower()
    cancel = ("cancel" in low and not _has(low, RETURN_WORDS)) or ev.mp_status == "cancelled" \
        or ev.mp_shipment == "cancelled" or (
        ev.mp_status == "refunded" and ev.mp_detail in ("by_admin", "by_payer", "expired")
        and ev.mp_shipment in ("", "cancelled", "ready_to_ship", "pending"))
    shipped = bool(v is not None and (v.shipped_at or v.delivered_at)) or ev.mp_shipment in SHIPPED_MP \
        or bool(ev.return_fee) or ev.triage is not None
    # a buyer-protection refund (bpp_*) is a return only once the parcel had left; a
    # buyer who cancels before dispatch is refunded under the same programme
    returned = _has(low, RETURN_WORDS) or bool(v is not None and (v.reviewed or v.result)) \
        or bool(ev.return_fee) or ev.triage is not None \
        or (ev.mp_detail.startswith("bpp_") and ev.mp_shipment in SHIPPED_MP) \
        or ev.mp_status in ("charged_back", "in_mediation")
    delivered = bool(v is not None and v.delivered_at) or _has(low, DELIVERED) or ev.mp_shipment == "delivered"
    seller_paid = "te dimos el dinero" in low      # ML covered the claim: the seller keeps the money
    refund_seen = ev.mp_refund > 0 or ev.mp_status == "refunded" or "reembolso al comprador" in low \
        or bool(v is not None and v.money_to == "Del comprador")
    return cancel, shipped, returned, delivered, seller_paid, refund_seen, "；".join(texts)


def classify_reversal(share, ev, lag_days):
    """(class, basis text, confidence) for a sale whose commission was reversed."""
    cancel, shipped, returned, _d, _p, _r, texts = flags(ev)
    if share < FULL:
        return PARTIAL, "佣金只冲销了 %.0f%%（部分退款）；%s" % (100 * share, texts), "账单"
    if returned:
        return RETURN, texts, "证据"
    if shipped:
        return RETURN, ("发货后取消/退回；" if cancel else "已发货；") + texts, "证据"
    if cancel:
        return CANCEL, texts, "证据"
    if ev.ventas is None and not ev.mp_status and lag_days is not None and lag_days <= CANCEL_LAG_DAYS:
        return CANCEL, "推定（时滞）：冲销距下单 %d 天、无退货费、销售报表与 MercadoPago 收款记录均无此单；%s" % (
            lag_days, texts), "推定"
    return UNPROVEN, texts, "无"


def classify_open(ev):
    """(class, basis text, flag) for a billed sale with no reversal in the files."""
    cancel, shipped, returned, delivered, seller_paid, refund_seen, texts = flags(ev)
    if seller_paid:
        return COMPLETED, "纠纷结案，平台把钱给了卖家；" + texts, ""
    if refund_seen and (returned or cancel):
        return OPEN, "有退款证据但账单尚未冲销；" + texts, "refund_not_reversed"
    if returned:
        return OPEN, "退货/纠纷处理中；" + texts, ""
    if cancel:
        return OPEN, "已取消但账单尚未冲销；" + texts, "cancel_not_reversed"
    if delivered:
        return COMPLETED, texts, ""
    return OPEN, ("在途/未送达；" if shipped else "状态未知；") + texts, ""


def back_to_stock(ev):
    """Did the returned goods go back on sale (so they keep no cost)?"""
    v = ev.ventas
    if v is not None and ("apto para la venta" == (v.result or "").lower() or
                          "de nuevo a la venta" in (v.estado or "").lower()):
        return True
    t = ev.triage
    return t is not None and ("again available" in (t.review_result or "").lower() or
                              "nuevamente" in (t.review_result or "").lower() or
                              "de nuevo a la venta" in (t.review_result or "").lower())
