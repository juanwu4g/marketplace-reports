# -*- coding: utf-8 -*-
"""MercadoPago files of a download folder, per payment (收款号 = SOURCE_ID =
the invoice's `Pago` column = Cobros' operation_id).

What the experiment takes from MercadoPago, and from which column (all checked
on the real downloads of 2026-10, see README):

* **Refund to the buyer** — settlement_v2 `REFUND` rows: −TRANSACTION_AMOUNT is the
  amount returned at list price (the same price the invoice's `Total de la venta`
  carries). `DISPUTE` rows with a non-zero TRANSACTION_AMOUNT are mediations /
  chargebacks the buyer won: money back to the buyer as well, kept apart.
  A `DISPUTE` row with TRANSACTION_AMOUNT 0 and a negative FEE_AMOUNT is an extra
  charge to the seller in the dispute (no money back to the buyer).
* **Seller-funded coupon** — not a column of settlement_v2: on a `SETTLEMENT` row
  TRANSACTION_AMOUNT + FEE_AMOUNT + TAXES_AMOUNT − REAL_AMOUNT is what was taken
  besides fee and tax. It equals the coupon the monthly report derives from Ventas
  (Total gap) order by order (TANKE_EE 2026-09: 847 of 860 orders to the cent).
  A refund gives it back (the same residual on the REFUND / DISPUTE row, negative).
  Cobros' `coupon_fee` ("Descuento a tu contraparte") is NOT it: on TANKE_EE it is
  the whole campaign discount (183.69 where the seller bore 115.12) or 0 where the
  seller bore 117.71.
* **Withheld tax (Mexico)** — TAXES_AMOUNT.
* **Cash** — REAL_AMOUNT by TRANSACTION_DATE, and MONEY_RELEASE_DATE for the release.

Without settlement_v2, Cobros (`collection-*.csv`) gives `amount_refunded` per
payment (refunds and mediations together): on TANKE_EE it equals settlement's
REFUND + DISPUTE on 876 of 878 payments. The coupon cannot be had from Cobros:
transaction − marketplace fee − shipping − tax (reserve-release) − net received
matched settlement on only 451 of 756 payments (Cobros' shipping_cost is not what
was deducted), so without settlement_v2 the coupon is "无法得到". Nothing here
fails on a missing or unreadable file: the result says what was there and what
was used."""

import glob
import os
import re

import pandas as pd

from platforms.mercadolibre.report.loading import _srcid

from experiments.ml_invoice_basis.common import key

GLOBS = (("settlement", "settlement_v2*.csv"), ("collection", "collection-*.csv"),
         ("release", "reserve-release-*.csv"), ("after_collection", "after_collection-*.csv"),
         ("withdraw", "withdraw-*.csv"), ("account_statement", "account_statement*.csv"))
FILE_ZH = {"settlement": "结算流水 settlement_v2", "collection": "收款 Cobros (collection)",
           "release": "资金释放 Liberaciones (reserve-release)",
           "after_collection": "售后 Poscobro (after_collection)", "withdraw": "提现 Retiros (withdraw)",
           "account_statement": "余额与流水 (account_statement)"}
NO_FILES = "本次下载没有 MercadoPago 文件"
SALE_TYPES = ("SETTLEMENT", "REFUND", "DISPUTE")


def files(folder):
    out = {}
    for k, pat in GLOBS:
        hits = sorted(p for p in glob.glob(os.path.join(folder, pat))
                      if not os.path.basename(p).startswith("~$"))
        out[k] = hits[-1] if hits else None
    return out


def _local_day(v):
    """'2026-10-13T13:35:42.000-06:00' -> 2026-10-13 (the date as written, i.e. local)."""
    if not isinstance(v, str) or len(v) < 10:
        return pd.NaT
    return pd.to_datetime(v[:10], errors="coerce")


def _num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _read_csv(path):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(path, sep=";", dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
        except Exception:
            return None
    return None


def _short_cols(df):
    """Cobros / Poscobro headers 'Monto (amount)' -> 'amount'."""
    df = df.copy()
    df.columns = [re.sub(r".*\(([^()]*)\)\s*$", r"\1", str(c)).strip() for c in df.columns]
    return df


# ════════════════════════════════════════════════════════════════════════
# settlement_v2 -> rows and per-payment numbers
# ════════════════════════════════════════════════════════════════════════

def settlement_rows(path):
    d = _read_csv(path) if path else None
    if d is None or "SOURCE_ID" not in d.columns:
        return None
    out = pd.DataFrame({
        "sid": d["SOURCE_ID"].map(_srcid), "type": d["TRANSACTION_TYPE"].fillna("").astype(str),
        "method": d.get("PAYMENT_METHOD_TYPE", pd.Series("", index=d.index)).fillna(""),
        "unit": d.get("BUSINESS_UNIT", pd.Series("", index=d.index)).fillna(""),
        "trans": _num(d["TRANSACTION_AMOUNT"]), "fee": _num(d["FEE_AMOUNT"]),
        "tax": _num(d["TAXES_AMOUNT"]), "real": _num(d["REAL_AMOUNT"]),
        "tdate": d["TRANSACTION_DATE"].map(_local_day),
        "rdate": d.get("MONEY_RELEASE_DATE", pd.Series("", index=d.index)).map(_local_day)})
    out["imp"] = out["trans"] + out["fee"] + out["tax"] - out["real"]
    return out


def payments_from_settlement(R):
    """One row per payment (SOURCE_ID) of the sale-type rows."""
    s = R[R["type"].isin(SALE_TYPES) & (R["sid"] != "")]
    if not len(s):
        return pd.DataFrame()
    st, rf, dp = (s["type"] == "SETTLEMENT"), (s["type"] == "REFUND"), (s["type"] == "DISPUTE")
    money_back = dp & (s["trans"].abs() > 0.004)
    charge = dp & ~money_back
    g = pd.DataFrame({
        "sid": s["sid"],
        "settle": s["trans"].where(st, 0.0), "settle_fee": s["fee"].where(st, 0.0),
        "coupon": s["imp"].where(st, 0.0),
        "refund": (-s["trans"]).where(rf, 0.0), "dispute": (-s["trans"]).where(money_back, 0.0),
        "dispute_charge": (-s["real"]).where(charge, 0.0),
        "coupon_back": (-s["imp"]).where(rf | money_back, 0.0),
        "fee_all": s["fee"], "tax_all": s["tax"], "real_all": s["real"],
        "settle_day": s["tdate"].where(st), "refund_day": s["tdate"].where(rf | dp)})
    agg = g.groupby("sid").agg(
        settle=("settle", "sum"), settle_fee=("settle_fee", "sum"), coupon=("coupon", "sum"),
        refund=("refund", "sum"), dispute=("dispute", "sum"),
        dispute_charge=("dispute_charge", "sum"), coupon_back=("coupon_back", "sum"),
        fee_all=("fee_all", "sum"), tax_all=("tax_all", "sum"), real_all=("real_all", "sum"),
        settle_day=("settle_day", "min"), refund_day=("refund_day", "max"))
    agg["n_settle"] = s[st].groupby("sid").size().reindex(agg.index).fillna(0).astype(int)
    agg["coupon_known"] = True
    return agg


# ════════════════════════════════════════════════════════════════════════
# fallback: Cobros (+ reserve-release for the tax)
# ════════════════════════════════════════════════════════════════════════

NOT_COLLECTED = ("rejected", "cancelled", "pending", "in_process", "authorized")


def payments_from_collection(cpath, rpath):
    c = _read_csv(cpath) if cpath else None
    if c is None:
        return pd.DataFrame(), None
    c = _short_cols(c)
    need = {"operation_id", "status", "transaction_amount", "amount_refunded"}
    if not need <= set(c.columns):
        return pd.DataFrame(), None
    tax = {}          # reserve-release's TAXES_AMOUNT: the withheld tax, for reference only
    r = _read_csv(rpath) if rpath else None
    if r is not None and {"SOURCE_ID", "DESCRIPTION", "TAXES_AMOUNT"} <= set(r.columns):
        p = r[r["DESCRIPTION"].fillna("") == "payment"]
        for sid, t in zip(p["SOURCE_ID"].map(_srcid), _num(p["TAXES_AMOUNT"])):
            if sid:
                tax[sid] = tax.get(sid, 0.0) + float(t)
    col = lambda n: _num(c[n]) if n in c.columns else pd.Series(0.0, index=c.index)
    st = c["status"].fillna("").astype(str).str.strip()
    collected = ~st.isin(NOT_COLLECTED)
    tamt = col("transaction_amount")
    mfee, ship = col("marketplace_fee"), col("shipping_cost")
    net = col("net_received_amount")
    refunded = col("amount_refunded")
    sid = c["operation_id"].map(_srcid)
    rows = []
    for i in c.index:
        k = sid[i]
        if not k:
            continue
        rows.append({"sid": k, "settle": float(tamt[i]) if collected[i] else 0.0,
                     "settle_fee": float(mfee[i] + ship[i]) if collected[i] else 0.0,
                     "coupon": 0.0, "refund": float(refunded[i]),
                     "dispute": 0.0, "dispute_charge": 0.0, "coupon_back": 0.0,
                     "fee_all": float(mfee[i] + ship[i]) if collected[i] else 0.0,
                     "tax_all": tax.get(k, 0.0), "real_all": float(net[i]),
                     "settle_day": pd.to_datetime(str(c.at[i, "date_created"])[:10], dayfirst=True,
                                                  errors="coerce") if "date_created" in c.columns
                     else pd.NaT, "refund_day": pd.NaT, "n_settle": int(collected[i]),
                     "coupon_known": False})
    P = pd.DataFrame(rows)
    if not len(P):
        return P, None
    P = P.groupby("sid").agg(dict(
        settle="sum", settle_fee="sum", coupon="sum", refund="sum", dispute="sum",
        dispute_charge="sum", coupon_back="sum", fee_all="sum", tax_all="sum", real_all="sum",
        settle_day="min", refund_day="max", n_settle="sum", coupon_known="all"))
    return P, (P["settle_day"].min(), P["settle_day"].max())


# ════════════════════════════════════════════════════════════════════════
# everything of one folder
# ════════════════════════════════════════════════════════════════════════

def load(folder):
    """{"files", "source", "pay" (per payment), "span", "rows" (settlement rows or None),
    "release", "withdraw", "after", "account", "notes"} — never raises."""
    F = files(folder)
    out = {"files": F, "source": None, "pay": pd.DataFrame(), "span": None, "rows": None,
           "release": None, "withdraw": None, "after": None, "account": None, "notes": []}
    if not any(F.values()):
        out["notes"].append(NO_FILES + "（settlement_v2 / collection / reserve-release / "
                                       "after_collection 都没有）。")
        return out
    R = settlement_rows(F["settlement"]) if F["settlement"] else None
    if R is not None and len(R):
        out["rows"] = R
        out["pay"] = payments_from_settlement(R)
        out["source"] = "settlement"
        out["span"] = (R["tdate"].min(), R["tdate"].max())
    elif F["collection"]:
        P, span = payments_from_collection(F["collection"], F["release"])
        if len(P):
            out["pay"], out["source"], out["span"] = P, "collection", span
            out["notes"].append("没有 settlement_v2：实际退款取 Cobros 的 amount_refunded（退款与争议不分开；"
                                "TANKE_EE 实测与结算流水 878 笔里 876 笔相同）；优惠券无法得到（Cobros 推算的优惠券"
                                "与结算流水只有 451/756 笔相同，不采用）；现金核对做不了（没有逐笔交易日/释放日）。")
    if out["source"] is None:
        out["notes"].append("有 MercadoPago 文件，但没有可用的 settlement_v2 或 Cobros：退款、优惠券无法得到。")
    if F["release"]:
        d = _read_csv(F["release"])
        if d is not None and {"DATE", "DESCRIPTION", "NET_CREDIT_AMOUNT", "NET_DEBIT_AMOUNT"} <= set(d.columns):
            out["release"] = pd.DataFrame({
                "day": d["DATE"].map(_local_day), "desc": d["DESCRIPTION"].fillna("").astype(str),
                "sid": d.get("SOURCE_ID", pd.Series("", index=d.index)).map(_srcid),
                "net": _num(d["NET_CREDIT_AMOUNT"]) - _num(d["NET_DEBIT_AMOUNT"])})
    if F["withdraw"]:
        d = _read_csv(F["withdraw"])
        if d is not None:
            d = _short_cols(d)
            if {"date_created", "status", "amount"} <= set(d.columns):
                out["withdraw"] = pd.DataFrame({
                    "day": pd.to_datetime(d["date_created"].astype(str).str[:10], errors="coerce"),
                    "status": d["status"].fillna("").astype(str), "amount": _num(d["amount"])})
    if F["after_collection"]:
        d = _read_csv(F["after_collection"])
        if d is not None:
            d = _short_cols(d)
            if {"flow", "status"} <= set(d.columns):
                out["after"] = pd.DataFrame({
                    "flow": d["flow"].fillna("").astype(str), "status": d["status"].fillna("").astype(str),
                    "order": d.get("order_id", pd.Series("", index=d.index)).map(key),
                    "sid": d.get("operation_id", pd.Series("", index=d.index)).map(_srcid),
                    "amount": _num(d.get("amount", pd.Series(0, index=d.index)))})
    if F["account_statement"]:
        d = _read_csv(F["account_statement"])
        if d is not None and {"INITIAL_BALANCE", "FINAL_BALANCE"} <= set(d.columns) and len(d):
            v = lambda c: float(str(d[c].iloc[0]).replace(",", "")) if c in d.columns else None
            try:
                out["account"] = {c: v(c) for c in ("INITIAL_BALANCE", "CREDITS", "DEBITS",
                                                    "FINAL_BALANCE")}
            except ValueError:
                out["account"] = None
    return out


def source_text(mp):
    if mp["source"] == "settlement":
        s0, s1 = mp["span"]
        return "结算流水 settlement_v2（交易日 %s 至 %s）" % (_day(s0), _day(s1))
    if mp["source"] == "collection":
        s0, s1 = mp["span"]
        return "收款报表 Cobros（下单日 %s 至 %s；没有 settlement_v2）" % (_day(s0), _day(s1))
    if any(mp["files"].values()):
        return "有 MercadoPago 文件，但没有 settlement_v2 / Cobros"
    return NO_FILES


def _day(d):
    return d.strftime("%Y-%m-%d") if d is not None and pd.notna(d) else "?"


# ════════════════════════════════════════════════════════════════════════
# invoice sale <-> payment
# ════════════════════════════════════════════════════════════════════════

def sale_payments(B):
    """{sale: {payment}}, {payment: {sale}}, {sale: list price total} over every
    downloaded invoice row (the `Pago` column of the sale's rows)."""
    s2p, p2s, tot = {}, {}, {}
    if "Pago" not in B.columns:
        return s2p, p2s, tot
    for ks, pg in zip(B["ksale"], B["Pago"].map(_srcid)):
        if ks and pg:
            s2p.setdefault(ks, set()).add(pg)
            p2s.setdefault(pg, set()).add(ks)
    if "Total de la venta" in B.columns:
        t = pd.to_numeric(B["Total de la venta"], errors="coerce")
        for ks, v in zip(B["ksale"], t):
            if ks and pd.notna(v) and ks not in tot:
                tot[ks] = float(v)
    return s2p, p2s, tot


def weight(sale, pay, p2s, tot):
    """The sale's share of a payment shared by several billed sales (by list price)."""
    sales = p2s.get(pay, {sale})
    if len(sales) <= 1:
        return 1.0
    s = sum(tot.get(x, 0.0) for x in sales)
    return (tot.get(sale, 0.0) / s) if s else 1.0 / len(sales)
