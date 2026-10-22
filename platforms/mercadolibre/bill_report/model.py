# -*- coding: utf-8 -*-
"""Load one store's download folder and derive every number of the workbook.

The baseline is the ledger: every row of every Facturación and Notas de crédito
file in the folder, each charge once (by Número del cargo), signed as
MercadoLibre states it (charges positive, "Anulación …" negative), **with IVA**.
The period is the calendar month by charge date (Fecha del cargo) — for Chile,
whose bills run from the 8th to the 7th, that is the tail of one invoice plus the
head of the next.

Sales come from the same ledger: each "Cargo por venta" row carries the sale it
charges (Número de venta, quantity, unit price, Total de la venta, buyer-paid
shipping). A sale is in the month's GMV when its commission row's charge date is
in the month. A commission reversal ("Anulación del cargo por venta") undoes the
sale in the share reversed / original commission; that share of Total de la
venta is the reversed sale amount (an estimate — the invoice does not state the
refund itself; MercadoPago's refund is shown next to it as evidence).

S (one dict per store) carries the inputs, the derived tables and S["m"], the
numbers the sheets print. Nothing here writes Excel.
"""
import calendar
import collections
import datetime
import os
import re

import pandas as pd

from core import report_kit as K
from platforms.mercadolibre.bill_report import catalog as C
from platforms.mercadolibre.bill_report import mapping as M
from platforms.mercadolibre.bill_report import profile as P
from platforms.mercadolibre.bill_report import readers as R
from platforms.mercadolibre.bill_report import status as ST

STORE_LEVEL = "（店铺级，无销售可分摊）"
NO_SALE_NO = "（无订单号的随单费用）"

# ------------------------------------------------------------------ period


def month_bounds(month):
    y, m = int(month[:4]), int(month[5:7])
    return datetime.date(y, m, 1), datetime.date(y, m, calendar.monthrange(y, m)[1])


def period_cn(month):
    return "%d 年 %d 月" % (int(month[:4]), int(month[5:7]))


def weeks(p0, p1):
    """Monday-based weeks clipped to the month: [(start, end, label)]."""
    out, d = [], p0
    while d <= p1:
        end = min(p1, d + datetime.timedelta(days=6 - d.weekday()))
        out.append((d, end, "%d/%d–%d/%d" % (d.month, d.day, end.month, end.day)))
        d = end + datetime.timedelta(days=1)
    return out


def week_of(d, wk):
    for i, (a, b, _) in enumerate(wk):
        if a <= d <= b:
            return i
    return None


# ------------------------------------------------------------------ load


def load(store, folder, month, site):
    """Read every file this report uses. Returns S with the raw frames and the
    list of files (name, role, facts) for ①."""
    prof = P.get(site)
    S = {"store": store, "folder": folder, "month": month, "site": prof.site, "prof": prof,
         "meta": R.run_meta(folder), "files": [], "load_errors": []}

    def note(name, role, facts=""):
        S["files"].append((name, role, facts))

    inv, facts = [], []
    for p in R.all_files(folder, "Reporte_Facturacion_MercadoLibre_*.xlsx") + \
            R.all_files(folder, "Reporte_Notas_Credito_MercadoLibre_*.xlsx"):
        try:
            df, f = R.read_invoice(p)
        except Exception as e:                     # noqa: BLE001 — reported as a check
            S["load_errors"].append((os.path.basename(p), str(e)))
            continue
        inv.append(df)
        facts.append(f)
        note(f["name"], "金额（基准）", "%d 行；计费日 %s → %s；平台更新于 %s" % (
            f["rows"], f["first"], f["last"], f["updated"] or "?"))
    S["inv_facts"] = facts
    S["L_raw"] = pd.concat(inv, ignore_index=True) if inv else pd.DataFrame(
        columns=list(set(R.INV_COLS.values())) + ["file", "doc", "bill_month", "row"])

    S["pagos"] = []
    for p in R.all_files(folder, "Reporte_Pagos_Facturas_*.xlsx"):
        try:
            summ, det, f = R.read_pagos(p)
        except Exception as e:                     # noqa: BLE001
            S["load_errors"].append((os.path.basename(p), str(e)))
            continue
        S["pagos"].append((summ, det, f))
        note(f["name"], "核对：账单付款与贷记单", "付款 %d 笔、贷记单 %d 张、付款明细 %d 行" % (
            int((summ["type"] != "Nota de crédito").sum()), int((summ["type"] == "Nota de crédito").sum()),
            len(det)))

    S["ventas"] = None
    p = R.first(folder, "*Ventas*.xlsx")
    if p:
        try:
            S["ventas"], f = R.read_ventas(p)
            note(f["name"], "SKU 查找；状态证据", "%d 行；下单日 %s → %s" % (f["rows"], f["first"], f["last"]))
        except Exception as e:                     # noqa: BLE001
            S["load_errors"].append((os.path.basename(p), str(e)))

    def many(pattern, reader, role, label):
        frames = []
        for path in R.all_files(folder, pattern):
            try:
                df, f = reader(path)
            except Exception as e:                 # noqa: BLE001
                S["load_errors"].append((os.path.basename(path), str(e)))
                continue
            frames.append(df)
            note(f["name"], role, label(df, f))
        return pd.concat(frames, ignore_index=True) if frames else None

    S["collection"] = many("collection-*.csv", R.read_collection, "SKU 查找；状态证据（MercadoPago）",
                           lambda df, f: "%d 笔收款" % len(df))
    S["settlement"] = many("settlement_v2-*.csv", R.read_settlement, "状态证据：实际退款（MercadoPago）",
                           lambda df, f: "%d 行资金流水" % len(df))
    S["cargos_full"] = many("Reporte_Cargos_Full_*.xlsx", R.read_cargos_full, "Full 费用归 SKU（不取金额）",
                            lambda df, f: "%d 行" % len(df))
    S["storage"] = many("*Costos_por_servicio_almacenamiento.xlsx", R.read_storage,
                        "Full 仓储费按 SKU 日明细归属（不取金额）",
                        lambda df, f: "日期 %s → %s" % (f.get("first"), f.get("last")))
    if S["storage"] is not None:
        S["storage"] = S["storage"].drop_duplicates(["inv_code", "day"])
    S["stock"] = S["conciliation"] = S["returns"] = None
    S["stock_snapshot"] = S["conciliation_range"] = ""
    p = R.first(folder, "stock_general_full_*.xlsx")
    if p:
        try:
            S["stock"], f = R.read_stock_general(p)
            S["stock_snapshot"] = f["snapshot"]
            note(f["name"], "⑨ 库存快照；SKU 查找", f["snapshot"])
        except Exception as e:                     # noqa: BLE001
            S["load_errors"].append((os.path.basename(p), str(e)))
    p = R.first(folder, "conciliation_*.xlsx")
    if p:
        try:
            S["conciliation"], f = R.read_conciliation(p)
            S["conciliation_range"] = f.get("range", "")
            note(f["name"], "⑨ 库存进出；SKU 查找", "区间 %s" % f.get("range", ""))
        except Exception as e:                     # noqa: BLE001
            S["load_errors"].append((os.path.basename(p), str(e)))
    p = R.first(folder, "Returns_*.xlsx")
    if p:
        try:
            S["returns"], f = R.read_returns(p)
            note(f["name"], "状态证据：仓库退货质检", "%d 行" % f["rows"])
        except Exception as e:                     # noqa: BLE001
            S["load_errors"].append((os.path.basename(p), str(e)))
    return S


# ------------------------------------------------------------------ derive


def derive(S):
    prof = S["prof"]
    p0, p1 = month_bounds(S["month"])
    S["p0"], S["p1"] = p0, p1
    m = S["m"] = {}
    issues = S["issues"] = collections.defaultdict(list)

    # ---- the ledger: each charge once
    L = S["L_raw"].copy()
    L["order"] = range(len(L))
    dup = L["charge_no"].ne("") & L.duplicated("charge_no", keep="first")
    S["dup_charges"] = L[dup][["file", "row", "charge_no", "detalle", "amount"]].to_dict("records")
    L = L[~dup].copy()
    ents = [C.entry(d) for d in L["detalle"]]
    L["line"] = [e.line for e, _ in ents]
    L["attrib"] = [e.attrib for e, _ in ents]
    L["reversal"] = [e.reversal for e, _ in ents]
    L["known"] = [k for _, k in ents]
    L["cn"] = [e.cn if k else "未登记：" + d for (e, k), d in zip(ents, L["detalle"])]
    L["in_period"] = L["charge_date"].map(lambda d: d is not None and p0 <= d <= p1)
    S["L"] = L
    P_ = L[L["in_period"]].copy()
    S["P"] = P_
    S["unknown"] = sorted(set(L.loc[~L["known"], "detalle"]))

    # ---- sales: one per commission row (any file), plus sales seen only on other sale rows
    sales = collections.OrderedDict()
    com = L[L["detalle"] == C.COMMISSION].sort_values(["charge_date", "order"])
    cohort = []
    for r in com.itertuples(index=False):
        if not r.sale_no:
            continue
        in_p = bool(r.charge_date and p0 <= r.charge_date <= p1)
        if r.sale_no in sales:
            # billed again (e.g. re-billed after a reversal): every billing is GMV
            issues["dup_sale"].append(r.sale_no)
            s0 = sales[r.sale_no]
            s0["n_billed"] += 1
            if in_p:
                ev = dict(s0, qty=r.qty, total=r.sale_total, buyer_ship=r.buyer_ship,
                          charge_date=r.charge_date, commission=r.amount, charge_no=r.charge_no)
                cohort.append(ev)
                s0["billed"] = True
            continue
        sales[r.sale_no] = dict(n_billed=1,
            sale_no=r.sale_no, pack_no=r.pack_no, pago=r.pago, listing=r.listing, title=r.title,
            qty=r.qty, unit_price=r.unit_price, total=r.sale_total, buyer_ship=r.buyer_ship,
            sale_date=r.sale_date, charge_date=r.charge_date, commission=r.amount,
            charge_no=r.charge_no, file=r.file, billed=bool(p0 <= r.charge_date <= p1) if r.charge_date else False,
            source="commission")
        if in_p:
            cohort.append(sales[r.sale_no])
        if r.qty and r.unit_price and abs(r.qty * r.unit_price - r.sale_total) > max(prof.tol, 0.011):
            issues["total_mismatch"].append(r.sale_no)
    sale_rows = L[(L["attrib"] == C.SALE) & (L["sale_no"] != "")].sort_values(["order"])
    for r in sale_rows.itertuples(index=False):
        if r.sale_no not in sales:
            sales[r.sale_no] = dict(
                sale_no=r.sale_no, pack_no=r.pack_no, pago=r.pago, listing=r.listing, title=r.title,
                qty=r.qty, unit_price=r.unit_price, total=r.sale_total, buyer_ship=r.buyer_ship,
                sale_date=r.sale_date, charge_date=None, commission=0.0, charge_no="", file=r.file,
                billed=False, source="other_row", n_billed=0)
    S["sales"] = sales
    S["cohort"] = cohort          # billing events in the period (a sale billed twice is here twice)

    # ---- commission reversals -> reversed share of the sale
    by_charge = L.set_index("charge_no")["amount"].to_dict() if len(L) else {}
    rate_by_listing = collections.defaultdict(list)
    for s in sales.values():
        if s["source"] == "commission" and s["total"] > 0 and s["commission"] > 0:
            rate_by_listing[s["listing"]].append(s["commission"] / s["total"])
    revs = []
    cum = collections.defaultdict(float)
    rv = L[L["detalle"] == "Anulación del cargo por venta"].sort_values(["charge_date", "order"])
    for r in rv.itertuples(index=False):
        s = sales.get(r.sale_no)
        total = (s["total"] if s and s["total"] else r.sale_total) or 0.0
        orig = by_charge.get(r.ref) if r.ref else None
        if orig and orig > 0:
            raw, method = -r.amount / orig, "原佣金行（账单内，按 Cargo que bonifica/anula 对上）"
        elif s and s["source"] == "commission" and s["commission"] > 0:
            raw, method = -r.amount / s["commission"], "同一订单的佣金行"
        elif rate_by_listing.get(r.listing) and total > 0:
            rates = sorted(rate_by_listing[r.listing])
            rate = rates[len(rates) // 2]
            raw = -r.amount / (total * rate)
            method = "原佣金行不在已下载账单中：按同一刊登的佣金率 %.1f%% 推算" % (100 * rate)
            if abs(raw - 1) <= 0.03:
                raw = 1.0
            issues["rev_estimated"].append(r.sale_no)
        else:
            raw, method = 1.0, "原佣金行不在已下载账单中、也无同刊登佣金率：按全额冲销处理"
            issues["rev_assumed_full"].append(r.sale_no)
        cap = float(max(1, s["n_billed"])) if s else 1.0
        share = max(0.0, min(raw, cap - cum[r.sale_no]))
        cum[r.sale_no] += share
        qty = (s["qty"] if s else r.qty) or 0.0
        buyer_ship = (s["buyer_ship"] if s else r.buyer_ship) or 0.0
        revs.append(dict(
            sale_no=r.sale_no, charge_date=r.charge_date, charge_no=r.charge_no, file=r.file,
            commission_rev=r.amount, share=share, method=method, total=total, qty=qty,
            amount=-total * share, buyer_ship=-buyer_ship * share,
            in_period=bool(r.charge_date and p0 <= r.charge_date <= p1),
            post=bool(r.charge_date and r.charge_date > p1),
            sale_date=(s["sale_date"] if s else r.sale_date)))
    S["revs"] = revs
    share_total = dict(cum)

    # ---- per-sale fee rows (all files) and evidence
    fees_by_sale = collections.defaultdict(lambda: collections.defaultdict(float))
    for r in sale_rows.itertuples(index=False):
        key = (r.line, "in" if r.in_period else ("post" if r.charge_date and r.charge_date > p1 else "pre"))
        fees_by_sale[r.sale_no][key] += r.amount
        fees_by_sale[r.sale_no][(r.detalle, "all")] += r.amount
        if r.in_period:
            fees_by_sale[r.sale_no][(r.detalle, "in")] += r.amount
    S["fees_by_sale"] = fees_by_sale

    ventas = S["ventas"]
    v_by_sale, v_single = {}, {}
    if ventas is not None:
        for r in ventas.itertuples(index=False):
            if r.kind == "parent":
                continue
            v_by_sale.setdefault(r.sale_no, r)
            if r.kind == "single":
                v_single.setdefault(r.sale_no, r)
    coll = S["collection"]
    c_by_order, c_by_op = {}, {}
    if coll is not None:
        for r in coll.itertuples(index=False):
            if r.order_id:
                c_by_order.setdefault(r.order_id, r)
            c_by_op.setdefault(r.op_id, r)
    refunds = collections.defaultdict(float)
    taxes = collections.defaultdict(float)
    if S["settlement"] is not None:
        for r in S["settlement"].itertuples(index=False):
            if r.type.upper().startswith("REFUND"):
                refunds[r.source_id] += -r.amount
            taxes[r.source_id] += r.taxes
    triage = {}
    if S["returns"] is not None:
        for r in S["returns"].itertuples(index=False):
            triage.setdefault(r.order, r)

    def evidence(s):
        v = v_by_sale.get(s["sale_no"]) or v_single.get(s["pack_no"])
        c = c_by_order.get(s["sale_no"]) or c_by_op.get(s["pago"])
        fee = fees_by_sale[s["sale_no"]]
        ret_fee = fee.get(("Cargo por devolución", "all"), 0.0)
        mp_ref = refunds.get(s["pago"], 0.0) if s["pago"] else 0.0
        if not mp_ref and c is not None and c.refunded:
            mp_ref = c.refunded
        return ST.Evidence(v, c.status if c is not None else "", c.status_detail if c is not None else "",
                           c.shipment if c is not None else "", mp_ref,
                           triage.get(s["sale_no"]) or triage.get(s["pack_no"]), ret_fee)

    status = {}
    first_rev = {}
    for e in revs:
        first_rev.setdefault(e["sale_no"], e)
    for sale_no, s in sales.items():
        ev = evidence(s)
        sh = min(1.0, share_total.get(sale_no, 0.0))
        if sh > 0:
            e = first_rev[sale_no]
            lag = (e["charge_date"] - s["sale_date"]).days if (e["charge_date"] and s["sale_date"]) else None
            cls, basis, conf = ST.classify_reversal(sh, ev, lag)
            flag = ""
        else:
            cls, basis, flag = ST.classify_open(ev)
            conf = ""
        status[sale_no] = dict(cls=cls, basis=basis, conf=conf, flag=flag, ev=ev, share=sh,
                               back_to_stock=ST.back_to_stock(ev))
        if flag and s["billed"]:
            issues[flag].append(sale_no)
    S["status"] = status
    for e in revs:
        e["cls"] = status[e["sale_no"]]["cls"]

    # ---- SKU mapping of every sale
    mapper = M.Mapper(ventas, coll, S["cargos_full"], S["storage"], S["stock"], S["conciliation"])
    S["mapper"] = mapper
    smap = {}
    for sale_no, s in sales.items():
        mp = mapper.sale(sale_no, s["pack_no"], s["pago"], s["listing"], s["title"], s["unit_price"])
        smap[sale_no] = mp
        s["sku"] = M.sku_label(mp, s["listing"])
    for ev in cohort:
        ev["sku"] = sales[ev["sale_no"]]["sku"]
    S["map"] = smap

    _sku_table(S)
    _pl(S)
    _weekly(S)
    _bridges(S)
    return S


# ------------------------------------------------------------------ per SKU


DIRECT_COLS = ["fee_commission", "fee_affiliate", "fee_shipping", "fee_return", "fee_full_storage",
               "fee_full_ops"]
ALLOC_COLS = ["fee_ads", "fee_full_storage", "fee_full_ops", "fee_store", "fee_other"]


def _new_sku():
    return collections.defaultdict(float)


def _sku_table(S):
    """Everything per SKU: sales, reversals by class, direct fees, then store-level
    charges attributed (Full by SKU, storage by the daily detail) or allocated."""
    sales, P_, mapper = S["sales"], S["P"], S["mapper"]
    sk = collections.defaultdict(_new_sku)
    titles = {}

    def title_for(sku, t):
        if t and sku not in titles:
            titles[sku] = t

    for s in S["cohort"]:
        x = sk[s["sku"]]
        x["units"] += s["qty"]
        x["gmv"] += s["total"]
        x["ship_income"] += s["buyer_ship"]
        x["n_sales"] += 1
        title_for(s["sku"], s["title"])
    for e in S["revs"]:
        if not e["in_period"]:
            continue
        s = sales.get(e["sale_no"])
        sku = s["sku"] if s else M.UNMAPPED_PREFIX + "无订单"
        x = sk[sku]
        x["rev_" + e["cls"]] += e["amount"]
        x["ship_income_rev"] += e["buyer_ship"]
        full = S["status"][e["sale_no"]]["share"] >= ST.FULL
        u = (e["qty"] * e["share"]) if full else 0.0
        x["units_rev"] += u
        if e["cls"] == ST.CANCEL:
            x["units_cancel"] += u
        elif e["cls"] == ST.RETURN and S["status"][e["sale_no"]]["back_to_stock"]:
            x["units_back"] += u
        if s:
            title_for(sku, s["title"])

    # direct, sale-tied fees in the period
    attributed = collections.defaultdict(float)     # (line, how) -> amount (P&L sign), for ③
    by_det = collections.defaultdict(float)         # (detalle, how) -> amount (invoice sign), for ④
    for r in P_[P_["attrib"] == C.SALE].itertuples(index=False):
        if r.sale_no and r.sale_no in sales:
            sku = sales[r.sale_no]["sku"]
        else:
            sku = NO_SALE_NO
        sk[sku]["d:" + r.line] -= r.amount
        attributed[(r.line, "direct")] -= r.amount
        by_det[(r.detalle, "direct")] += r.amount

    # Full charges with a SKU: Full charges report by charge number, else the row's listing
    full_sku = {}
    if S["cargos_full"] is not None:
        for r in S["cargos_full"].itertuples(index=False):
            if r.sku:
                full_sku.setdefault(r.charge_no, r.sku)
    unresolved_full = []
    alloc_rows = []
    for r in P_[P_["attrib"] == C.FULL_SKU].itertuples(index=False):
        sku = full_sku.get(r.charge_no)
        how = "Full 费用明细（按费用号）"
        if not sku and r.listing:
            mp = mapper.listing(r.listing, r.title)
            sku, how = mp.sku, "刊登号 %s → %s" % (r.listing, M.TIER_CN[mp.tier])
        if sku:
            sk[sku]["d:" + r.line] -= r.amount
            attributed[(r.line, "direct")] -= r.amount
            by_det[(r.detalle, "direct")] += r.amount
            S["issues"]["full_how"].append((r.charge_no, sku, how))
        else:
            unresolved_full.append(r)
            alloc_rows.append(r)
    S["unresolved_full"] = unresolved_full

    # storage: invoice charge on day d = storage of day d-1 in the platform's detail
    st = S["storage"]
    stor_stats = dict(exact=0.0, nearest=0.0, none=0.0, report_exact=0.0)
    srows = P_[P_["attrib"] == C.STORAGE]
    by_day = {}
    if st is not None and len(st):
        for day, g in st.groupby("day"):
            by_day[day] = g.groupby("sku")["cost"].sum()
    days = sorted(by_day)
    for d, g in srows.groupby("charge_date"):
        amt = float(g["amount"].sum())
        target = d - datetime.timedelta(days=1)
        if target in by_day:
            shares, kind = by_day[target], "exact"
            stor_stats["report_exact"] += float(shares.sum())
        elif days:
            near = min(days, key=lambda x: abs((x - target).days))
            shares, kind = by_day[near], "nearest"
        else:
            shares, kind = None, "none"
        if shares is None or float(shares.sum()) <= 0:
            for r in g.itertuples(index=False):
                alloc_rows.append(r)
            stor_stats["none"] += amt
            continue
        stor_stats[kind] += amt
        tot = float(shares.sum())
        parts = [(sku, amt * float(v) / tot) for sku, v in shares.items()]
        parts[-1] = (parts[-1][0], amt - sum(p for _, p in parts[:-1]))
        for sku, p in parts:
            sk[sku]["d:fee_full_storage"] -= p
            title_for(sku, mapper.title_of(sku))
        attributed[("fee_full_storage", "direct")] -= amt
        for r in g.itertuples(index=False):
            by_det[(r.detalle, "direct")] += r.amount
    S["storage_stats"] = stor_stats

    # store-level: allocate by net GMV share
    for r in P_[P_["attrib"] == C.ALLOC].itertuples(index=False):
        alloc_rows.append(r)
    for x in sk.values():
        x["net_gmv"] = x["gmv"] + sum(x["rev_" + c] for c in ST.REVERSAL_CLASSES)
    base = dict((k, max(0.0, x["net_gmv"])) for k, x in sk.items() if not k.startswith("（"))
    btot = sum(base.values())
    by_line = collections.defaultdict(float)
    for r in alloc_rows:
        by_line[r.line] += r.amount
        by_det[(r.detalle, "alloc")] += r.amount
    S["alloc_by_line"] = dict(by_line)
    S["alloc_rows"] = alloc_rows
    for line, amt in list(by_line.items()):
        amt = -amt                      # P&L sign: a charge is negative
        attributed[(line, "alloc")] += amt
        if btot <= 0:
            sk[STORE_LEVEL]["a:" + line] += amt
            continue
        keys = sorted((k for k, v in base.items() if v > 0), key=lambda k: base[k])
        done = 0.0
        for k in keys[:-1]:
            part = amt * base[k] / btot
            sk[k]["a:" + line] += part
            done += part
        sk[keys[-1]]["a:" + line] += amt - done
    S["attributed"] = dict(attributed)
    S["by_det"] = dict(by_det)

    for k, x in sk.items():
        x["ship_net"] = x["ship_income"] + x["ship_income_rev"]
        x["direct"] = sum(v for kk, v in x.items() if kk.startswith("d:"))
        x["alloc"] = sum(v for kk, v in x.items() if kk.startswith("a:"))
        x["contrib"] = x["net_gmv"] + x["ship_net"] + x["direct"] + x["alloc"]
        x["units_net"] = x["units"] - x["units_rev"]
        x["cost_units_excluded"] = x["units_cancel"] + x["units_back"]
        if not titles.get(k):
            titles[k] = mapper.title_of(k)
    S["sku"] = sk
    S["sku_titles"] = titles
    # tier per SKU (worst tier among its billed sales)
    tiers = collections.defaultdict(collections.Counter)
    for s in S["cohort"]:
        tiers[s["sku"]][S["map"][s["sale_no"]].tier] += 1
    S["sku_tiers"] = tiers


# ------------------------------------------------------------------ ② P&L numbers


def _pl(S):
    P_, m = S["P"], S["m"]
    sk = S["sku"]
    pl = {}
    pl["gmv_billed"] = sum(s["total"] for s in S["cohort"])
    for c in ST.REVERSAL_CLASSES:
        pl["rev_" + c] = sum(e["amount"] for e in S["revs"] if e["in_period"] and e["cls"] == c)
    pl["ship_income"] = sum(s["buyer_ship"] for s in S["cohort"])
    pl["ship_income_rev"] = sum(e["buyer_ship"] for e in S["revs"] if e["in_period"])
    # fees are charges to the seller: positive in the invoice, negative in the P&L
    for k in C.PL_FEE_KEYS:
        pl[k] = -float(P_.loc[P_["line"] == k, "amount"].sum())
    pl["net_gmv"] = pl["gmv_billed"] + sum(pl["rev_" + c] for c in ST.REVERSAL_CLASSES)
    pl["revenue"] = pl["net_gmv"] + pl["ship_income"] + pl["ship_income_rev"]
    pl["fees_total"] = sum(pl[k] for k in C.PL_FEE_KEYS)
    pl["contrib"] = pl["revenue"] + pl["fees_total"]
    m["pl"] = pl
    m["ledger"] = float(P_["amount"].sum())
    m["n_cohort"] = len(set(s["sale_no"] for s in S["cohort"]))
    m["units_billed"] = sum(s["qty"] for s in S["cohort"])
    prof = S["prof"]
    m["vat_out"] = P.vat_in(pl["revenue"], prof)
    m["vat_in"] = P.vat_in(-pl["fees_total"], prof)
    m["fee_discount"] = float(P_["fee_discount"].sum())
    m["deducted"] = dict((k or "（空）", float(v)) for k, v in P_.groupby("deducted")["amount"].sum().items())
    post = [e for e in S["revs"] if e["post"] and S["sales"].get(e["sale_no"], {}).get("billed")]
    m["post_n"] = len(set(e["sale_no"] for e in post))
    m["post_amount"] = sum(e["amount"] for e in post)
    m["post_commission"] = sum(e["commission_rev"] for e in post)
    # MercadoPago: withheld taxes on the month's billed sales (Mexico: not in this baseline)
    m["mp_taxes"] = None
    if S["settlement"] is not None:
        pagos = set(s["pago"] for s in S["cohort"] if s["pago"])
        st = S["settlement"]
        m["mp_taxes"] = -float(st.loc[st["source_id"].isin(pagos), "taxes"].sum())
    # MercadoPago refunds next to the in-period reversals (evidence)
    m["mp_refund_in"] = sum(S["status"][sn]["ev"].mp_refund
                            for sn in set(e["sale_no"] for e in S["revs"] if e["in_period"]))
    # payments made against the period's invoiced charges
    paid = collections.defaultdict(float)
    for _, det, _f in S["pagos"]:
        for r in det.itertuples(index=False):
            paid[r.charge_no] += r.applied
    S["paid_by_charge"] = paid
    m["paid_in_period"] = float(sum(paid.get(c, 0.0) for c in P_["charge_no"]))
    # cohort status counts
    counts = collections.OrderedDict((c, [0, 0.0]) for c in ST.CLASS_CN)
    for s in S["cohort"]:
        c = S["status"][s["sale_no"]]["cls"]
        counts[c][0] += 1
        counts[c][1] += s["total"]
    m["cohort_status"] = counts
    # mapping coverage of billed GMV
    cov = collections.OrderedDict((t, [0, 0.0]) for t in M.TIER_ORDER)
    for s in S["cohort"]:
        t = S["map"][s["sale_no"]].tier
        cov[t][0] += 1
        cov[t][1] += s["total"]
    m["coverage"] = cov
    m["mapped_share"] = (1.0 - cov[M.NONE][1] / pl["gmv_billed"]) if pl["gmv_billed"] else 1.0
    m["sku_sum"] = dict((k, sum(x[k] for x in sk.values())) for k in (
        "gmv", "net_gmv", "ship_net", "direct", "alloc", "contrib", "units", "units_rev"))


# ------------------------------------------------------------------ ⑧ weekly


def _weekly(S):
    wk = weeks(S["p0"], S["p1"])
    S["weeks"] = wk
    rows = [collections.defaultdict(float) for _ in wk]
    for s in S["cohort"]:
        i = week_of(s["charge_date"], wk)
        rows[i]["n"] += 1
        rows[i]["units"] += s["qty"]
        rows[i]["gmv"] += s["total"]
        rows[i]["ship"] += s["buyer_ship"]
    for e in S["revs"]:
        if e["in_period"]:
            i = week_of(e["charge_date"], wk)
            rows[i]["rev"] += e["amount"]
            rows[i]["ship"] += e["buyer_ship"]
    for r in S["P"].itertuples(index=False):
        i = week_of(r.charge_date, wk)
        rows[i][r.line] -= r.amount
    for x in rows:
        x["net_gmv"] = x["gmv"] + x["rev"]
        x["fees"] = sum(x[k] for k in C.PL_FEE_KEYS)
        x["contrib"] = x["net_gmv"] + x["ship"] + x["fees"]
    S["weekly"] = rows
    # top SKUs by net GMV, weekly net GMV
    top = sorted((k for k in S["sku"] if not k.startswith("（")),
                 key=lambda k: -S["sku"][k]["net_gmv"])[:10]
    tw = dict((k, [0.0] * len(wk)) for k in top)
    for s in S["cohort"]:
        if s["sku"] in tw:
            tw[s["sku"]][week_of(s["charge_date"], wk)] += s["total"]
    for e in S["revs"]:
        s = S["sales"].get(e["sale_no"])
        if e["in_period"] and s and s["sku"] in tw:
            tw[s["sku"]][week_of(e["charge_date"], wk)] += e["amount"]
    S["top_weekly"] = [(k, tw[k]) for k in top]


# ------------------------------------------------------------------ ③ bridges


def _bridges(S):
    L, m = S["L"], S["m"]
    # per invoice document vs the invoice payments report
    refs = L[L["ref"] != ""].groupby("ref")["amount"].sum().to_dict()
    docs = []
    for f in [x for x in S["inv_facts"] if x["doc"] == "FAC"]:
        rows = L[L["file"] == f["name"]]
        no = rows[rows["deducted"] == "No"]
        invoiced = float(no["amount"].sum())
        annulled = float(sum(refs.get(c, 0.0) for c in no["charge_no"]))
        paid = float(sum(S["paid_by_charge"].get(c, 0.0) for c in no["charge_no"]))
        has_pay = any(c in S["paid_by_charge"] for c in no["charge_no"])
        docs.append(dict(name=f["name"], bill_month=f["bill_month"], total=float(rows["amount"].sum()),
                         deducted=float(rows.loc[rows["deducted"] == "Si", "amount"].sum()),
                         invoiced=invoiced, annulled=annulled, expected=invoiced + annulled,
                         paid=paid, has_pay=has_pay, n=len(rows),
                         in_period=bool(rows["in_period"].any())))
    S["docs"] = docs
    # credit notes: NC file (by NC number) vs the payments report's credit-note list
    nc = L[L["doc"] == "NC"].groupby("fiscal_no")["amount"].sum()
    listed = collections.OrderedDict()
    for summ, _d, _f in S["pagos"]:
        for r in summ[summ["type"] == "Nota de crédito"].itertuples(index=False):
            key = r.nc_no if r.nc_no and r.nc_no != "En proceso" else "En proceso#%d" % len(listed)
            listed.setdefault(key, r.total)
    matched, nc_left = [], []
    left = dict(listed)
    for no, amt in nc.items():
        if no in left:
            matched.append((no, -amt, left.pop(no), "号码"))
        else:
            nc_left.append((no, -amt))
    still = []
    for no, amt in nc_left:
        hit = next((k for k, v in left.items() if abs(v - amt) <= S["prof"].tol), None)
        if hit:
            matched.append((no, amt, left.pop(hit), "金额（%s）" % hit))
        else:
            still.append((no, amt))
    # credit notes not issued yet: the NC file groups them under a temporary text, the
    # payments report lists them one by one as "En proceso" — compare the totals
    pend_nc = [(n, a) for n, a in still if not re.match(r"^[A-Z]\d{5,}$", n)]
    pend_pg = [(k, v) for k, v in left.items() if k.startswith("En proceso")]
    if pend_nc and pend_pg and abs(sum(a for _, a in pend_nc) - sum(v for _, v in pend_pg)) <= S["prof"].tol:
        matched.append(("、".join(n for n, _ in pend_nc), sum(a for _, a in pend_nc),
                        sum(v for _, v in pend_pg), "待开票合计（%d 笔 En proceso）" % len(pend_pg)))
        still = [x for x in still if x not in pend_nc]
        for k, _ in pend_pg:
            left.pop(k)
    S["nc_match"] = dict(matched=matched, nc_only=still, pagos_only=list(left.items()))
    m["nc_file_total"] = -float(L.loc[L["doc"] == "NC", "amount"].sum())


# ------------------------------------------------------------------ checks


def residuals(S):
    """name -> residual; every one must be 0 (within the currency's tolerance)."""
    m, pl, ss = S["m"], S["m"]["pl"], S["m"]["sku_sum"]
    res = collections.OrderedDict()
    res["账本 ＝ ② 平台费用合计"] = m["ledger"] + pl["fees_total"]
    res["② 计费 GMV ＝ ⑦ 合计"] = pl["gmv_billed"] - ss["gmv"]
    res["② 净商品销售额 ＝ ⑦ 合计"] = pl["net_gmv"] - ss["net_gmv"]
    res["② 买家运费（净）＝ ⑦ 合计"] = pl["ship_income"] + pl["ship_income_rev"] - ss["ship_net"]
    res["② 平台费用 ＝ ⑦ 直接 ＋ 分摊"] = pl["fees_total"] - (ss["direct"] + ss["alloc"])
    res["② 经营贡献 ＝ ⑦ 合计"] = pl["contrib"] - ss["contrib"]
    att = S["attributed"]
    res["店铺级费用 ＝ ⑦ 分摊列合计"] = sum(v for (l, h), v in att.items() if h == "alloc") - ss["alloc"]
    res["直接归属费用 ＝ ⑦ 直接列合计"] = sum(v for (l, h), v in att.items() if h == "direct") - ss["direct"]
    cov = m["coverage"]
    res["SKU 映射：各级合计 ＝ 计费 GMV"] = sum(v[1] for v in cov.values()) - pl["gmv_billed"]
    cs = m["cohort_status"]
    res["状态分类：各类合计 ＝ 计费 GMV"] = sum(v[1] for v in cs.values()) - pl["gmv_billed"]
    return res


def validate(S):
    st, prof, m = S["store"], S["prof"], S["m"]
    out = []
    ck = lambda *a, **k: out.append(K.check(*a, **k))
    fac = [f for f in S["inv_facts"] if f["doc"] == "FAC"]
    ncs = [f for f in S["inv_facts"] if f["doc"] == "NC"]
    p0, p1 = S["p0"], S["p1"]
    for name, err in S["load_errors"]:
        ck("load:" + name, st, "读取 %s" % name, False, "读不出来：%s" % err)
    # invoices cover the month
    ok, why = bool(fac), "没有任何 Reporte_Facturacion 文件"
    if fac:
        rng = sorted((f["first"], f["last"], f["name"]) for f in fac if f["first"])
        relevant = [x for x in rng if x[1] >= p0 - datetime.timedelta(days=3) and x[0] <= p1]
        gaps = [(a, b) for a, b in zip(relevant, relevant[1:]) if (b[0] - a[1]).days > 3]
        ok = bool(relevant) and relevant[0][0] <= p0 + datetime.timedelta(days=2) and \
            relevant[-1][1] >= p1 - datetime.timedelta(days=2) and not gaps
        why = "；".join("%s：%s → %s" % (n, a, b) for a, b, n in rng)
    ck("invoice_cover", st, "账单覆盖整个月（按计费日）", ok,
       ("账单计费日覆盖 %s → %s。%s" % (p0, p1, why)) if ok else
       ("账单没有覆盖 %s → %s（缺一张账单？）。现有：%s" % (p0, p1, why)))
    cover_end = [f for f in fac if f["first"] and f["first"] <= p1]
    last_fac = max(cover_end, key=lambda f: f["first"]) if cover_end else None
    if last_fac:
        upd = last_fac["updated"][:10]
        closed = bool(upd) and upd > (p1 + datetime.timedelta(days=1)).isoformat()
        ck("invoice_closed", st, "覆盖月末的账单已在月后更新", closed,
           "%s 平台更新于 %s（月末 %s）" % (last_fac["name"], last_fac["updated"] or "?", p1),
           severity="warn")
    used = set(S["P"]["file"]) if "P" in S else set()
    open_ = [f["name"] for f in fac if f.get("open") and f["name"] in used]
    S["provisional"] = open_
    ck("invoice_open", st, "本月用到的账单都已关账", not open_,
       ("%s 尚未关账（平台注明“podría tener cambios hasta el cierre”）：本月落在它里面的费用以后还可能变，"
        "报表为暂定，关账后请重出" % "、".join(open_)) if open_ else "是", severity="warn")
    months = set(f["bill_month"] for f in fac)
    nc_months = set(f["bill_month"] for f in ncs)
    miss_nc = sorted(months - nc_months)
    ck("nc_files", st, "每张账单都有贷记单文件", not miss_nc,
       "缺：%s（该月可能没有开过贷记单；有的话请补下载）" % "、".join(miss_nc) if miss_nc else
       "贷记单文件：%s" % "、".join(f["name"] for f in ncs), severity="warn")
    ck("unknown_detalle", st, "账单科目都已登记", not S["unknown"],
       "未登记：%s —— 金额已计入“其他平台扣款”，请在 catalog.py 登记后重出" % "；".join(S["unknown"])
       if S["unknown"] else "共 %d 种科目，全部在目录里" % n_detalles(S), severity="error")
    ck("dup_charges", st, "同一费用号只出现一次", not S["dup_charges"],
       "%d 行重复（同一 Número del cargo 出现在两个文件），已只计一次" % len(S["dup_charges"])
       if S["dup_charges"] else "没有重复", severity="warn")
    iss = S["issues"]
    ck("dup_sale", st, "每笔销售只计费一次", not iss["dup_sale"],
       "%d 笔销售被计费不止一次（冲销后重新计费）：每次计费都计入 GMV，冲销按次抵减：%s" % (
           len(iss["dup_sale"]), "、".join(iss["dup_sale"][:10])) if iss["dup_sale"] else "是",
       severity="warn")
    ck("total_identity", st, "Total de la venta ＝ 数量 × 单价", not iss["total_mismatch"],
       "%d 笔不等：%s" % (len(iss["total_mismatch"]), "、".join(iss["total_mismatch"][:10]))
       if iss["total_mismatch"] else "全部相等", severity="warn")
    est = iss["rev_estimated"] + iss["rev_assumed_full"]
    ck("rev_method", st, "冲销比例都能从原佣金行算出", not est,
       "%d 笔的原佣金行不在已下载账单中（多为上月的销售）：%d 笔按同刊登佣金率推算，%d 笔按全额" % (
           len(est), len(iss["rev_estimated"]), len(iss["rev_assumed_full"])) if est else "全部能算出",
       severity="warn")
    unp = [e for e in S["revs"] if e["in_period"] and e["cls"] == ST.UNPROVEN]
    ck("unproven", st, "本期冲销都有状态证据", not unp,
       "%d 笔冲销找不到取消或退货的证据，金额 %s，单列在②“其他冲销”" % (
           len(unp), "{:,.2f}".format(sum(e["amount"] for e in unp))) if unp else "是", severity="warn")
    for flag, name in (("refund_not_reversed", "有退款证据但账单未冲销的销售"),
                       ("cancel_not_reversed", "已取消但账单未冲销的销售")):
        xs = iss[flag]
        ck(flag, st, name, not xs, ("%d 笔：%s —— 冲销可能在下一张账单里；本期按已计费列示" % (
            len(xs), "、".join(xs[:12]))) if xs else "没有", severity="warn")
    # mapping coverage
    share = m["mapped_share"]
    ck("map_cover", st, "SKU 映射覆盖率（按计费 GMV）≥ 99%", share >= 0.99,
       "%.2f%% 的计费 GMV 已映射到 SKU（目标 ≥ 99%%）；未匹配 %d 笔，见⑦表二与⑩" % (
           100 * share, m["coverage"][M.NONE][0]),
       severity="error" if share < 0.95 else "warn")
    if S["ventas"] is None:
        ck("ventas", st, "销售报表（仅用于 SKU 与状态）", False, "没有 *Ventas*.xlsx：SKU 只能靠 MercadoPago "
           "与刊登号，状态证据变少", severity="warn")
    mp = S["collection"] is not None or S["settlement"] is not None
    ck("mp", st, "MercadoPago 文件（状态证据）", mp,
       "有：%s" % "、".join(n for n, r, _ in S["files"] if "MercadoPago" in r) if mp else
       "本文件夹没有 MercadoPago 文件：取消/退货只能靠销售报表和冲销时滞判断，退款金额无旁证",
       severity="warn", skipped=not mp)
    ss = S["storage_stats"]
    stot = ss["exact"] + ss["nearest"] + ss["none"]
    if stot:
        ck("storage_attr", st, "Full 仓储费按 SKU 日明细归属", ss["none"] == 0 and ss["nearest"] <= 0.25 * stot,
           "本期仓储费 %s：%s 按当日明细，%s 按最近一天的结构，%s 无明细按销售额分摊" % tuple(
               "{:,.2f}".format(x) for x in (stot, ss["exact"], ss["nearest"], ss["none"])), severity="warn")
    if S["unresolved_full"]:
        ck("full_sku", st, "Full 费用都能归到 SKU", False,
           "%d 行找不到 SKU，按销售额分摊：%s" % (len(S["unresolved_full"]), "、".join(
               "%s %s" % (r.detalle, r.charge_no) for r in S["unresolved_full"][:6])), severity="warn")
    # documents vs payments
    for d in S["docs"]:
        resid = d["expected"] - d["paid"]
        if not d["in_period"]:
            ck("pay:" + d["name"], st, "%s 付款核对" % d["name"], True,
               "这张账单没有本期计费日的行，只在③列示：应付 %s，已付 %s" % (
                   "{:,.2f}".format(d["expected"]), "{:,.2f}".format(d["paid"])), severity="warn", skipped=True)
        elif not d["has_pay"]:
            ck("pay:" + d["name"], st, "%s 已付款" % d["name"], False,
               "付款报表还没有这张账单的付款明细（账单未到期或未付）：应付 %s" % "{:,.2f}".format(d["expected"]),
               severity="warn")
        else:
            ck("pay:" + d["name"], st, "%s 月结费用 ＝ 已付款" % d["name"], abs(resid) <= prof.tol,
               "应付 %s，已付 %s，差 %s" % tuple("{:,.2f}".format(x) for x in (d["expected"], d["paid"], resid)),
               severity="warn")
    nm = S["nc_match"]
    ck("nc_match", st, "贷记单文件 ＝ 付款报表里的贷记单", not nm["nc_only"] and all(
        abs(a - b) <= prof.tol for _, a, b, _ in nm["matched"]),
       "对上 %d 张；只在贷记单文件 %d 张（%s）；只在付款报表 %d 张（%s）" % (
           len(nm["matched"]), len(nm["nc_only"]), "、".join("%s %s" % (n, "{:,.2f}".format(a))
                                                           for n, a in nm["nc_only"][:6]),
           len(nm["pagos_only"]), "、".join("%s %s" % (n, "{:,.2f}".format(a)) for n, a in nm["pagos_only"][:6])),
       severity="warn")
    # residuals
    bad = [(k, v) for k, v in residuals(S).items() if abs(v) > prof.tol]
    ck("residuals", st, "全部残差为 0", not bad,
       "；".join("%s：%s" % (k, "{:,.4f}".format(v)) for k, v in bad) if bad else
       "%d 条勾稽残差全部为 0" % len(residuals(S)), severity="error")
    if not S["prof"].has_full:
        ck("full", st, "Full 仓（库存、仓储）", True, "%s没有 Full：⑨ 不适用" % prof.country_cn,
           severity="warn", skipped=True)
    elif S["stock"] is None:
        ck("stock", st, "Full 库存快照（⑨）", False, "没有 stock_general_full 文件：⑨ 只列销量",
           severity="warn", skipped=True)
    return out


def n_detalles(S):
    return int(S["L"]["detalle"].nunique())
