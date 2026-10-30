# -*- coding: utf-8 -*-
"""MercadoLibre report — derive(): one store's files -> the dict of numbers S.

S["m"] holds every figure the sheets and checks use (GMV, fees per Detalle,
withheld tax, bridges …). The accounting rules live here; sheets only lay out
what derive() computed."""

import pandas as pd

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (CH_NA, CH_NO, CH_SI, FEE_BY_ES, FEE_CATALOG,
    F_DETALLE, MONTH_FULL_ES, STATES_BACK_ON_SALE, STATES_SELLER_KEEPS, STATES_STOCK_LOST,
    V_PACK_FLAG, V_SKU, V_STATUS, V_UNITS, num)
from platforms.mercadolibre.report.loading import (STORAGE_BILL_LAG_DAYS, _cycle_file,
    bill_cycles, credit_note_audit, infer_cycle_day, load_billing, load_cargos_full,
    load_cargos_full_cycles, load_pagos, load_pagos_cycles, load_returns, load_run_meta,
    load_settlement, load_stock, load_storage, load_storage_daily, load_ventas, mp_by_order,
    open_bills, payable_of, stock_snapshot_date)


# ════════════════════════════════════════════════════════════════════════
# 推导层 —— 把"发现"固化成公式
# ════════════════════════════════════════════════════════════════════════

def commission_column_match(A, L):
    """智利"没有代扣"的逐单证据：Ventas 佣金列与账单纯佣金逐单相等的订单数。

    墨西哥这两者差 ~9%（代扣税），所以这个比对是区分两种税制最直接的办法：
    若智利哪天开始代扣，相等的比例会掉到接近 0，no_withholding 校验随之失败。
    只比能按订单号直接挂上账单佣金、且佣金列不为 0 的订单；包裹（佣金记在母
    行、账单记在子行）和退货（佣金列并进了退货处理费）本来就不相等，照样算在
    分母里 —— 比例因此是保守的。
    """
    ids = set(A["oid"]) | set(A["pid"])
    ids.discard(""); ids.discard("nan")
    com = L[(~L["is_rev"]) & L["parent_detalle"].map(
        lambda d: FEE_BY_ES.get(str(d), (None, None, None))[1] == "commission").astype(bool)]
    by = {}
    for ks, kp, amt in zip(com["k_sale"].astype(str).str.strip(),
                           com["k_pack"].astype(str).str.strip(), com["amount"]):
        k = ks if ks in ids else (kp if kp in ids else None)
        if k:
            by[k] = by.get(k, 0.0) + float(amt)
    n = of = 0
    for oid, v in zip(A["oid"], -num(A, site.V_COMMISSION)):
        b = by.get(oid)
        if b is None or abs(v) <= site.ZERO_TOL:
            continue
        of += 1
        if abs(float(v) - b) <= site.TOL_CENT:
            n += 1
    return {"com_match_n": n, "com_match_of": of}


def _derive_pagos_cycles(S, m, folder, B, FA, p0, p1):
    """智利的桥 B 输入：会计月内的费用实际付了多少。

    应付（m["payable"]）照旧按计费日在会计月内的 No 通道费用算；实付取两张
    账单的付款明细页里"付给计费日在会计月内的费用"的那部分。另外逐张账单记
    下"整张账单应付 vs 整张账单实付"，供桥 B 和 ⑩ 页说明（两张都应各自平）。

    贷记单（pay_nc）取计费日在会计月内的 Notas de Crédito 行：智利的 Pagos
    汇总页把贷记单记在"用来抵哪一期"的账单下，和计费日对不上自然月。
    """
    cyc = load_pagos_cycles(folder, S["bill_cycles"], S["load_errors"])
    frames, info, why = [], [], []
    in_month, types, dates = 0.0, set(), []
    for c in cyc:
        y, mth = c["bill"]
        label = "%s %d" % (MONTH_FULL_ES[mth], y)
        span = "%s–%s" % (c["start"].strftime("%m/%d"), c["end"].strftime("%m/%d"))
        P, D = c["P"], c["D"]
        rows = B[(B["cdate"] >= c["start"]) & (B["cdate"] <= c["end"])]
        no, rev = payable_of(rows)
        paid_total, nc_ids = 0.0, set()
        if len(P) and "Tipo de pago" in P.columns:
            P = P.copy()
            P["_bill"] = label
            frames.append(P)
            tp = P["Tipo de pago"].astype(str)
            nota = tp.str.contains("Nota", na=False)
            paid = ~nota & tp.str.strip().ne("") & tp.ne("nan")
            paid_total = float(pd.to_numeric(P[paid]["Importe total"], errors="coerce")
                               .fillna(0.0).sum())
            types |= set(tp[paid])
            dates += list(pd.to_datetime(P[paid].get("Fecha de pago / Emisión de NC"),
                                         errors="coerce").dropna())
            if "Número de nota de crédito" in P.columns:
                nc_ids = set(P.loc[nota, "Número de nota de crédito"].astype(str))
        applied = 0.0
        need = {"Parte del pago aplicado a cargos", "Fecha del cargo"}
        if len(D) and need <= set(D.columns):
            fc = pd.to_datetime(D["Fecha del cargo"], errors="coerce")
            part = pd.to_numeric(D["Parte del pago aplicado a cargos"], errors="coerce").fillna(0.0)
            mine = (fc >= p0) & (fc <= p1)
            if "Número del pago" in D.columns:
                mine &= ~D["Número del pago"].astype(str).isin(nc_ids)
            applied = float(part[mine].sum())
        if not c["path"]:
            has_bill = _cycle_file("Reporte_Facturacion_MercadoLibre_*.xlsx", folder, y, mth)
            why.append("%s 账单（%s）%s" % (
                label, span, "还没有 Reporte_Pagos_Facturas —— 账单尚未关账或尚未扣款，平台还没出付款报表"
                if has_bill else "的 Reporte_Pagos_Facturas 没有下载"))
        elif not paid_total:
            why.append("%s 账单（%s）还没有付款记录（账单多半尚未关账或尚未扣款）"
                       % (label, span))
        elif not (len(D) and need <= set(D.columns)):
            why.append("%s 账单的付款明细页（Detalle de Pagos del mes）读不出来，"
                       "无法把付款拆到计费日" % label)
        in_month += applied
        info.append({"label": label, "span": span, "path": c["path"],
                     "payable": no + rev, "paid": paid_total, "applied_month": applied})
    S["pagos"] = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    paths = [c["path"] for c in cyc if c["path"]]
    S["files"]["pagos"] = " + ".join(paths) if paths else None
    m["pay_auto"] = in_month
    m["pay_cycles"] = info
    m["pay_complete"] = not why
    m["pay_missing"] = why
    m["pay_types"] = sorted(types)
    m["pay_date"] = max(dates).strftime("%Y-%m-%d") if dates else ""
    m["pay_nc"] = -float(FA[FA["_doc"] == "NC"]["amount"].sum())


def derive(store, folder, p0, p1):
    """读一家店一个月的全部原始文件，算出报表需要的每一个数。"""
    # load_errors：读不出来的**可选**文件 [(文件名, 错误)]，validate() 把每一条
    # 记成 ⑩ 页的 load:<文件名> 校验（见 loading.unreadable）。Ventas 与账单不在
    # 其列 —— 它们读不出来下面照旧抛异常，报表没有它们就不存在。
    S = {"store": store, "folder": folder, "files": {}, "anomalies": [],
         "p0": p0, "p1": p1, "meta": load_run_meta(folder), "load_errors": []}
    LE = S["load_errors"]

    vpath, V = load_ventas(folder)
    S["files"]["ventas"] = vpath
    S["ventas_warning"] = V.attrs.get("incomplete_notice", "")
    S["ventas_all"] = V
    A = V[(V["fecha"] >= p0) & (V["fecha"] <= p1 + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))].copy()
    S["ventas"] = A

    B = load_billing(folder)
    S["billing_all"] = B
    # 会计月归属规则（两国相同）：账单行按它自己的计费日 Fecha del cargo 归入
    # 自然月，**不按账单文件名**。所有账单文件先合并再筛，所以：
    #   墨西哥：账单 = 自然月，9 月只取到 Septiembre 那一张；
    #   智利：账单是 8 号到次月 7 号，9 月取到的是 Septiembre 账单里 9/1–9/7 的行
    #         ＋ Octubre 账单里 9/8–9/30 的行（见 bill_cycles）。
    # 订单级费用另走 billing_linked（按订单号跨全部账单关联，不看计费日）。
    FA = B[(B["cdate"] >= p0) & (B["cdate"] <= p1)].copy()
    # 墨西哥是自然月账单，不推；智利按店从账单/仓储文件推切期日。
    S["cycle_day"] = 1 if site.BILL_CYCLE_DAY == 1 else infer_cycle_day(B, folder)
    S["bill_cycles"] = bill_cycles(p0.year, p0.month, S["cycle_day"])
    S["bill_open"] = open_bills(folder, S["bill_cycles"]) if S["cycle_day"] != 1 else []
    S["billing"] = FA

    ids = set(A["oid"]) | set(A["pid"])
    ids.discard(""); ids.discard("nan")
    L = B[B["k_sale"].isin(ids) | B["k_pack"].isin(ids)].copy()   # 跨账期，8月订单的费用可能落在9月账单
    S["billing_linked"] = L

    def bucket_sum(frame, bucket, reversal=None):
        # astype(bool) 不能省：frame 为空时 .map() 返回的是 object dtype 的空
        # Series，pandas 会把它当成**列名列表**而不是布尔掩码，索引出一个零列的
        # 表，下一行取 f["is_rev"] 就 KeyError。做历史月份时（当月账单里没有
        # 那个月的订单）必然走到这条路径。
        f = frame[frame["parent_detalle"].map(
            lambda d: FEE_BY_ES.get(str(d), (None, None, None))[1] == bucket).astype(bool)]
        if reversal is True:
            f = f[f["is_rev"]]
        elif reversal is False:
            f = f[~f["is_rev"]]
        return float(f["amount"].sum())

    m = {}
    vf = V["fecha"].dropna()
    m["sales_from"] = vf.min() if len(vf) else pd.NaT
    m["sales_to"] = vf.max() if len(vf) else pd.NaT
    m["rows"] = int(len(A))
    m["orders_unique"] = int(A["oid"].nunique())
    m["units"] = float(num(A, V_UNITS).sum())
    m["ing"] = float(num(A, site.V_REVENUE).sum())
    m["env_ing"] = float(num(A, site.V_SHIP_INCOME).sum())
    if site.BONIF_COL:
        # 秘鲁：运费收入里并进了 Flex 配送补贴（见 load_ventas），拆开记一笔供说明
        m["bonif"] = float(num(A, "bonif").sum())
        m["bonif_n"] = int((num(A, "bonif") > site.ZERO_TOL).sum())
        m["env_ing_buyer"] = m["env_ing"] - m["bonif"]
    m["com_imp"] = float(-num(A, site.V_COMMISSION).sum())      # 佣金 + 代扣税（Ventas 口径）
    m["env_cost"] = float(-num(A, site.V_SHIP_COST).sum())
    m["refund"] = float(-num(A, site.V_REFUND).sum())
    m["total"] = float(num(A, site.V_TOTAL).sum())
    m["coupon"] = float(A["coupon"].sum())
    m["coupon_n"] = int((A["coupon"].abs() > site.COUPON_TOL).sum())
    denoms = A.loc[A["coupon"].abs() > site.COUPON_TOL, "coupon"].round(2).value_counts()
    m["coupon_denoms"] = {float(k): int(v) for k, v in denoms.items()}

    m["com_gross"] = bucket_sum(L, "commission", reversal=False)
    m["com_anul"] = bucket_sum(L, "commission", reversal=True)
    m["env_gross"] = bucket_sum(L, "shipping", reversal=False)
    m["env_anul"] = bucket_sum(L, "shipping", reversal=True)
    m["dev_gross"] = bucket_sum(L, "returns", reversal=False)
    m["dev_anul"] = bucket_sum(L, "returns", reversal=True)
    m["dev_net"] = m["dev_gross"] + m["dev_anul"]

    # ★ 代扣代缴税金：Ventas 的"佣金及税金" − 账单的纯佣金（毛额）
    m["tax_ret"] = m["com_imp"] - m["com_gross"]
    m["tax_rate_gross"] = (m["tax_ret"] / m["ing"]) if m["ing"] else 0.0
    m["iva_in_gmv"] = m["ing"] * site.IVA_RATE / (1 + site.IVA_RATE)
    m["net_revenue_exvat"] = m["ing"] - m["iva_in_gmv"]
    m["tax_vs_iva"] = (m["tax_ret"] / m["iva_in_gmv"]) if m["iva_in_gmv"] else 0.0
    if not site.WITHHOLDING:
        # 智利没有代扣：上面的 tax_ret 只是"Ventas 佣金列 − 账单纯佣金"这个
        # 算术差（各座桥要靠它勾稽，所以键名不变），它不是税。拆成两块：
        #   账单里关联到本期订单的退货处理费（Ventas 把它并进了佣金列）
        #   ＋ 剩余（退货处理费尚未出账单、包裹母子行错位等时间/关联差）
        m["vcol_dev"] = m["dev_net"]
        m["vcol_rest"] = m["tax_ret"] - m["dev_net"]
        m.update(commission_column_match(A, L))

    # 桥 A 残差：等式成立时应为 0
    m["resid"] = m["total"] - (m["ing"] + m["env_ing"] - m["com_imp"]
                               - m["env_cost"] - m["refund"] - m["coupon"])

    # ---- 结算通道 ----
    m["bill_total"] = float(FA["amount"].sum())
    m["bill_by_channel"] = {c: float(FA[FA["channel"] == c]["amount"].sum())
                            for c in (CH_SI, CH_NO, CH_NA)}
    m["bill_cnt_by_channel"] = {c: int((FA["channel"] == c).sum())
                                for c in (CH_SI, CH_NO, CH_NA)}
    # 应付 = No 桶的费用 ＋ 同期冲销掉 No 桶费用的那些行
    rev_of_no = FA[FA["is_rev"] & (FA["parent_channel"] == CH_NO)]
    if not len(rev_of_no):   # ID 关联不上时退回按科目性质判断
        rev_of_no = FA[FA["is_rev"] & FA["parent_detalle"].map(
            lambda d: FEE_BY_ES.get(str(d), (None, None, None))[2] is False).astype(bool)]
    m["invoice_internal_reversal"] = float(rev_of_no["amount"].sum())
    m["payable"] = m["bill_by_channel"][CH_NO] + m["invoice_internal_reversal"]

    # ---- 账单级（与任何订单都不挂钩）的费用 ----
    other = FA[FA["sdate"].isna()].copy()
    S["billing_other"] = other
    m["other_total"] = float(other["amount"].sum())
    # 按科目净额（费用 + 其冲销）
    per_fee = {}
    for es, zh, bucket, order_level in FEE_CATALOG:
        if order_level:
            continue
        v = float(other[other["parent_detalle"].astype(str) == es]["amount"].sum())
        per_fee[es] = v
    m["bill_fee_net"] = per_fee
    named_sum = sum(per_fee.values())
    # 兜底行：保证"月度账单费用小计"恒等于 other_total，任何意外都看得见而不是消失
    m["unc_bill"] = m["other_total"] - named_sum
    # 订单级的未登记科目
    unc_order_rows = FA[(~FA["known"]) & FA["sdate"].notna()]
    m["unc_order"] = float(unc_order_rows["amount"].sum())
    m["order_linked_billed"] = m["payable"] - m["other_total"]

    # ---- 未登记科目清单（供 ⑨ 页与校验使用）----
    unk = FA[~FA["known"]]
    S["unknown_fees"] = []
    if len(unk):
        g = unk.groupby(unk[F_DETALLE].astype(str))["amount"].agg(["count", "sum"])
        for name, row in g.sort_values("sum", key=abs, ascending=False).iterrows():
            chans = sorted(set(unk[unk[F_DETALLE].astype(str) == name]["channel"]))
            S["unknown_fees"].append({"detalle": name, "count": int(row["count"]),
                                      "amount": float(row["sum"]),
                                      "channels": ", ".join(chans)})
    # 已登记但出现在意外通道的科目（信息性）
    for es, zh, bucket, order_level in FEE_CATALOG:
        sub = FA[(FA[F_DETALLE].astype(str) == es)]
        if not len(sub):
            continue
        has_sale = sub["sdate"].notna()
        if order_level and not has_sale.any():
            S["anomalies"].append("科目「%s」按字典应为订单级，但本月全部行都没有销售日期" % es)
        if (not order_level) and has_sale.any():
            S["anomalies"].append("科目「%s」按字典应为账单级，但本月有 %d 行带销售日期"
                                  % (es, int(has_sale.sum())))

    # ---- 账单支付 ----
    if S["cycle_day"] != 1:
        # 智利：会计月横跨两张账单，实付按 Pagos 明细页逐笔拆到计费日（见
        # load_pagos_cycles）。墨西哥一张账单就是一个月，走下面原来的路径。
        _derive_pagos_cycles(S, m, folder, B, FA, p0, p1)
    else:
        ppath, P = load_pagos(folder, p0.year, p0.month, LE)
        S["files"]["pagos"] = ppath
        S["pagos"] = P
        if len(P) and "Tipo de pago" in P.columns:
            tp = P["Tipo de pago"].astype(str)
            # 实际付款 = 除贷记单以外的所有支付行。以前只认类型里带 "Cobro" 的行，
            # 而平台实际写的是 "Pago con débito automático"（自动扣款）或 "Pago
            # voluntario"（主动付款）—— 于是实扣永远是 0，桥B 每家店每个月都不过。
            # 2026-10-14 用 11 家店 9 月的真实数据核对：9 家的付款行金额与应付分毫
            # 不差；另两家的差额是真实的（TOOL_TA03 少付 5,181.17，UNIT_TA04 多付
            # 1.90）。贷记单另算：它冲的是订单级费用，不是这张账单的付款。
            nota = tp.str.contains("Nota", na=False)
            paid = ~nota & tp.str.strip().ne("") & tp.ne("nan")
            m["pay_auto"] = float(P[paid]["Importe total"].sum())
            m["pay_nc"] = float(P[nota]["Importe total"].sum())
            m["pay_types"] = sorted(set(tp[paid]))
            pd_col = "Fecha de pago / Emisión de NC"
            dts = pd.to_datetime(P[paid].get(pd_col), errors="coerce").dropna()
            m["pay_date"] = dts.max().strftime("%Y-%m-%d") if len(dts) else ""
        else:
            m["pay_auto"] = m["pay_nc"] = 0.0
            m["pay_date"] = ""

    # ---- Full 费用交叉核对 ----
    if S["cycle_day"] != 1:
        # 智利：两期 Cargos Full 文件里计费日落在本月的行（见 load_cargos_full_cycles）。
        # 缺任何一期就无法核对整月，按"没有文件"跳过并说明缺哪一期。
        cpaths, csheets, ctotal, cmiss = load_cargos_full_cycles(
            folder, S["bill_cycles"], p0, p1, LE)
        cpath = " + ".join(cpaths) if cpaths and not cmiss else None
        S["cargos_full_missing"] = cmiss
    else:
        cpath, csheets, ctotal = load_cargos_full(folder, p0.year, p0.month, LE)
    S["files"]["cargos_full"] = cpath
    S["cargos_full_sheets"] = csheets
    m["cargos_full_report_total"] = ctotal
    full_buckets = ("storage", "fulfillment", "penalty")
    m["full_bill_total"] = float(sum(
        v for es, v in per_fee.items()
        if FEE_BY_ES[es][1] in full_buckets))
    m["ads_net"] = float(sum(v for es, v in per_fee.items() if FEE_BY_ES[es][1] == "ads"))
    m["subscription_net"] = float(sum(v for es, v in per_fee.items()
                                      if FEE_BY_ES[es][1] == "subscription"))
    m["misc_net"] = float(sum(v for es, v in per_fee.items()
                              if FEE_BY_ES[es][1] == "misc"))

    # ---- 仓储费报表 ----
    if S["cycle_day"] != 1:
        # 智利：仓储费报表也按账单切期日切期，只取占用日落在本月的那几天，并拿账单里
        # **同样那几天**（计费日晚一天，见 STORAGE_BILL_LAG_DAYS）的仓储费来比。
        spaths, accrual, srange, sdays = load_storage_daily(folder, p0, p1, LE)
        spath = " + ".join(spaths) if spaths else None
        m["storage_days"] = sdays
        m["storage_month_days"] = int((p1 - p0).days) + 1
        st = FA[FA["parent_detalle"].astype(str) == "Cargo por servicio de almacenamiento Full"]
        if srange:
            lag = pd.Timedelta(days=STORAGE_BILL_LAG_DAYS)
            b0, b1 = srange[0] + lag, srange[1] + lag
            stb = B[(B["parent_detalle"].astype(str) == "Cargo por servicio de almacenamiento Full")
                    & (B["cdate"] >= b0) & (B["cdate"] <= b1)]
            m["storage_bill_cmp"] = float(stb["amount"].sum())
        else:
            m["storage_bill_cmp"] = float(st["amount"].sum())
    else:
        spath, accrual, srange = load_storage(folder, p0, p1, LE)
    S["files"]["storage"] = spath
    m["storage_accrual"] = accrual
    m["storage_range"] = srange
    m["storage_bill"] = per_fee.get("Cargo por servicio de almacenamiento Full", 0.0)

    # ---- 退货质检 ----
    rpath, R = load_returns(folder, LE)
    S["files"]["returns"] = rpath
    S["returns_all"] = R
    if len(R):
        S["returns"] = R[(R["fr"] >= p0) & (R["fr"] <= p1)].copy()
    else:
        S["returns"] = pd.DataFrame()

    # ---- Full 仓库存快照 ----
    spath, snap, ST = load_stock(folder, LE)
    S["files"]["stock"] = spath
    S["stock"] = ST
    S["stock_snap_text"] = snap
    S["stock_snap"] = stock_snapshot_date(snap, folder)

    # ---- MercadoPago 结算流水 ----
    mpath, MP = load_settlement(folder, p0, p1, LE)
    S["files"]["settlement"] = mpath
    S["mp"] = MP
    try:
        S["mp_orders"] = mp_by_order(A, S["billing_all"], folder, LE)
    except Exception as e:                          # 验证本身坏了不能拖垮报表
        S["mp_orders"] = None
        S["anomalies"].append("MercadoPago 逐单验证没能完成：%s" % str(e)[:120])
    if not site.WITHHOLDING and S.get("mp_orders"):
        # 智利：mp_by_order 原样运行（不改口径），只在结果上标注一类已查明的差额。
        # 买家另付的运费（Ingresos por envío）在 MercadoPago 里是一笔单独的
        # SETTLEMENT_SHIPPING，收款号和商品那笔不同，账单的 Pago 列只指向商品那笔
        # —— 所以这些订单"平台少到账"的正好是运费收入。DAB_CL 2026-09：250 单
        # 不一致里 216 单的差额分毫不差地等于该单的运费收入；Tecnohub_CL 3/3。
        env = dict(zip(A["oid"], num(A, site.V_SHIP_INCOME)))
        ship = [x for x in S["mp_orders"]["differ"]
                if env.get(x[0], 0.0) > site.CN_TOL and abs((x[1] - x[2]) - env[x[0]]) <= site.CN_TOL]
        S["mp_orders"]["ship_explained"] = (len(ship), float(sum(x[1] - x[2] for x in ship)))
        S["mp_orders"]["ship_ids"] = set(x[0] for x in ship)

    # ---- 组合单 / 多件单 / 换货单 ----
    m["pack_flag_yes"] = int((A[V_PACK_FLAG] == "Sí").sum()) if V_PACK_FLAG in A.columns else 0
    m["multi_unit_rows"] = int((num(A, V_UNITS) > 1).sum())
    nsku = A.groupby("oid")[V_SKU].nunique() if V_SKU in A.columns else pd.Series(dtype=int)
    m["multi_sku_orders"] = int((nsku > 1).sum())
    nsku_all = V.groupby("oid")[V_SKU].nunique() if V_SKU in V.columns else pd.Series(dtype=int)
    m["multi_sku_orders_all"] = int((nsku_all > 1).sum())
    m["blank_total_rows"] = int(pd.to_numeric(A[site.V_TOTAL], errors="coerce").isna().sum()) if site.V_TOTAL in A.columns else 0
    m["abnormal_orders"] = int(A["abnormal"].sum())

    # ---- 售后与库存去向 ----
    st = A[V_STATUS].astype(str) if V_STATUS in A.columns else pd.Series("", index=A.index)
    m["n_back_on_sale"] = int(st.isin(STATES_BACK_ON_SALE).sum())
    m["n_stock_lost"] = int(st.isin(STATES_STOCK_LOST).sum())
    m["n_seller_keeps"] = int(st.isin(STATES_SELLER_KEEPS).sum())
    m["amt_stock_lost"] = float(num(A[st.isin(STATES_STOCK_LOST)], site.V_REVENUE).sum())

    S["m"] = m

    # ---- 贷记明细 × 销售报表 逐单交叉比对 ----
    S["cn"] = credit_note_audit(S, folder)

    return S
