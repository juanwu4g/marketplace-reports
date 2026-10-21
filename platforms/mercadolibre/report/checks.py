# -*- coding: utf-8 -*-
"""MercadoLibre report — validate(): the check records (core.report_kit.check) of
every store, and check_hints() — where to start looking when one fails."""

import glob
import os

import pandas as pd

from core import report_kit

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.site import localize
from platforms.mercadolibre.report.catalog import (MONTH_FULL_ES, SHEET_NAMES, TAX_RATE_MAX,
    TAX_RATE_MIN, file_month_token, fmt, pct)
from platforms.mercadolibre.report.loading import cycle_cn, cycle_day, sales_window_start


# ════════════════════════════════════════════════════════════════════════
# 校验层
# ════════════════════════════════════════════════════════════════════════

_chk = report_kit.check


def check_hints(c, S):
    """一条未通过的校验该从哪儿查起。返回若干行文字；没有可说的返回 []。

    只是给人看的排查线索，不参与任何计算。S 是这条校验所属店铺的数据，被剔除
    的店铺（store_usable）没有，传 None。
    """
    key, files = c["key"], (S or {}).get("files", {})
    if key == "bridge_b" and S is not None:
        P = S["pagos"]
        types = sorted(set(P["Tipo de pago"].dropna().astype(str))) if "Tipo de pago" in P else []
        return ["支付源文件：%s；行数 %d；支付类型：%s" % (
                    files.get("pagos"), len(P), "、".join(types) or "（无可识别支付类型）"),
                "排查提示：实付 = 除贷记单以外的所有支付行（自动扣款、主动付款）。差额为正"
                "多半是账单还没付清（或付款落在下载之后），为负是多付；对照平台账单页的付款记录。"]
    if key == "mp_settlement":
        return ["结算源文件：%s" % files.get("settlement"),
                "排查提示：逐单比的是销售报表 Total 和该订单收款在 MercadoPago 的净到账（不限日期）。"
                "一次买几件的子行（金额为 0）已跳过，钱记在包裹行上。剩下的差额原因还没有逐类查明，"
                "请按订单号在卖家后台核对退款、争议和运费；清单见【%s】。" % SHEET_NAMES[9]]
    if key == "no_withholding":
        return ["排查提示：看③桥 E。剩余差额变大时，先按订单号找销售报表佣金列与账单 Cargo por venta "
                "不相等的订单（退货、包裹以外的）；MercadoPago 流水 TAXES_AMOUNT 不为 0 说明平台开始代扣，"
                "智利报表的结构需要重新设计。"]
    if key in ("billing_present", "sales_window", "tax_rate"):
        return ["排查提示：先核对本期账单及销售窗口是否完整；税率异常可能由缺账单、关联失败或税务状态变化引起。"]
    if key.startswith("cn_"):
        return ["排查提示：见报表⑩逐单异常清单；核对订单/包裹关联、冲销及现金凭证，推定调整不等于现金已证实。"]
    return []


def _billing_present_cycles(S):
    """智利版"本期账单已下载"：会计月横跨的两张账单都得在。

    只有一张时费用行会静默缺一段（缺 Octubre 就少了 9/8–9/30 的广告和仓储），
    行数却不是 0，墨西哥那条"有没有行"的判断看不出来。按账单文件名判断。
    """
    name, p0 = S["store"], S["p0"]
    have = set()
    for f in glob.glob(os.path.join(S["folder"], "Reporte_Facturacion_MercadoLibre_*.xlsx")):
        tok = file_month_token(os.path.basename(f))
        if tok:
            have.add(tok)
    need = [(y, mth, c0, c1) for y, mth, c0, c1 in S["bill_cycles"]]
    miss = ["%s %d（%s–%s 的费用）" % (MONTH_FULL_ES[mth], y, c0.strftime("%m/%d"),
                                       c1.strftime("%m/%d"))
            for y, mth, c0, c1 in need if (y, mth) not in have]
    n_bill = len(S["billing"])
    labels = "＋".join("%s %d（%s–%s）" % (MONTH_FULL_ES[mth], y, c0.strftime("%m/%d"),
                                          c1.strftime("%m/%d")) for y, mth, c0, c1 in need)
    return _chk(
        "billing_present", name, "本期账单已下载",
        n_bill > 0 and not miss,
        ("本期账单 %d 行，计费日落在 %s 内；本店账单按 %s 切期，本月取自 %s 两张账单%s"
         % (n_bill, p0.strftime("%Y-%m"), cycle_cn(cycle_day(S)), labels,
            ("。⚠ %s 账单尚未关账（平台注明“podría tener cambios hasta el cierre”），"
             "本月落在它里面的费用以后还可能变，关账后建议重出" % "、".join(S["bill_open"]))
            if S.get("bill_open") else "")
         if not miss else
         "⚠ 本店的账单是 %s，%s 这个会计月要用 %s 两张账单，"
         "目录里缺：%s。缺的那段计费日上的广告、仓储等费用全部缺失，桥 B 和 Full "
         "核对也会跟着失真。等账单关账后用 run_batch.py --month %s "
         "重新下载再重出。" % (cycle_cn(cycle_day(S)), p0.strftime("%Y-%m"), labels,
                             "；".join(miss), p0.strftime("%Y-%m"))))


def _bridge_b_cycles(S):
    """智利版桥 B：会计月内的应付 vs 两张账单里付给本月费用的那部分。"""
    m, name = S["m"], S["store"]
    per = "；".join("%s 账单（%s）整张应付 %s、实付 %s%s" % (
        c["label"], c["span"], fmt(c["payable"]), fmt(c["paid"]),
        "" if not c["paid"] else (" ✓" if abs(c["payable"] - c["paid"]) <= site.TOL_CENT else " ✗"))
        for c in m.get("pay_cycles", []))
    if not m.get("pay_complete"):
        return _chk("bridge_b", name, "桥B 账单应付 vs 实际扣款", True,
                    "本月横跨两张账单，付款记录不全，无法核对整月：%s。（%s）"
                    % ("；".join(m.get("pay_missing", [])), per), skipped=True)
    d = m["payable"] - m["pay_auto"]
    return _chk("bridge_b", name, "桥B 账单应付 vs 实际扣款",
                abs(d) <= site.TOL_CENT,
                "计费日在本月的应付 %s，两张账单的付款里付给这些费用的部分合计 %s，差 %s。"
                "（%s）" % (fmt(m["payable"]), fmt(m["pay_auto"]), fmt(d), per))


# 智利"佣金列里不是佣金的部分"扣掉退货处理费后，剩余占 GMV 的上限。实测 DAB_CL
# 2026-09 是 −0.005%（时间差：退货处理费已进 Ventas、还没出账单等）。若平台开始
# 代扣（墨西哥是 ~9%），这个比例会跳到几个百分点，远超此线。
CL_VCOL_REST_MAX = 0.005


def _no_withholding_check(S):
    """智利：确认"平台不代扣税"这个前提在本期数据里仍然成立。

    三条证据，任一不成立就失败（error 级 —— 前提一变，损益表的结构就不对了）：
      ① MercadoPago 结算流水的 TAXES_AMOUNT 合计为 0（有流水时）；
      ② Ventas 佣金列 − 账单纯佣金 − 退货处理费，剩余不超过 GMV 的 0.5%；
      ③ （说明用）佣金列与账单纯佣金逐单相等的订单比例。
    """
    m, name = S["m"], S["store"]
    mp = S.get("mp") or {}
    taxes = mp.get("tax") if mp else None
    rest = m.get("vcol_rest", 0.0)
    rate = (rest / m["ing"]) if m["ing"] else 0.0
    ok = abs(rate) <= CL_VCOL_REST_MAX and (taxes is None or abs(taxes) <= site.TOL_CENT)
    n, of = m.get("com_match_n", 0), m.get("com_match_of", 0)
    return _chk(
        "no_withholding", name, "平台未代扣税（智利）",
        ok,
        "Ventas 佣金列比账单纯佣金多 %s，其中退货处理费 %s，剩余 %s（占 GMV %s，上限 %s）；"
        "佣金列与账单佣金逐单相等 %d/%d 单；MercadoPago 流水 TAXES_AMOUNT %s。%s"
        % (fmt(m["tax_ret"]), fmt(m.get("vcol_dev", 0.0)), fmt(rest), pct(rate, 3),
           pct(CL_VCOL_REST_MAX, 1), n, of,
           "无流水，未核" if taxes is None else "合计 %s" % fmt(taxes),
           "前提成立：智利平台不代扣 IVA/所得税，报表不设代扣税行。" if ok else
           "⚠ 前提不成立：平台可能开始代扣税金，或佣金列混进了新的扣款。"
           "智利报表按“不代扣”设计，必须先查明再用本报表。"))


def validate(stores, tol=None):
    if tol is None:
        tol = site.BRIDGE_A_TOL       # 墨西哥 0.05；智利按 1 比索
    out = []
    for S in stores:
        m, name = S["m"], S["store"]
        if S.get("ventas_warning"):
            out.append(_chk("ventas_incomplete", name, "Ventas 导出数据不完整", False,
                            "平台提示「%s」；已用现有数据生成报表，收入和利润可能偏低。"
                            "请重新下载核对。源文件：%s"
                            % (S["ventas_warning"], os.path.basename(S["files"]["ventas"])),
                            severity="warn"))

        out.append(_chk("bridge_a", name, "桥A 逐笔订单勾稽",
                        abs(m["resid"]) <= tol,
                        "残差 %s（阈值 %s）。不为零说明还有未识别的隐藏费用或折扣。"
                        % (fmt(m["resid"]), fmt(tol))))

        if cycle_day(S) != 1:
            out.append(_bridge_b_cycles(S))
        elif S["files"].get("pagos"):
            d = m["payable"] - m["pay_auto"]
            out.append(_chk("bridge_b", name, "桥B 账单应付 vs 实际扣款",
                            abs(d) <= site.TOL_CENT,
                            "应付 %s，实扣 %s，差 %s" % (fmt(m["payable"]), fmt(m["pay_auto"]), fmt(d))))
        else:
            out.append(_chk("bridge_b", name, "桥B 账单应付 vs 实际扣款", True,
                            "本期没有 Reporte_Pagos_Facturas 文件，跳过", skipped=True))

        unk = S["unknown_fees"]
        out.append(_chk("fee_catalog", name, "费用科目全部已登记",
                        not unk,
                        "全部科目已在 FEE_CATALOG 登记" if not unk else
                        "发现 %d 个未登记科目，合计 %s：%s" % (
                            len(unk), fmt(sum(u["amount"] for u in unk)),
                            "；".join("%s(%s)" % (u["detalle"], fmt(u["amount"])) for u in unk))))

        out.append(_chk("bill_block_balance", name, "账单费用小计自平",
                        abs(m["unc_bill"]) <= site.TOL_CENT,
                        "兜底行金额 %s。不为零说明有科目未在 FEE_CATALOG 里、或已登记科目出现在非预期位置。"
                        % fmt(m["unc_bill"]), severity="warn"))

        if site.WITHHOLDING:
            r = m["tax_rate_gross"]
            out.append(_chk("tax_rate", name, "代扣税率在预期区间",
                            TAX_RATE_MIN <= r <= TAX_RATE_MAX,
                            "实测 %s（预期 %s–%s）。越界说明税制或账号税务状态变了，需重新确认。"
                            % (pct(r, 3), pct(TAX_RATE_MIN, 1), pct(TAX_RATE_MAX, 1))))
        else:
            out.append(_no_withholding_check(S))

        if S["files"].get("cargos_full"):
            d = m["cargos_full_report_total"] - m["full_bill_total"]
            out.append(_chk("cargos_full", name, "Full 费用报表 vs 账单",
                            abs(d) <= site.TOL_CENT,
                            "Full 报表 %s，账单口径 %s，差 %s（两者必须相等，相加即重复计数）"
                            % (fmt(m["cargos_full_report_total"]), fmt(m["full_bill_total"]), fmt(d))))
        else:
            out.append(_chk("cargos_full", name, "Full 费用报表 vs 账单", True,
                            ("%s没有 Full 仓，平台不提供 Full 费用报表，这一项不适用"
                             % site.COUNTRY_CN) if not site.HAS_FULL else
                            ("缺 %s 的 Reporte_Cargos_Full（本月横跨两期账单，两期都要有），跳过"
                             % "、".join(S["cargos_full_missing"]))
                            if S.get("cargos_full_missing") else
                            "本期没有 Reporte_Cargos_Full 文件，跳过", skipped=True))

        MO = S.get("mp_orders")
        if MO and MO["linked"]:
            # 提醒级：个别订单对不上是要人看的事，不说明报表口径错了。
            k = len(MO["differ"])
            out.append(_chk("mp_settlement", name, "MercadoPago 逐单交叉验证",
                            k == 0,
                            "%d 单能对上收款：%d 单分毫不差%s%s。流水覆盖 %s"
                            % (MO["linked"], MO["agree"],
                               "，%d 单不一致（合计差 %s，见 ⑩ 页）"
                               % (k, fmt(MO["sum_mp"] - MO["sum_v"])) if k else "",
                               "；另有 %d 单对不上收款号" % (MO["unlinked"] + MO["no_cash"])
                               if MO["unlinked"] + MO["no_cash"] else "",
                               MO["span"])
                            + (("。其中 %d 单（合计 %s）的差额正好等于该单的运费收入 —— 秘鲁买家另付的运费"
                                "（SETTLEMENT_SHIPPING）和 Flex 配送补贴（CASHBACK）在 MercadoPago 都是"
                                "单独一笔，收款号不同、账单的 Pago 列关联不到，不是少付；其余 %d 单需人工核对"
                                if site.BONIF_COL else
                                "。其中 %d 单（合计 %s）的差额正好等于买家另付的运费 —— 智利运费在 "
                                "MercadoPago 是单独一笔 SETTLEMENT_SHIPPING，收款号不同、账单的 Pago "
                                "列关联不到，不是少付；其余 %d 单需人工核对")
                               % (MO["ship_explained"][0], fmt(MO["ship_explained"][1]),
                                  k - MO["ship_explained"][0])
                               if MO.get("ship_explained") and MO["ship_explained"][0] else ""),
                            severity="warn"))
        elif MO is not None:
            out.append(_chk("mp_settlement", name, "MercadoPago 逐单交叉验证", True,
                            "有结算流水，但没有一个订单能对上收款号（账单里缺 Pago 列？），跳过",
                            skipped=True))
        else:
            out.append(_chk("mp_settlement", name, "MercadoPago 流水交叉验证", True,
                            "本店没有 settlement 流水（多为该登录账号无报表权限），跳过",
                            severity="warn", skipped=True))

        # ---- 本期账单是否存在 ----
        # 放在最前面：账单整块缺失会让后面一串校验一起失真（代扣税率会变成
        # 0% 或几十个百分点、桥B 和 Full 核对直接跳过），但那些都是**症状**。
        # 没有这一条，人只能从症状倒推病因。
        n_bill = len(S["billing"])
        if cycle_day(S) != 1:
            out.append(_billing_present_cycles(S))
        else:
            out.append(_chk(
              "billing_present", name, "本期账单已下载",
              n_bill > 0,
              ("本期账单 %d 行，计费日落在 %s 内" % (n_bill, S["p0"].strftime("%Y-%m"))
               if n_bill else
               "⚠ 这个下载目录里**没有 %s 的账单**（目录里的账单是 %s）。"
               "平台费用、佣金、代扣税全部缺失，后面几条校验的异常都是这一条的连带结果。"
               "用 run_batch.py --month %s 重新下载该月账单再重出。"
               % (S["p0"].strftime("%Y-%m"),
                  "、".join(sorted({d.strftime("%Y-%m") for d in S["billing_all"]["cdate"]
                                    if pd.notna(d)})) or "（一份都没有）",
                  S["p0"].strftime("%Y-%m")))))

        # ---- 销售报表是否覆盖整个会计月 ----
        # 销售报表是**滚动窗口**导出：窗口起点晚于月初，这个月就是被截断的。
        # 桥A 仍然平（它只在导出内部勾稽），所以不单独查就看不出来 ——
        # 实测 EWTTO_SM 7 月被截到 07-13，GMV 少了 36% 而全部校验照样通过。
        # 判断依据是**导出时请求的窗口**，不是最早一笔订单的日期 —— 两者
        # 完全不同：窗口起点晚于月初才是数据被截断；最早订单晚于月初也可能
        # 只是这家店当月才开张（BOCINA_TA02 就是 7 月 16 号开的）。
        w0 = sales_window_start(S.get("meta") or {})
        if w0 is None:
            out.append(_chk(
                "sales_window", name, "销售数据覆盖整个会计月", True,
                "这个下载目录没有 run_meta.json（老目录），无法确认销售窗口是否"
                "覆盖整月，跳过。重新下载一次即可开始检查。",
                severity="warn", skipped=True))
        else:
            late = (w0 - S["p0"]).days
            f0 = m.get("sales_from")
            out.append(_chk(
                "sales_window", name, "销售数据覆盖整个会计月",
                late <= 0,
                ("销售窗口自 %s 起，覆盖了 %s 整月（最早一笔订单 %s）"
                 % (w0.strftime("%Y-%m-%d"), S["p0"].strftime("%Y-%m"),
                    f0.strftime("%Y-%m-%d") if f0 is not None and pd.notna(f0) else "—")
                 if late <= 0 else
                 "⚠ 销售窗口只回溯到 %s，比 %s 月初晚 %d 天 —— 本月前 %d 天的订单"
                 "全部缺失，收入、费用、SKU 全部偏低。用 run_batch.py --month %s "
                 "重新下载（会自动把窗口放大到覆盖整月）再重出。"
                 % (w0.strftime("%Y-%m-%d"), S["p0"].strftime("%Y-%m"), late,
                    late, S["p0"].strftime("%Y-%m")))))

        # ---- 贷记明细 × 销售报表 逐单交叉比对 ----
        # 这三项都是逐单的，落在其余校验的盲区里：桥A 只在销售报表内部勾稽，
        # 桥B 只比账单总额，MercadoPago 那条只比店铺汇总差额。
        c = S["cn"]
        out.append(_chk(
            "cn_orphan", name, "账单有费用但销售报表无此单",
            not c["orphan"],
            ("全部对得上，或虽查无此单但计费已被全额冲平（净额 0）"
             if not c["orphan"] else
             "%d 单在账单里计了费、销售报表里查无此单，且未被冲平，合计 %s。"
             "这笔钱不在销售报表里，桥A 查不到，需人工向平台核实。"
             % (len(c["orphan"]), fmt(c["orphan_total"])))))

        out.append(_chk(
            "cn_unreversed", name, "取消单的费用是否已冲销",
            not c["unrev"],
            ("取消单要么平台压根没计费，要么已全额冲销"
             if not c["unrev"] else
             "%d 单已取消、账单却计了费且没有冲销，合计 %s。"
             "（发货前取消平台通常不下账单，所以只统计确实计过费的。）"
             % (len(c["unrev"]), fmt(c["unrev_total"])))))

        n_cash = sum(1 for x in c["partial"] if x[12] == "现金凭证")
        out.append(_chk(
            "cn_partial", name, "部分退款订单重复扣佣金",
            not c["partial"] and not c["mismatch"],
            ("本期没有部分冲销的佣金"
             if not c["partial"] and not c["mismatch"] else
             "%d 单佣金被部分冲销，销售报表把原佣金重复扣了一次，合计 %s"
             "（其中 %d 单有 MercadoPago 现金凭证，%d 单按规律推定）。"
             "已在损益表“加：部分退款重复扣佣金调整”行修正。%s"
             % (len(c["partial"]), fmt(c["partial_total"]), n_cash,
                len(c["partial"]) - n_cash,
                ("另有 %d 单金额对不上已知模式，未计入调整，需人工核对。"
                 % len(c["mismatch"])) if c["mismatch"] else "")),
            severity="warn"))
    for c in out:
        c["name"], c["detail"] = localize(c["name"]), localize(c["detail"])
    return out
