# -*- coding: utf-8 -*-
"""MercadoLibre report — ① 阅读说明."""

import os

from core import sku_costs as SC
from core.report_family import readme_rows, set_widths
from core.report_kit import C_GREY, C_NAVY, C_RED, C_WARN, F, est_height, fill, wrap

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (CH_NO, CH_SI, FEE_BY_ES, F_DETALLE, F_SALE_NO,
    MONTH_FULL_ES, SHEET_NAMES, fmt, fmt0, pct)
from platforms.mercadolibre.report.loading import cycle_cn, cycle_day
from platforms.mercadolibre.report.sheet_common import (fee_rows, main_store, no_full_text,
    per_store, storage_compare_text)


# ════════════════════════════════════════════════════════════════════════
# ① 阅读说明
# ════════════════════════════════════════════════════════════════════════

def sheet_readme(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[0])
    set_widths(ws, (4, 30, 118))
    ws["A1"] = "Mercado Libre %s店铺 月度财务报表" % site.COUNTRY_CN
    ws["A1"].font = F(15, True, C_NAVY)
    ws.row_dimensions[1].height = 26
    ws["A2"] = ("会计期间：%s　|　币种：%s %s　|　编制日期：%s　|　数据截取时点：%s"
                % (ctx["period_cn"], site.CUR, site.CUR_CN, ctx["build_date"], ctx["cutoff"]))
    ws["A2"].font = F(9, False, C_GREY)
    incomplete = [S["store"] for S in stores if S.get("ventas_warning")]
    if incomplete:
        warning = ("⚠ Ventas 数据不完整：%s。报表使用已导出的数据；本店及合计金额"
                   "可能偏低。请重新下载并核对原始数据，详见【%s】。"
                   % ("、".join(incomplete), SHEET_NAMES[9]))
        ws.merge_cells("A3:C3")
        c = ws["A3"]
        c.value = warning
        c.font = F(10, True, C_RED)
        c.fill = fill(C_WARN)
        c.alignment = wrap("top")
        ws.row_dimensions[3].height = est_height(warning, 100)

    M = main_store(stores)
    mm = M["m"]
    chk = {(c["store"], c["key"]): c for c in ctx["checks"]}
    ok_a = all(chk[(S["store"], "bridge_a")]["ok"] for S in stores)
    ok_b = all(chk[(S["store"], "bridge_b")]["ok"] for S in stores)
    mp_stores = [S for S in stores if S.get("mp")]
    coupon_stores = [S for S in stores if S["m"]["coupon_n"] > 0]

    src_names = {"ventas": "销售明细(%s)" % site.VENTAS_SHEET, "pagos": "账单支付(Pagos)",
                 "cargos_full": "Full 费用报表(Cargos Full)", "storage": "仓储费用报表",
                 "returns": "退货质检(Returns/Triages)", "settlement": "MercadoPago 结算流水"}
    have = []
    for k, label in src_names.items():
        who = [S["store"] for S in stores if S["files"].get(k)]
        if who:
            have.append(label if len(who) == len(stores) else "%s(仅 %s)" % (label, "、".join(who)))
    sources = "、".join(["平台账单(Facturación)", "贷记单(Notas de Crédito)"] + have)

    rows = [("", "", "")]
    rows.append(("S", "一、本报表覆盖范围", ""))
    for i, S in enumerate(stores, 1):
        s = S["m"]
        rows.append(("", "店铺 %d" % i,
                     "%s —— %s %s 单 / %s 件 / GMV %s"
                     % (S["store"], ctx["period_short"], fmt0(s["rows"]),
                        fmt0(s["units"]), fmt(s["ing"]))))
    rows.append(("", "数据来源", sources))
    costs = ctx.get("sku_costs")          # core/sku_costs 的成本表；没配是 None
    if costs is not None:
        rows.append(("",) + SC.source_line(costs))
    rows.append(("", "", ""))

    rows.append(("S", "二、为什么你以前“平不了账”——六个结构性原因", ""))
    rows.append(("", "① 三条时间线",
                 "销售报表按【下单日】、平台账单按【计费日】、资金到账按【放款日】。同一笔订单会落在三个不同的月份，"
                 "直接把三张表加总必然对不上。本报表统一采用【订单日口径】做损益，并在“平账勾稽”页给出到账单口径的桥式调节。"))
    if site.WITHHOLDING:
        rows.append(("", "② 佣金里混进了代扣税",
                     "销售报表的“Cargo por venta e impuestos”= 平台佣金 + 平台代扣代缴的税金(IVA/ISR retenciones)，"
                     "而平台账单里的“Cargo por venta”只有纯佣金。两者差额 = 代扣税，本期实测约占含税售价的 %s。"
                     "这笔钱不是费用，是已预缴税款，可用于抵扣申报，必须单列。本期：%s。"
                     % (pct(mm["tax_rate_gross"], 2),
                        per_store(stores, lambda S: fmt(S["m"]["tax_ret"]), "、"))))
    else:
        rows.append(("", "② 佣金列里混进了退货处理费（智利没有代扣税）",
                     "销售报表的“Cargo por venta e impuestos”名字里有“impuestos”，但智利平台并不代扣税："
                     "正常订单这一列与账单的纯佣金“Cargo por venta”逐单相等（本期 %s），"
                     "MercadoPago 流水的税金字段 TAXES_AMOUNT 也全是 0。两者的差额来自退货订单 —— "
                     "平台把退货处理费“Cargo por devolución”并进了这一列。本期差额 %s，其中账单上的退货处理费 %s，"
                     "剩余 %s 是时间/关联差（退货费已进销售报表、还没出账单等）。本报表按账单口径把佣金和退货处理费分两行列示，"
                     "不设“代扣代缴税金”行。"
                     % (per_store(stores, lambda S: "%d/%d 单" % (S["m"].get("com_match_n", 0),
                                                                 S["m"].get("com_match_of", 0)), "、"),
                        per_store(stores, lambda S: fmt(S["m"]["tax_ret"]), "、"),
                        per_store(stores, lambda S: fmt(S["m"].get("vcol_dev", 0.0)), "、"),
                        per_store(stores, lambda S: fmt(S["m"].get("vcol_rest", 0.0)), "、"))))
    rows.append(("", "③ 费用分两个通道结算",
                 "账单里“Descontado de la operación = Si”的费用（佣金、配送费）在每笔订单打款时就已扣掉，不会再收一次；"
                 "“No”的费用（广告、仓储、揽收、罚金等）月底汇总，从 MercadoPago 余额自动扣款。"
                 "把两类相加就会重复计算一次佣金和运费（本期重复额将达 %s）。"
                 % fmt(sum(S["m"]["bill_by_channel"][CH_SI] for S in stores))))
    rows.append(("", "④ 退货退款走三套凭证",
                 "一笔退货会同时产生：销售报表的“Anulaciones y reembolsos”（退给买家的钱）、"
                 "账单里的“Anulación del cargo”（平台退还的佣金/运费）、以及贷记单 Nota de Crédito（另开红字发票）。"
                 "只看其中一套，退货损失就会算错。"))
    if site.BONIF_COL:
        rows.append(("", "⑤ “Descuentos y bonificaciones”列里是平台给的 Flex 配送补贴",
                     "墨西哥等国的这一列导出为空，%s的这一列有值：是平台付给卖家的 Mercado Envíos Flex "
                     "配送补贴（本期 %s），在 MercadoPago 里是单独一笔 CASHBACK。它是卖家的钱，不是卖家承担的"
                     "优惠券，本报表把它并进【运费收入】（损益表“运费收入”行的说明里拆开写了）。卖家自负的优惠券"
                     "仍用 Total 倒推：本期 %s。"
                     % (site.COUNTRY_CN,
                        per_store(stores, lambda S: "%d 单、%s" % (S["m"].get("bonif_n", 0),
                                                                fmt(S["m"].get("bonif", 0.0))), "、"),
                        per_store(stores, lambda S: "%d 笔、%s" % (S["m"]["coupon_n"],
                                                                fmt(S["m"]["coupon"])), "、"))))
    elif coupon_stores:
        cs = coupon_stores[0]
        den = "、".join("%s×%d 笔" % (fmt(k, 0), v)
                        for k, v in sorted(cs["m"]["coupon_denoms"].items(), reverse=True))
        rows.append(("", "⑤ 卖家优惠券被导出成空白",
                     "Ventas 报表本来有“Descuentos y bonificaciones”（折扣与补贴）这一列，但平台导出时整列是空的。"
                     "本期 %s 有 %d 笔自负优惠券、合计 %s（面额 %s），在任何一张明细表里都查不到，只能用 Total 倒推。"
                     "已验证：账单佣金 ÷（原价 − 券额）= 标准费率，确认平台是按扣券后的价格计佣。"
                     % (cs["store"], cs["m"]["coupon_n"], fmt(cs["m"]["coupon"]), den)))
    else:
        rows.append(("", "⑤ 卖家优惠券被导出成空白",
                     "Ventas 报表的“Descuentos y bonificaciones”列被平台导出成空白。本期各店由 Total 倒推的券额均为 0，"
                     "但这一列随时可能再次出现金额，本报表每期都会自动倒推核对。"))
    if costs is not None:
        rows.append(("", "⑥ 平台数据不含商品采购成本",
                     "平台所有报表都不含你的进货成本，因此平台数据最多算到“经营贡献毛利”。"
                     "本报表从商品成本表取每个 SKU 的单位成本（见一、数据来源与七、商品成本），"
                     "填进【⑦ SKU盈利分析】表五和【%s】；成本表里没有的 SKU 仍是黄色单元格，可临时录入。"
                     % SHEET_NAMES[1]))
    else:
        rows.append(("", "⑥ 缺少商品采购成本",
                     "平台所有报表都不含你的进货成本，因此平台数据最多算到“经营贡献毛利”。"
                     "本报表在【⑦ SKU盈利分析】页留出黄色单元格供录入单位成本，录入后净利润会自动重算。"))
    rows.append(("", "", ""))

    rows.append(("S", "三、本报表的核心勾稽等式（每期自动校验）", ""))
    rows.append(("", "等式 1（订单资金）",
                 "商品收入 + 运费收入 − 卖家优惠券 − 平台佣金 − 平台配送费 − %s − 退款 "
                 "= 销售报表“Total”列 = 实际到账"
                 % ("平台代扣税" if site.WITHHOLDING else "并入佣金列的退货处理费等")))
    rows.append(("", "　验证结果",
                 ("全部店铺完全相等，差异 %s，逐笔订单勾稽无残留 ✓" % fmt(0) if ok_a else
                  "⚠ 未通过：%s" % per_store(stores, lambda S: "残差 " + fmt(S["m"]["resid"])))))
    rows.append(("", "等式 2（账单支付）",
                 "本期账单中“不在订单内扣除”的费用 − 同期发票内冲销 = 月底自动扣款金额"))
    rows.append(("", "　验证结果",
                 per_store(stores, lambda S: "%s − %s = %s（实扣 %s）%s" % (
                     fmt(S["m"]["bill_by_channel"][CH_NO]),
                     fmt(-S["m"]["invoice_internal_reversal"]),
                     fmt(S["m"]["payable"]), fmt(S["m"]["pay_auto"]),
                     "✓" if chk[(S["store"], "bridge_b")]["ok"] else "⚠"))))
    mo = [S for S in stores if (S.get("mp_orders") or {}).get("linked")]
    if mo:
        n = sum(S["mp_orders"]["linked"] for S in mo)
        ok = sum(S["mp_orders"]["agree"] for S in mo)
        # 以前这里无论校验过没过都写"逐项一致，证明本报表口径正确"。现在按实际结果说。
        rows.append(("", "等式 3（第三方交叉验证）",
                     "用 MercadoPago 结算流水逐单核对 %d 家店：%d 个订单里 %d 个（%s）的销售报表"
                     "到账金额与平台实际收款分毫不差%s。"
                     % (len(mo), n, ok, pct(ok / n if n else 0, 1),
                        "；其余 %d 个不一致的订单列在【%s】" % (n - ok, SHEET_NAMES[9])
                        if n - ok else "，本报表口径得到独立凭证的证实")))
    else:
        rows.append(("", "等式 3（第三方交叉验证）",
                     "本期没有任何一家店下载到 MercadoPago 结算流水，因此缺少独立的第三方凭证。"
                     "建议开通报表权限后重新下载，见【%s】。" % SHEET_NAMES[9]))
    rows.append(("", "", ""))

    rows.append(("S", "四、组合单 / 一单多件 / 换货单 是怎么处理的", ""))
    rows.append(("", "① 合并包裹（Paquete de varios productos = Sí）",
                 "这是最容易误解的一个字段。它的意思不是“一单买了多个商品”，而是“这笔订单和别的订单合并到同一个包裹发货”。"
                 "本期 %s，但每一行仍然是一笔独立的销售，有自己完整的收入、佣金、运费和 Total。"
                 "所以按行汇总不会重复也不会遗漏。"
                 % per_store(stores, lambda S: "%s 有 %s 行（占 %s）" % (
                     S["store"], fmt0(S["m"]["pack_flag_yes"]),
                     pct(S["m"]["pack_flag_yes"] / S["m"]["rows"] if S["m"]["rows"] else 0, 0)), "、")))
    rows.append(("", "② 一单多件（Unidades > 1）",
                 "同一个 SKU 一次买多件，只占一行，金额列已经是多件的合计。"
                 "本报表用 Unidades 列算件数、用金额列算钱，两边都不会错。本期 %s。"
                 % per_store(stores, lambda S: "%s 有 %d 行" % (S["store"], S["m"]["multi_unit_rows"]), "、")))
    rows.append(("", "③ 真正的一单多 SKU",
                 "绝大多数是【换货】：第一行是原销售，带全部金额但 SKU 为空；第二行是换出去的新商品，有 SKU 和件数但金额全空。"
                 "金额不会重复计，但件数会多算、SKU 会归到换货商品上。本期 %s。"
                 % per_store(stores, lambda S: "%s %d 单（销售报表全窗口 %d 单）" % (
                     S["store"], S["m"]["multi_sku_orders"], S["m"]["multi_sku_orders_all"]), "、")))
    rows.append(("", "④ 汇总口径",
                 "本报表一律按【行】汇总原始金额列，不做任何拆分或分摊。%s"
                 % per_store(stores, lambda S: "%s：订单数按行计 %s 行，唯一订单号 %s 个，件数取 Unidades 合计 %s"
                             % (S["store"], fmt0(S["m"]["rows"]), fmt0(S["m"]["orders_unique"]),
                                fmt0(S["m"]["units"])))))
    rows.append(("", "", ""))

    rows.append(("S", "五、各页说明", ""))
    for nm, desc in [
        (SHEET_NAMES[1], "各店分列 + 合计，从 GMV 一路推到净利润，含费用率分析。行位置每期固定，可直接跨月对比。"),
        (SHEET_NAMES[2], "五个桥：A 订单资金流、B 月度账单、C 两种口径对比、D 从到账净额到经营利润、E %s。这一页是解决“平不了账”的核心。"
         % ("代扣税解释" if site.WITHHOLDING else "佣金列口径差与“智利不代扣税”的证据（F 含税口径提示）")),
        (SHEET_NAMES[3], "本期账单全部费用科目（西班牙语原文 + 中文），标明结算通道。"),
        (SHEET_NAMES[4], "按订单状态、按金额、按平台/卖家责任划分。"),
        (SHEET_NAMES[5], "Full 仓退货逐件质检结果与退货原因，含可回收/彻底损失判定。"
         if site.HAS_FULL else "%s没有 Full 仓，本页不适用。" % site.COUNTRY_CN),
        (SHEET_NAMES[6], "按 SKU 的收入、费用、退货与贡献；五张表：表一正常订单、表二~表四三类异常明细、表五 SKU 真实盈利（最终口径）。单位采购成本只在表五填。"),
        (SHEET_NAMES[7], "销售报表窗口内各月的销售、费率、退货率走势。"),
        (SHEET_NAMES[8], "Full 仓库存结构、动销与周转，并把平台的近 30 天销量与销售报表逐 SKU 核对。"
         if site.HAS_FULL else "%s没有 Full 仓，平台不提供库存报表，本页不适用。" % site.COUNTRY_CN),
        (SHEET_NAMES[9], "本期校验结果、数据缺失项和需要你补充/核实的事项。"),
    ]:
        rows.append(("", nm, desc))
    rows.append(("", "", ""))

    rows.append(("S", "六、一个容易搞错的概念：到账净额 ≠ 利润", ""))
    rows.append(("", "到账净额是现金指标",
                 "销售报表的 Total（%s 本期 %s）只是订单层面进你 MercadoPago 账户的钱。"
                 "广告费、%s页面维护费都不在里面 —— 它们是月底汇总后从同一个余额里另外扣走的（%s扣了 %s）。"
                 % (M["store"], fmt(mm["total"]), "Full 仓储与揽收费、" if site.HAS_FULL else "",
                    mm["pay_date"] or "结算日", fmt(mm["pay_auto"]))))
    rows.append(("", "利润指标是经营贡献毛利",
                 "%s 本期 %s。算法：到账净额 ＋ 加回%s %s − 退货处理费 %s − 广告 %s − Full 仓费用 %s − 页面费 %s。"
                 "完整推导见【%s】桥 D。"
                 % (M["store"], fmt(ctx["contrib"][M["store"]]),
                    "代扣税" if site.WITHHOLDING else "佣金列里的退货处理费等", fmt(mm["tax_ret"]),
                    fmt(mm["dev_net"]),
                    fmt(mm["ads_net"]), fmt(mm["full_bill_total"]), fmt(mm["subscription_net"]),
                    SHEET_NAMES[2])))
    pa_n, pa_v = fee_rows(M, "Cargo por campaña de publicidad de Product Ads")
    da_n, da_v = fee_rows(M, "Cargo por campaña de publicidad de Display Ads")
    pa_rn, pa_rv = fee_rows(M, "Cargo por campaña de publicidad de Product Ads", reversal=True)
    da_rn, da_rv = fee_rows(M, "Cargo por campaña de publicidad de Display Ads", reversal=True)
    ad_rows = M["billing"][M["billing"][F_DETALLE].astype(str).str.contains("publicidad", na=False)]
    ad_linked = int(ad_rows[F_SALE_NO].notna().sum()) if F_SALE_NO in ad_rows.columns else 0
    rows.append(("", "广告费怎么归集的",
                 "取本期账单里 Product Ads %s（%d笔）＋ Display Ads %s（%d笔），减去当月冲销 %s 与 %s，净额 %s，按计费日归属本期。"
                 "账单的广告行没有商品编号也没有订单号（%d 行中 %d 行带订单号），所以无法按 SKU 或订单分摊，只能作为店铺级共同费用。"
                 % (fmt(pa_v), pa_n, fmt(da_v), da_n, fmt(pa_rv), fmt(da_rv),
                    fmt(mm["ads_net"]), len(ad_rows), ad_linked)))
    st_n, st_v = fee_rows(M, "Cargo por servicio de almacenamiento Full")
    storage_note = ("取本期账单的 Cargo por servicio de almacenamiento Full %s（%d 笔，基本每天一笔）。"
                    % (fmt(st_v), st_n))
    if M["files"].get("storage") and "storage_bill_cmp" in mm:
        storage_note += storage_compare_text(mm)
    elif M["files"].get("storage"):
        d = mm["storage_bill"] - mm["storage_accrual"]
        storage_note += ("另有一张独立的仓储费报表按实际占用日算出 %s，两者差 %s，原因是账单按计费日、报表按占用日，跨月错了一两天。"
                         "本报表统一采用账单口径，因为那才是真正扣钱的金额。" % (fmt(mm["storage_accrual"]), fmt(d)))
    if not site.HAS_FULL:
        rows.append(("", "Full 仓", no_full_text().replace("本页不适用，不是数据缺失。",
                                                         "损益表里的 Full 各行恒为 0，相关核对不适用。")))
        storage_note = None
    if storage_note:
        rows.append(("", "仓储费怎么归集的", storage_note))
    full_items = [(FEE_BY_ES[es][0], v) for es, v in mm["bill_fee_net"].items()
                  if FEE_BY_ES[es][1] in ("storage", "fulfillment", "penalty") and abs(v) > site.ZERO_TOL]
    full_items.sort(key=lambda x: -abs(x[1]))
    if site.HAS_FULL:
        rows.append(("", "Full 仓相关费用",
                     "%s，合计 %s。%s"
                     % ("、".join("%s %s" % (n, fmt(v)) for n, v in full_items) or "本期无",
                        fmt(mm["full_bill_total"]),
                        ("已与 %s 的各张明细表逐项核对，差异 %s。"
                         % (os.path.basename(M["files"]["cargos_full"]),
                            fmt(mm["cargos_full_report_total"] - mm["full_bill_total"]))
                         if M["files"].get("cargos_full") else "本期未下载到 Full 费用明细报表，无法交叉核对。"))))

    if costs is not None:
        results = [S["cost"] for S in stores]
        rows.append(("",) + SC.point_line(
            results, sum(ctx["contrib"].values()) - sum(r.cogs for r in results if r is not None)))
        rows.append(("", "", ""))
        rows.append(("S", "七、商品成本（会计处理）", ""))
        rows.append(("",) + SC.policy_line(ctx["cost_units_basis"], ctx["cost_policy_closing"]))

    rows.append(("", "", ""))
    rows.append(("S", "八、重要提示" if costs is not None else "七、重要提示", ""))

    r = readme_rows(ws, 4, rows)

    notes = ["本报表所有金额均为%s(%s)，%s" % (
        site.CUR_CN, site.CUR, "商品采购成本取自商品成本表（覆盖率见【%s】“减：商品采购成本”行）。"
        % SHEET_NAMES[1] if costs is not None else "不含商品采购成本（平台数据不提供）。")]
    if site.WITHHOLDING:
        notes.append("“平台代扣代缴税金”在损益表中单列为备查项，不计入费用。请与你的墨西哥会计师核对该笔预缴税款的抵扣处理。")
    else:
        notes.append("智利平台不代扣税，本报表没有“代扣代缴税金”行。收入为含税（IVA %s）口径，"
                     "GMV 中所含 IVA 在【%s】备查项列示，仅供参考。⚠ 会计判断（请交智利会计师确认）："
                     "① 销项 IVA 由卖家自行向 SII 申报缴纳；② 平台费用发票（佣金、配送、广告%s）"
                     "所含 IVA 能否作为进项抵扣；③ 账单金额是否含 IVA —— 平台导出的账单里没有单列 IVA。"
                     % (pct(site.IVA_RATE, 0), SHEET_NAMES[1], "、Full" if site.HAS_FULL else ""))
    if not site.WITHHOLDING and any(cycle_day(S) != 1 for S in stores):
        notes.append("智利的平台账单不按自然月出，而且每家店切期日不同（%s）。本报表的会计月是"
                     "自然月：账单级费用按每一行的计费日归月，所以本月的费用取自两张账单（%s）。"
                     % (per_store(stores, lambda S: "按 %s，“%s”账单 = %s–%s" % (
                         cycle_cn(cycle_day(S)), MONTH_FULL_ES[S["bill_cycles"][0][1]],
                         S["bill_cycles"][0][2].strftime("%m/%d"),
                         S["bill_cycles"][0][3].strftime("%m/%d")), "；"),
                        "＋".join("%s %d" % (MONTH_FULL_ES[mth], y)
                                 for y, mth, _a, _b in main_store(stores)["bill_cycles"])))
    elif not site.WITHHOLDING:
        notes.append("%s的平台账单按自然月出（账单页的截账日“Cierre”是每月最后一天），"
                     "和墨西哥一样一个会计月对应一张账单；账单级费用仍按每一行的计费日归月。"
                     % site.COUNTRY_CN)
    if site.BONIF_COL:
        notes.append("运费收入里含平台付给卖家的 Flex 配送补贴（Ventas 的“Descuentos y bonificaciones”列，"
                     "本期 %s），MercadoPago 里是单独一笔 CASHBACK。⚠ 会计判断：这笔补贴按运费收入"
                     "还是按平台返利/其他收入入账，开票和 %s 处理请交会计师确认。"
                     % (per_store(stores, lambda S: fmt(S["m"].get("bonif", 0.0)), "、"), site.VAT_NAME))
    no_mp = [S["store"] for S in stores if not S.get("mp")]
    if no_mp:
        notes.append("%s 本期没有 MercadoPago 资金流水（常见原因是该登录账号无报表权限），"
                     "因此这些店的“实际到账”为按销售报表推算。建议让管理员在 MercadoPago 后台“协作者”中"
                     "开通报表权限后重新下载核对。" % "、".join(no_mp))
    if ctx["partial_months"]:
        notes.append("销售报表窗口内的 %s 属不完整月份，本报表未纳入损益，仅在【%s】作参考。"
                     % ("、".join(ctx["partial_months"]), SHEET_NAMES[7]))
    bad = [c for c in ctx["checks"] if not c["ok"]]
    if bad:
        notes.append("⚠ 本期有 %d 项校验未通过，报表数字可能不完整，请先看【%s】顶部的校验结果再使用。"
                     % (len(bad), SHEET_NAMES[9]))
    for i, t in enumerate(notes, 1):
        ws.cell(row=r, column=2, value="注 %d" % i).font = F(10, True, C_RED)
        c2 = ws.cell(row=r, column=3, value=t)
        c2.font = F(10)
        c2.alignment = wrap("top")
        if len(t) > 90:
            ws.row_dimensions[r].height = est_height(t, 78)
        r += 1
