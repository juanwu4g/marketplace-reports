# -*- coding: utf-8 -*-
"""MercadoLibre report — ⑩ 数据缺口与待办: failed checks, missing files, to-dos."""

import os

import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from core.report_kit import (BOX, C_AMBER, C_BAD, C_BAND, C_GREEN, C_GREY, C_NAVY, C_OK, C_RED,
    C_SUB, C_WARN, F, est_height, fill, header_cells, wrap)

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.site import is_pe
from platforms.mercadolibre.report.catalog import F_INVOICE, SHEET_NAMES, V_SKU, fmt, money, pct
from platforms.mercadolibre.report.loading import cycle_day, date_span
from platforms.mercadolibre.report.sheet_common import (agg_tax_rate, storage_compare_text,
    store_count_cn)


# ══════════════════════════════════════════════════════════════════════
# ⑩ 数据缺口与待办（含本期校验结果）
# ══════════════════════════════════════════════════════════════════════

def sheet_gaps(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[9])
    ws["A1"] = "本期校验结果 · 数据缺口 · 待办清单"
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = "上半部分是脚本每期自动跑的硬校验；下半部分是需要人处理的事项，按优先级排序。"
    ws["A2"].font = F(9, False, C_GREY)
    for i, w in enumerate([10, 30, 60, 54, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 4
    ws.cell(row=r, column=1, value="一、本期校验结果").font = F(12, True, C_NAVY)
    r += 1
    header_cells(ws, r, ["结果", "店铺", "校验项", "明细", ""], wrap_text=False)
    r += 1
    for c0 in ctx["checks"]:
        mark = "跳过" if c0["skipped"] else ("通过 ✓" if c0["ok"] else "未通过 ✗")
        col = C_GREY if c0["skipped"] else (C_GREEN if c0["ok"] else C_RED)
        bg = C_WARN if c0["skipped"] else (C_OK if c0["ok"] else C_BAD)
        ws.cell(row=r, column=1, value=mark).font = F(10, True, col)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=2, value=c0["store"]).font = F(10)
        ws.cell(row=r, column=3, value=c0["name"]).font = F(10)
        d = ws.cell(row=r, column=4, value=c0["detail"])
        d.font = F(9)
        d.alignment = wrap("top")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        ws.row_dimensions[r].height = est_height(c0["detail"], 70)
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BOX
            ws.cell(row=r, column=c).fill = fill(bg)
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="二、数据缺口与待办").font = F(12, True, C_NAVY)
    r += 1
    header_cells(ws, r, ["优先级", "事项", "现状 / 发现", "建议动作", "对财务的影响"], wrap_text=False)
    ws.row_dimensions[r].height = 22
    r += 1

    items = []
    for S in stores:
        if S.get("ventas_warning"):
            items.append(("高", "%s 的 Ventas 导出不完整 ⚠" % S["store"],
                          "平台提示「%s」；源文件：%s。报表已使用文件内现有销售数据。"
                          % (S["ventas_warning"], os.path.basename(S["files"]["ventas"])),
                          "重新生成并下载 Ventas，核对订单数和金额后重新出报表。",
                          "本店及包含本店的合计金额可能偏低，当前数字不可视为完整结果。"))
    nsku = sum(S["ventas"][V_SKU].nunique() for S in stores)
    items.append(("高", "补齐商品采购成本",
                  "平台所有报表都不含进货成本，因此本报表只能算到“经营贡献毛利”（本期合计 %s）。"
                  % fmt(sum(ctx["contrib"].values())),
                  "在【%s】表五的黄色“单位采购成本”列填入每个 SKU 的到岸单位成本（进货价＋头程运费＋关税＋清关杂费），"
                  "本期共 %d 个 SKU。" % (SHEET_NAMES[6], nsku),
                  "这是算出真实净利润的唯一缺口，填完即可得出完整损益。"))

    unk_all = [(S, u) for S in stores for u in S["unknown_fees"]]
    if unk_all:
        items.append(("高", "出现未登记的费用科目 ⚠",
                      "本期发现 %d 个 FEE_CATALOG 里没有的科目：%s。"
                      "它们的金额已计入报表的“未登记科目”行以保证勾稽不断，但科目归属未确认。"
                      % (len(unk_all), "；".join("%s（%s，%s，%d 笔）" % (
                          u["detalle"], S["store"], fmt(u["amount"]), u["count"]) for S, u in unk_all)),
                      "确认这些科目的业务含义后，加入 report_build.py 的 FEE_CATALOG 常量表"
                      "（写明中文名、大类、是否订单级），并在 PL_LINES 里给它一行。",
                      "未登记科目会让损益表的费用分类不完整；金额本身没有丢，但看不出是什么钱。"))

    no_mp = [S["store"] for S in stores if not S.get("mp")]
    if no_mp:
        items.append(("高", "MercadoPago 报表权限",
                      "%s 本期没有下载到 settlement 结算流水，无法用第三方凭证交叉验证到账金额。"
                      % "、".join(no_mp),
                      "让账号管理员在 MercadoPago 后台“协作者(Colaboradores)”中开通"
                      "“查看收款与账单报表”和“查看操作报表”两项权限，然后重新下载结算流水。",
                      "缺少独立凭证时，到账金额只能按销售报表推算，无法验证平台是否少付。"))

    for S in stores:
        MO = S.get("mp_orders") or {}
        dif = MO.get("differ") or []
        if not dif:
            continue
        zero = [x for x in dif if abs(x[1]) <= site.CN_TOL and x[2] > 0]
        items.append(("中", "MercadoPago 逐单不一致：%s" % S["store"],
                      "%d 个订单的销售报表到账与平台实际收款不一致，合计差 %s（平台 − 销售报表）。"
                      "其中 %d 单销售报表记为 0、平台却有到账（合计 %s）。逐单（订单号：销售报表 → "
                      "平台）：%s%s"
                      % (len(dif), fmt(sum(x[2] - x[1] for x in dif)), len(zero),
                         fmt(sum(x[2] for x in zero)),
                         "；".join("%s：%s → %s" % (o, fmt(v), fmt(mp)) for o, v, mp in dif[:40]),
                         "；…另 %d 单" % (len(dif) - 40) if len(dif) > 40 else ""),
                      (("其中 %d 单（%s）的差额等于该单的运费收入：秘鲁买家另付的运费（SETTLEMENT_SHIPPING）"
                        "和 Flex 配送补贴（CASHBACK）在 MercadoPago 都是单独一笔，不是少付，可忽略。其余的"
                        if site.BONIF_COL else
                        "其中 %d 单（%s）的差额等于买家另付的运费：智利运费在 MercadoPago 是单独一笔"
                        "（SETTLEMENT_SHIPPING），不是少付，可忽略。其余的") % (
                           MO["ship_explained"][0], fmt(MO["ship_explained"][1]))
                       if MO.get("ship_explained") and MO["ship_explained"][0] else "")
                      + "在卖家后台按订单号查这些订单的退款 / 争议处理结果，确认哪一边是对的。",
                      "销售报表若漏记了这些到账，本店的到账和毛利偏低；反之则偏高。"))

    for S in stores:
        m = S["m"]
        if m["ing"] and m["ads_net"] / m["ing"] > 0.04:
            share = m["ads_net"] / ctx["contrib"][S["store"]] if ctx["contrib"][S["store"]] else 0
            items.append(("高", "广告投入产出需要复核（%s）" % S["store"],
                          "本期广告净支出 %s，占 GMV %s，占经营贡献毛利 %s。"
                          % (fmt(m["ads_net"]), pct(m["ads_net"] / m["ing"], 1), pct(share, 1)),
                          "按 Product Ads 后台的广告活动(campaña)导出 ACOS/ROAS 明细找出无效活动。"
                          "平台账单不提供按商品的广告拆分，必须从广告后台单独导。",
                          "这是金额最大的可控成本项之一。"))
        if len(S["returns"]):
            R = S["returns"]
            lost = int((R["Resultado de la revisión"] == "Producto para retirar en centro de distribución").sum())
            if lost and lost / len(R) > 0.3:
                top = R.groupby("Estado del producto").size().sort_values(ascending=False)
                items.append(("中", "Full 仓退货“需自提”占比过高（%s）" % S["store"],
                              "本期质检 %d 件中 %d 件（%s）被判定“需到配送中心自提”，即不可再售。"
                              "主因：%s。" % (len(R), lost, pct(lost / len(R), 0),
                                            "、".join("%s %d 件" % (k, v) for k, v in top.head(3).items())),
                              "① 尽快决定这批货是自提翻新还是放弃（放着会产生长龄库存费）；"
                              "② 针对封条破损/包装类原因，检查外包装与封装工艺，这类退货多数可以从源头避免。",
                              "每件都是“钱退了、货也卖不掉”的双重损失，同时还要付退仓费。"))
        space = sum(v for es, v in m["bill_fee_net"].items()
                    if es in ("Cargo por comprar más espacio en Full",
                              "Cargo por sobrepasar espacio Full",
                              "Cargo por stock antiguo en Full"))
        if space > 0:
            items.append(("中", "Full 库存空间与库龄成本（%s）" % S["store"],
                          "本期购买额外仓位 %s、超仓位罚金 %s、长龄库存费 %s，合计 %s。"
                          % (fmt(m["bill_fee_net"].get("Cargo por comprar más espacio en Full", 0)),
                             fmt(m["bill_fee_net"].get("Cargo por sobrepasar espacio Full", 0)),
                             fmt(m["bill_fee_net"].get("Cargo por stock antiguo en Full", 0)), fmt(space)),
                          "核对 stock_general_full 报表中“需清理”的库存，把滞销 SKU 从 Full 仓撤出或做清仓。",
                          "仓位与库龄费合计占经营贡献毛利 %s，属可优化项。"
                          % pct(space / ctx["contrib"][S["store"]] if ctx["contrib"][S["store"]] else 0, 1)))
        pen = m["bill_fee_net"].get("Cargo por incumplimiento en Envíos Full", 0)
        if pen > 0:
            items.append(("中", "Full 发货违规罚金（%s）" % S["store"],
                          "本期被罚 %s。" % fmt(pen),
                          "核对相关发货批次的实发数量与申报数量，找出差异原因，必要时向平台申诉。",
                          "金额通常不大，但反映入仓流程有问题，容易反复发生。"))
        if m["coupon_n"]:
            items.append(("中", "优惠券列被平台导出成空白（%s）" % S["store"],
                          "Ventas 报表的 Descuentos y bonificaciones 列整列为空，但本期实际发生 %d 笔"
                          "卖家自负优惠券、合计 %s。本报表用 Total 倒推还原。"
                          % (m["coupon_n"], fmt(m["coupon"])),
                          "① 到促销后台核对优惠券活动的实际投放与预算；"
                          "② 每月导出时留意这一列，若平台修复了导出，可直接取用不必倒推。",
                          "占 GMV %s。这是唯一一笔“账面完全看不见”的支出，不还原就永远差这一块对不上。"
                          % pct(m["coupon"] / m["ing"] if m["ing"] else 0, 2)))
        if S["files"].get("returns") and len(S.get("returns_all", [])):
            ra = S["returns_all"]
            items.append(("中", "退货质检报表期间不匹配（%s）" % S["store"],
                          "Returns 报表按【质检日】导出，本次覆盖 %s；本期订单的退货可能下月才质检。"
                          % (date_span(ra["fr"])
                             or "（质检日一列读不出来 —— 表头或日期格式可能变了，"
                                "先看运行日志里的警告）"),
                          "每月固定在月结后 15 天再导一次 Returns，用【订单号】与销售报表关联，"
                          "才能把退货损失准确归到销售月份。",
                          "本期退货损失可能被低估（部分本期订单的退货尚未质检完）。"))
        if not S["files"].get("cargos_full") and site.HAS_FULL:
            items.append(("中", "缺少 Full 费用明细报表（%s）" % S["store"],
                          "没有找到本期的 Reporte_Cargos_Full 文件，无法与账单的 Full 科目交叉核对。",
                          "在 Full 后台重新导出本期的 Cargos Full 报表。",
                          "少了一道校验；账单口径仍然准确，但无法独立验证。"))
        if S["files"].get("storage") and "storage_bill_cmp" in m \
                and m.get("storage_days", 0) < m.get("storage_month_days", 0):
            items.append(("低", "仓储费报表只覆盖本月一部分（%s）" % S["store"],
                          storage_compare_text(m),
                          "无需处理。智利仓储报表按店铺的切期日切期；想覆盖整月，下载时需多拉一期"
                          "（上月切期日起的那一期）。本报表统一采用账单口径。",
                          "不影响损益：损益表的仓储费取自账单，按计费日完整覆盖本月。"))
        elif S["files"].get("storage") and "storage_bill_cmp" not in m:
            d = m["storage_bill"] - m["storage_accrual"]
            if abs(d) > site.TOL_CENT:
                items.append(("低", "仓储费两个口径的时间差（%s）" % S["store"],
                              "账单（按计费日）%s vs 独立仓储报表（按占用日）%s，差 %s。"
                              % (fmt(m["storage_bill"]), fmt(m["storage_accrual"]), fmt(d)),
                              "无需处理。本报表统一采用账单口径，因为那才是真正扣钱的金额。",
                              "影响极小，属正常的跨月错位。"))

    for S in stores:
        for a in S["anomalies"]:
            items.append(("中", "科目出现在非预期通道（%s）" % S["store"], a,
                          "确认该科目的计费方式是否变化，必要时调整 FEE_CATALOG 里的“是否订单级”。",
                          "会影响费用在损益表中的归类位置，不影响合计。"))

    if ctx["partial_months"]:
        items.append(("低", "销售报表窗口两端的月份不完整",
                      "%s 为不完整月份（销售报表是滚动窗口导出）。" % "、".join(ctx["partial_months"]),
                      "下月初重新下载完整数据；长期看应按月累积保存指标，才能得到真实趋势。",
                      "本报表未将这些月份计入损益，仅在【%s】作参考。" % SHEET_NAMES[7]))
    if site.WITHHOLDING:
        items.append(("低", "代扣代缴税金的税务处理",
                      "本期%s被代扣 %s，约为含税售价的 %s。"
                      % (store_count_cn(stores),
                         fmt(sum(S["m"]["tax_ret"] for S in stores)),
                         pct(agg_tax_rate(stores), 2)),
                      "把 MercadoPago 的 Constancia de retenciones（预扣凭证）交给墨西哥会计师，"
                      "确认能否全额抵扣 IVA/ISR 申报。",
                      "若能全额抵扣，这笔钱不是成本；若不能，需重新计入费用，会显著改变净利润。"))
    else:
        items.append(("中", "IVA 与平台费用发票的税务处理（会计判断）",
                      "智利平台不代扣税（MercadoPago 流水 TAXES_AMOUNT 为 0，佣金列与账单佣金逐单相等）。"
                      "本期%s GMV 所含 IVA（%s）约 %s，全部由卖家自行申报。平台账单只给“Valor del cargo”"
                      "一个金额，没有单列 IVA。"
                      % (store_count_cn(stores), pct(site.IVA_RATE, 0),
                         fmt(sum(S["m"]["iva_in_gmv"] for S in stores))),
                      "请智利会计师确认：① 销项 IVA 的申报口径（含 B2B 开票订单、退货冲减）；"
                      "② 平台费用（佣金、配送、广告%s、退货处理费、信誉保障计划扣款）的发票是否含 IVA、"
                      "能否作为进项抵扣；③ 需不需要从 SII 下载平台开给你的电子发票（DTE）逐张核对。"
                      % ("、Full" if site.HAS_FULL else ""),
                      "本报表的经营贡献毛利是含税口径、未扣任何税；税后利润取决于上述判断。"))
        if is_pe():
            # 秘鲁的电子发票按 SUNAT 的系列号分：F 开头是 factura（能抵进项 IGV），
            # B 开头是 boleta de venta（开给最终消费者，不能抵）。IDG_PE 2026-09 的
            # 平台账单全部是 B001-…，说明平台多半没按 RUC 给这家店开 factura。
            for S in stores:
                inv = S["billing"][F_INVOICE].astype(str).str.strip() \
                    if F_INVOICE in S["billing"].columns else pd.Series(dtype=str)
                ser = inv[inv.str.match(r"^[A-Z]\d{3}-\d+$")].str[0].value_counts()
                if ser.get("B", 0) and not ser.get("F", 0):
                    items.append(("中", "平台费用开的是 boleta 不是 factura（%s）" % S["store"],
                                  "本期账单 %d 行的发票号（N° de factura fiscal）都是 B 开头（如 %s），"
                                  "按 SUNAT 的编号规则这是 boleta de venta，不是 factura。"
                                  % (int(ser["B"]), inv[inv.str.startswith("B")].iloc[0]),
                                  "请会计师确认：boleta 上的 IGV 一般不能作为进项（crédito fiscal）抵扣；"
                                  "若公司有 RUC，查一下卖家账号里登记的开票资料，看能否让平台改开 factura。",
                                  "平台费用（本期账单 %s）里的 IGV 能不能抵扣，取决于这一点。"
                                  % fmt(S["m"]["bill_total"])))
        for S in stores:
            if cycle_day(S) != 1 and not S["m"].get("pay_complete"):
                items.append(("中", "桥 B 付款记录不全（%s）" % S["store"],
                              "；".join(S["m"].get("pay_missing", [])) + "。",
                              "等该账单自动扣款后（通常次月 7 号关账、约一周后扣款）重新下载账单再重出。",
                              "不影响损益表（费用按账单行计），只是暂时无法证实这些费用已付清。"))

    order = {"高": 0, "中": 1, "低": 2}
    items.sort(key=lambda x: order.get(x[0], 9))
    col = {"高": C_RED, "中": C_AMBER, "低": C_GREEN}
    bg = {"高": C_BAD, "中": C_WARN, "低": C_OK}
    for pri, tit, now, act, imp in items:
        ws.cell(row=r, column=1, value=pri).font = F(11, True, col[pri])
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=1).fill = fill(bg[pri])
        ws.cell(row=r, column=2, value=tit).font = F(10, True)
        ws.cell(row=r, column=2).alignment = wrap("top")
        for c, txt in ((3, now), (4, act), (5, imp)):
            cc = ws.cell(row=r, column=c, value=txt)
            cc.font = F(9)
            cc.alignment = wrap("top")
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BOX
        ws.row_dimensions[r].height = max(est_height(now, 48), est_height(act, 44), est_height(imp, 32))
        r += 1

    # ---- 三、贷记明细 × 销售报表 逐单比对的异常清单 ----
    anyrow = any(S["cn"]["orphan"] or S["cn"]["unrev"] or S["cn"]["partial"]
                 or S["cn"]["mismatch"] for S in stores)
    r += 2
    ws.cell(row=r, column=1, value="三、贷记明细 × 销售报表 逐单比对").font = F(12, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "按订单号把账单/贷记单与销售报表逐单对照。关联键有个坑：账单的 Número de venta "
        "与销售报表的 # de venta 不是同一个号段，可靠的桥是 Número de paquete。"
        "下面三类异常，桥A（销售报表内部勾稽）、桥B（账单总额）、MercadoPago（店铺汇总）都看不见。"
    )).font = F(9, False, C_GREY)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.row_dimensions[r].height = est_height(ws.cell(row=r, column=1).value, 100)
    r += 2
    if not anyrow:
        ws.cell(row=r, column=1, value="本期三类异常均为 0，逐单比对全部通过。").font = F(10, False, C_GREEN)
        r += 1

    def block(title, intro, cols, widths, rows, sum_cols=()):
        nonlocal r
        if not rows:
            return
        ws.cell(row=r, column=1, value=title).font = F(11, True, C_NAVY)
        r += 1
        ws.cell(row=r, column=1, value=intro).font = F(9, False, C_GREY)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(5, len(cols)))
        ws.row_dimensions[r].height = est_height(intro, 150)
        r += 1
        header_cells(ws, r, cols, color=C_BAND)
        for i, w in enumerate(widths, 1):        # 只加宽、不收窄：上面的表已经定过列宽
            if ws.column_dimensions[get_column_letter(i)].width < w:
                ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[r].height = 28
        r += 1
        start = r
        for vals in rows:
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                c.font = F(9)
                c.border = BOX
                if isinstance(v, float):
                    c.number_format = site.MNY
            r += 1
        ws.cell(row=r, column=1, value="合计 %d 单" % len(rows)).font = F(10, True)
        # 要合计哪几列由调用方指定，不靠表头文字去猜 —— 猜法会把"MP实际净额"
        # 这种加起来没有意义的列也求和。
        for i in sum_cols:
            cl = get_column_letter(i)
            ws.cell(row=r, column=i, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)
                    ).number_format = site.MNY
        for i in range(1, len(cols) + 1):
            ws.cell(row=r, column=i).fill = fill(C_SUB)
            ws.cell(row=r, column=i).font = F(10, True)
            ws.cell(row=r, column=i).border = BOX
        r += 2

    block("3.1 账单有费用、销售报表查无此单（且未被冲平）",
          "下单后立刻作废的单，平台会计费再全额冲销，净额为 0，属正常，已排除。"
          "留在这里的是净额不为零的 —— 平台为一笔销售报表上不存在的订单收了钱。"
          "这笔钱既不在销售报表里，也不会被桥A 发现，只能人工向平台核实。",
          ["店铺", "账单上的订单键", "账单净额", "账单行数"], [18, 22, 14, 10],
          [(S["store"], k, money(amt), cnt)
           for S in stores for k, amt, cnt in S["cn"]["orphan"]], sum_cols=(3,))

    block("3.2 取消单计了费却没有冲销",
          "发货前取消的订单，平台通常压根不下账单，所以“没有贷记单”本身不是问题 —— "
          "这里只列确实计过费、账单净额不为零的。",
          ["店铺", "订单号", "账单净额", "订单状态", "销售日", "SKU"],
          [18, 22, 14, 34, 13, 18],
          [(S["store"], k, money(net), st,
            ("" if d is None or pd.isna(d) else d.strftime("%Y-%m-%d")), sku)
           for S in stores for k, net, st, d, sku in S["cn"]["unrev"]], sum_cols=(3,))

    block("3.3 部分退款订单：销售报表把原佣金重复扣了一次",
          ("平台按退款比例 p 同时退还佣金与代扣税（代扣税在账单里没有科目，只在 MercadoPago 流水可见）。"
           "应退 ＝ p ×（GMV − 佣金 − 代扣税）；" if site.WITHHOLDING else
           "平台按退款比例 p 退还佣金（智利不代扣税；公式里“佣金列其余部分”是并进佣金列的退货处理费等）。"
           "应退 ＝ p ×（GMV − 佣金 − 佣金列其余部分）；")
          + "而销售报表的“退款与取消”列写的是 应退 ＋ 原佣金，"
          "于是多扣一个原佣金。“认定依据”为现金凭证的，是用 MercadoPago 实际净额与销售报表 Total "
          "的差值直接证实（该差值精确等于原佣金）；无流水的店只能按规律推定。"
          "本表合计已计入损益表的“加：部分退款重复扣佣金调整”行。",
          ["店铺", "订单号", "SKU", "退款比例", "原佣金＝多扣金额", "GMV",
           "销售报表退款", "闭式应退", "销售报表Total", "MP实际净额", "认定依据"],
          [18, 22, 18, 10, 16, 12, 14, 12, 14, 13, 34],
          [(S["store"], hit, sku, pct(p, 2), money(com), money(gmv), money(vref),
            money(want), money(vtot), (money(real) if real is not None else "—"),
            "现金凭证确认" if basis == "现金凭证" else "按规律推定（%s）" % proof)
           for S in stores
           for hit, p, com, gmv, tax, vref, want, vtot, real, proof, d, sku, basis
           in S["cn"]["partial"]], sum_cols=(5, 6, 7, 8))

    block("3.4 部分冲销、但金额对不上已知模式（未计入调整，需人工核对）",
          "这些单的佣金确实被部分冲销了，但销售报表的退款金额与上面的规律对不上，"
          "说明还叠加了别的调整。没有把它们计入损益表调整 —— 宁可少调也不要凭猜测调。",
          ["店铺", "订单号", "SKU", "退款比例", "原佣金", "销售报表退款",
           "按规律应为", "差额", "MP实际净额"],
          [18, 22, 18, 10, 12, 14, 14, 12, 13],
          [(S["store"], hit, sku, pct(p, 2), money(com), money(vref),
            money(vref - gap), money(gap), (money(real) if real is not None else "—"))
           for S in stores
           for hit, p, com, gmv, tax, vref, want, vtot, real, proof, d, sku, basis, gap
           in S["cn"]["mismatch"]], sum_cols=(5, 6))
    ws.freeze_panes = "A5"
