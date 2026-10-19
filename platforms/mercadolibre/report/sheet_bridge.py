# -*- coding: utf-8 -*-
"""MercadoLibre report — ③ 平账勾稽表: the bridges (orders ↔ money, bill ↔ fees,
withheld tax or the "nothing withheld" check)."""

from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from core.report_kit import (BOX, C_BAND, C_CALC, C_GREEN, C_GREY, C_INPUT, C_NAVY, C_OK, C_RED,
    C_SUB, C_WARN, F, INT, PCT, PCT2, PCT3, est_height, fill, wrap)

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (CH_NA, CH_NO, CH_SI, MONTH_FULL_ES,
    SHEET_NAMES, fmt, money, pct)
from platforms.mercadolibre.report.loading import cycle_cn, cycle_day
from platforms.mercadolibre.report.checks import CL_VCOL_REST_MAX
from platforms.mercadolibre.report.sheet_common import Sheet, main_store, per_store, store_cols


# ════════════════════════════════════════════════════════════════════════
# ③ 平账勾稽表
# ════════════════════════════════════════════════════════════════════════

def _bridge_head(wb, stores, ctx):
    n = len(stores)
    sh = Sheet(wb, SHEET_NAMES[2], first_store_col=3, nstores=n)
    ws = sh.ws
    ncol_idx = sh.tcol_idx + 1
    ws["A1"] = "平账勾稽表（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("把“销售报表 / 平台账单 / 实际资金”三张表逐项打通。"
                "每个桥的最后一行都是差异检验，✓ 表示完全勾稽。单位：%s" % site.CUR)
    ws["A2"].font = F(9, False, C_GREY)
    for i, w in enumerate([6, 44] + [16] * n + [16, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    state = {"r": 4}

    def block(t, sub=""):
        r = state["r"]
        c = ws.cell(row=r, column=1, value=t)
        c.font = F(12, True, "FFFFFF")
        for k in range(1, ncol_idx + 1):
            ws.cell(row=r, column=k).fill = fill(C_NAVY)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol_idx)
        ws.row_dimensions[r].height = 20
        r += 1
        if sub:
            c = ws.cell(row=r, column=1, value=sub)
            c.font = F(9, False, C_GREY)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol_idx)
            c.alignment = wrap()
            ws.row_dimensions[r].height = est_height(sub, 110, 14, 28)
            r += 1
        state["r"] = r

    def head(labels=None):
        r = state["r"]
        cols = labels or (["#", "项目"] + store_cols(stores) + ["合计", "说明"])
        for i, t in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_BAND)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BOX
        state["r"] = r + 1

    def line(no, label, values, note, kind="d", total=True, fmt_code=site.MNY):
        """values: 可以是 list（每店一个值/公式），也可以是 callable(col_letter, j)。"""
        r = state["r"]
        ws.cell(row=r, column=1, value=no).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=label)
        for j in range(n):
            v = values(sh.scol(j), j) if callable(values) else values[j]
            c = ws.cell(row=r, column=sh.c0 + j)
            if isinstance(v, str) or v is None:
                c.value = v
            else:
                c.value = money(v)
            c.number_format = fmt_code
        if total:
            ws.cell(row=r, column=sh.tcol_idx,
                    value="=SUM(%s%d:%s%d)" % (sh.scol(0), r, sh.scol(n - 1), r)
                    ).number_format = fmt_code
        nc = ws.cell(row=r, column=ncol_idx, value=note)
        nc.font = F(9, False, "404040")
        nc.alignment = wrap()
        for c in range(1, ncol_idx + 1):
            ws.cell(row=r, column=c).border = BOX
            if kind == "s":
                ws.cell(row=r, column=c).fill = fill(C_SUB)
                ws.cell(row=r, column=c).font = F(10, True)
            elif kind == "t":
                ws.cell(row=r, column=c).fill = fill(C_CALC)
                ws.cell(row=r, column=c).font = F(10, True, C_NAVY)
            elif kind == "chk":
                ws.cell(row=r, column=c).fill = fill(C_OK)
                ws.cell(row=r, column=c).font = F(10, True, C_GREEN)
            else:
                ws.cell(row=r, column=c).font = F(10)
        nc.font = F(9, False, "404040")
        nc.alignment = wrap()
        if note and len(note) > 60:
            ws.row_dimensions[r].height = est_height(note, 56)
        state["r"] = r + 1
        return r

    def gap():
        state["r"] += 1

    M = main_store(stores)
    mm = M["m"]
    V = lambda f: [f(S) for S in stores]

    # ───────── 桥 A ─────────
    block("桥 A　订单资金流：从销售报表推算到实际到账",
          "销售报表(%s)的每一行是一笔订单。平台在给你打款时，已经把佣金、配送费和%s直接扣掉了，"
          "所以“Total”列就是这笔订单真正进你 MercadoPago 账户的钱。"
          % (site.VENTAS_SHEET, "代扣税" if site.WITHHOLDING else "退货处理费"))
    head()
    a1 = line("A1", "商品销售收入 (GMV)", V(lambda S: S["m"]["ing"]),
              "%s → Ingresos por productos，按下单日归集" % site.VENTAS_SHEET)
    line("A2", "＋ 运费收入", V(lambda S: S["m"]["env_ing"]),
         "%s → Ingresos por envío%s" % (site.VENTAS_SHEET, " ＋ Descuentos y bonificaciones（Flex 配送补贴）"
                                          if site.BONIF_COL else ""))
    line("A3", "− 平台销售佣金", V(lambda S: -S["m"]["com_gross"]),
         "账单 Facturación → Cargo por venta（纯佣金，含联盟佣金，不含税）")
    line("A4", "− 平台配送费", V(lambda S: -S["m"]["env_cost"]),
         "%s → Costos de envío（已扣平台运费补贴后的净额）" % site.VENTAS_SHEET)
    if site.WITHHOLDING:
        line("A5", "− 平台代扣代缴税金 (IVA/ISR)", V(lambda S: -S["m"]["tax_ret"]),
             "★ 关键项。= Ventas 的“Cargo por venta e impuestos” − 账单的纯佣金。本期实测约为含税售价的 %s。"
             "这是平台替你预缴给墨西哥税局的钱，不是费用。" % pct(mm["tax_rate_gross"], 2))
    else:
        line("A5", "− 佣金列中的退货处理费等（非税）", V(lambda S: -S["m"]["tax_ret"]),
             "＝ %s 的“Cargo por venta e impuestos” − 账单的纯佣金。智利平台不代扣税，"
             "这一列只是把退货处理费（Cargo por devolución）也算了进去，拆解见桥 E。" % site.VENTAS_SHEET)
    line("A6", "− 退款与取消（净）", V(lambda S: -S["m"]["refund"]),
         "%s → Anulaciones y reembolsos" % site.VENTAS_SHEET)
    cs = [S for S in stores if S["m"]["coupon_n"] > 0]
    coupon_note = ("★ 平台把 Descuentos y bonificaciones 这一列导出成了空白，所以这笔钱在任何明细列里都看不到，"
                   "只能用 Total 倒推。已交叉验证：账单佣金 ÷（原价 − 券额）＝ 标准费率。")
    if site.BONIF_COL:
        coupon_note = ("%s的 Descuentos y bonificaciones 列有值，但那是平台给的 Flex 配送补贴（已并入 A2 运费收入），"
                       "不是卖家优惠券。卖家自负的优惠券仍用 Total 倒推。" % site.COUNTRY_CN)
    if cs:
        coupon_note += "本期 " + "；".join(
            "%s %d 笔、合计 %s（面额 %s）" % (
                S["store"], S["m"]["coupon_n"], fmt(S["m"]["coupon"]),
                "、".join("%s×%d" % (fmt(k, 0), v)
                          for k, v in sorted(S["m"]["coupon_denoms"].items(), reverse=True)))
            for S in cs) + "。"
    else:
        coupon_note += "本期各店倒推券额均为 0。"
    a6b = line("A6b", "− 卖家承担优惠券/折扣", V(lambda S: -S["m"]["coupon"]), coupon_note)
    a7 = line("A7", "＝ 推算到账净额",
              lambda c, j: "=SUM(%s%d:%s%d)" % (c, a1, c, a6b),
              "按上述公式推算出的应到账金额", kind="t")
    a8 = line("A8", "销售报表 “Total (%s)” 列实际合计" % site.CUR, V(lambda S: S["m"]["total"]),
              "直接取 %s 的 Total 列求和，作为独立验证" % site.VENTAS_SHEET, kind="s")
    ok_a = all(abs(S["m"]["resid"]) <= site.BRIDGE_A_TOL for S in stores)
    line("A9", "差异检验（A8 − A7）",
         lambda c, j: "=%s%d-%s%d" % (c, a8, c, a7),
         ("全部店铺均为 %s —— 逐笔订单完全勾稽，无任何未解释差额 ✓" % fmt(0) if ok_a else
          "⚠ 存在残差，说明还有未识别的隐藏费用或折扣，请按【%s】的排查步骤处理。" % SHEET_NAMES[9]),
         kind="chk")

    # ───────── 桥 A′ MercadoPago 交叉验证 ─────────
    for S in [s for s in stores if (s.get("mp_orders") or {}).get("linked")]:
        gap()
        MO = S["mp_orders"]
        block("桥 A′　第三方交叉验证：%s 的 MercadoPago 结算流水（逐单）" % S["store"],
              "每个订单的销售报表 Total 对它自己那笔收款在 MercadoPago 的净到账（结算、退款、"
              "争议全算，不限日期）。这是唯一独立于销售报表和账单的第三方凭证。流水覆盖 %s。"
              % MO["span"])
        r = state["r"]
        for i, t in enumerate(("#", "项目", "销售报表推算", "MercadoPago 实际流水", "差异", "说明"), 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_BAND)
            c.alignment = Alignment(horizontal="center")
            c.border = BOX
        state["r"] = r + 1
        cross = [("V1", "能对上收款的订单：到账合计（%d 单）" % MO["linked"],
                  MO["sum_v"], MO["sum_mp"],
                  "%d 单分毫不差；差额全部来自下面列出的 %d 单"
                  % (MO["agree"], len(MO["differ"])))]
        for i, (oid, vt, mp) in enumerate(MO["differ"][:15], 1):
            cross.append(("V1.%d" % i, "订单 %s" % oid, vt, mp,
                          "销售报表把它记成 0，平台却有到账 —— 可能是退款/争议后来按卖家有利结案"
                          if abs(vt) <= site.CN_TOL and mp > 0 else
                          ("差额＝该单运费收入（买家运费 SETTLEMENT_SHIPPING / Flex 配送补贴 CASHBACK，"
                           "单独一笔，收款号不同）" if site.BONIF_COL else
                           "差额＝买家另付的运费（单独一笔 SETTLEMENT_SHIPPING，收款号不同）")
                          if oid in MO.get("ship_ids", ()) else "逐单核对"))
        if len(MO["differ"]) > 15:
            cross.append(("…", "另有 %d 单不一致，全部清单见【%s】"
                          % (len(MO["differ"]) - 15, SHEET_NAMES[9]),
                          sum(v for _, v, _ in MO["differ"][15:]),
                          sum(m_ for _, _, m_ in MO["differ"][15:]), ""))
        for no, lab, calc, act, note in cross:
            r = state["r"]
            ws.cell(row=r, column=1, value=no).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=2, value=lab)
            ws.cell(row=r, column=3, value=money(calc)).number_format = site.MNY
            ws.cell(row=r, column=4, value=money(act)).number_format = site.MNY
            ws.cell(row=r, column=5, value="=C%d-D%d" % (r, r)).number_format = site.MNY
            c = ws.cell(row=r, column=6, value=note)
            c.font = F(9, False, "404040")
            c.alignment = wrap()
            same = abs(money(calc) - money(act)) <= site.CN_TOL
            for k in range(1, 7):
                ws.cell(row=r, column=k).border = BOX
                ws.cell(row=r, column=k).fill = fill(C_OK if same else C_WARN)
            state["r"] = r + 1

    # ───────── 桥 B ─────────
    gap()
    block("桥 B　月度账单：从账单总额到实际扣款",
          "平台账单(Facturación)里的费用分两个结算通道，混在同一张表里。搞不清这一点，就会把佣金和运费重复计算一次。")
    head()
    line("B1", "本期账单全部费用合计", V(lambda S: S["m"]["bill_total"]),
         "Facturación ＋ Notas de Crédito 全部行求和（含冲销行）")
    line("B2", "其中：已在订单打款时扣除 (Descontado = Si)",
         V(lambda S: S["m"]["bill_by_channel"][CH_SI]),
         "佣金、配送费、部分退货费。★ 这部分不会再收第二次，已体现在桥 A 里")
    line("B3", "其中：取消订单的费用与冲销 (No aplica)",
         V(lambda S: S["m"]["bill_by_channel"][CH_NA]),
         "被取消的订单原计费与其冲销，净额为负表示平台净退还")
    b4 = line("B4", "其中：需单独支付 (Descontado = No)",
              V(lambda S: S["m"]["bill_by_channel"][CH_NO]),
              "广告、Full 仓储/揽收/仓位/罚金、页面维护费等", kind="s")
    b5 = line("B5", "− 同一张发票内的冲销",
              V(lambda S: S["m"]["invoice_internal_reversal"]),
              "按“Cargo que bonifica → Número del cargo”关联到被冲销的原始费用，只取原费用属于 No 通道的那些行。")
    b6 = line("B6", "＝ 应付账单金额",
              lambda c, j: "=%s%d+%s%d" % (c, b4, c, b5),
              "这才是真正要掏钱的部分", kind="t")
    if all(cycle_day(S) == 1 for S in stores):
        b7_note = ("Pagos 报表里除贷记单以外的支付行：自动扣款（Pago con débito automático）或"
                   "主动付款（Pago voluntario），最近一笔在 %s" % (mm["pay_date"] or "结算日"))
    else:
        b7_note = ("智利账单不按自然月出（%s），本月的费用分属两张账单。取两张账单 Pagos 报表"
                   "“Detalle de Pagos del mes”页里付给本月计费日费用的那部分（该页逐行写明每笔付款"
                   "付了哪一张费用），最近一笔付款在 %s。各张账单整张的应付与实付见 B10。"
                   % (per_store(stores, lambda S: cycle_cn(cycle_day(S)), "；"),
                      mm["pay_date"] or "结算日"))
    b7 = line("B7", "实际支付金额", V(lambda S: S["m"]["pay_auto"]), b7_note, kind="s")
    ok_b = all(abs(S["m"]["payable"] - S["m"]["pay_auto"]) <= site.TOL_CENT
               for S in stores if S["files"].get("pagos"))
    pay_gap = [S for S in stores if cycle_day(S) != 1 and not S["m"].get("pay_complete")]
    if pay_gap:
        b8_note = ("⚠ 付款记录不全（%s），差异不可作结论，见【%s】"
                   % ("；".join(x for S in pay_gap for x in S["m"].get("pay_missing", [])),
                      SHEET_NAMES[9]))
    elif ok_b:
        b8_note = "全部店铺均为 %s，完全勾稽 ✓" % fmt(0)
    else:
        b8_note = "⚠ 存在差异，见【%s】" % SHEET_NAMES[9]
    line("B8", "差异检验（B7 − B6）",
         lambda c, j: "=%s%d-%s%d" % (c, b7, c, b6), b8_note, kind="chk")
    line("B7a", "其中：与订单挂钩、但走账单支付的费用",
         V(lambda S: S["m"]["order_linked_billed"]),
         "多为联盟推广佣金等。这部分已计入【%s】的订单级费用行，"
         "所以损益表的“月度账单费用小计”会比实际扣款少这一块 —— 不是漏记，是归类不同，全表无重复无遗漏。"
         % SHEET_NAMES[1])
    line("B9", "另：贷记单退还 (Notas de Crédito)", V(lambda S: S["m"]["pay_nc"]),
         "针对取消/退货订单退还的佣金与运费，走红字发票单独退回，"
         "已在桥 A 的“退款与取消（净）”中体现，不要重复计入。")
    if any(cycle_day(S) != 1 for S in stores):
        # 智利：两张账单各自整张的应付 / 实付。一个会计月的桥 B 只用到每张账单的
        # 一部分，整张账单自己平不平要单独看，免得两头的差异互相抵掉看不出来。
        ncyc = max(len(S["m"].get("pay_cycles", [])) for S in stores)
        for k in range(ncyc):
            def cyc(S, key, k=k):
                cs = S["m"].get("pay_cycles", [])
                return cs[k][key] if k < len(cs) else 0.0
            lab = next((S["m"]["pay_cycles"][k] for S in stores
                        if k < len(S["m"].get("pay_cycles", []))), None)
            # 账单名各店相同，切期日各店不同（DAB 8 号、Tecnohub 24 号），周期写在说明里
            name = "%s 账单" % lab["label"] if lab else ""
            line("B10.%d" % (k + 1), "%s整张应付" % name, V(lambda S: cyc(S, "payable")),
                 "计费日在该账单周期内的 No 通道费用＋冲销（不限本月）。周期：%s"
                 % per_store(stores, lambda S: cyc(S, "span") or "—", "、"))
            line("", "%s整张实付" % name, V(lambda S: cyc(S, "paid")),
                 "其中付给本月计费日的部分：%s"
                 % per_store(stores, lambda S: fmt(cyc(S, "applied_month")), "、"))
    # ───────── 桥 C ─────────
    gap()
    block("桥 C　为什么“销售报表的扣费”和“账单的费用”对不上",
          "这是最常见的算错点。两张表统计的是不同的东西，差额可以完全解释。")
    head()
    c1 = line("C1", "销售报表显示的平台扣费（佣金及税金＋配送费）",
              V(lambda S: S["m"]["com_imp"] + S["m"]["env_cost"]),
              "%s 中 Cargo por venta e impuestos ＋ Costos de envío 的绝对值" % site.VENTAS_SHEET)
    if site.WITHHOLDING:
        c2 = line("C2", "− 其中属于代扣代缴税金，不是平台费用",
                  V(lambda S: -S["m"]["tax_ret"]), "★ 差异的最大来源")
    else:
        c2 = line("C2", "− 其中属于退货处理费等（并进了佣金列）",
                  V(lambda S: -S["m"]["tax_ret"]),
                  "智利没有代扣税；这部分在账单里是独立科目 Cargo por devolución，见桥 E")
    c3 = line("C3", "＝ 销售报表口径的真实平台费用",
              lambda c, j: "=%s%d+%s%d" % (c, c1, c, c2), "", kind="t")
    c4 = line("C4", "账单中与订单挂钩的费用（佣金＋配送费，毛额）",
              V(lambda S: S["m"]["com_gross"] + S["m"]["env_gross"]),
              "账单 Cargo por venta ＋ Cargo por envíos（已匹配到本期订单的部分，可能跨两张账单）", kind="s")
    line("C5", "差异（C4 − C3）",
         lambda c, j: "=%s%d-%s%d" % (c, c4, c, c3),
         "剩余小额差异来自：被取消订单在账单里保留了原计费行、以及多商品包裹的行归属。"
         "本期 %s。" % per_store(stores, lambda S: fmt(
             S["m"]["com_gross"] + S["m"]["env_gross"] - (S["m"]["com_imp"] + S["m"]["env_cost"] - S["m"]["tax_ret"])), "、"),
         kind="chk")
    gap()
    c6 = line("C6", "账单中与任何订单都无关的费用", V(lambda S: S["m"]["other_total"]),
              "广告、Full 仓储/揽收/仓位/长龄库存/罚金、页面维护费。"
              "★ 销售报表里完全看不到这些，只看销售报表会高估利润。", kind="s")
    r = line("C7", "占 GMV 比例",
             lambda c, j: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, a1, c, c6, c, a1),
             "本期 %s；其中广告占 %s。" % (
                 per_store(stores, lambda S: "%s %s" % (
                     S["store"], pct(S["m"]["other_total"] / S["m"]["ing"] if S["m"]["ing"] else 0, 1)), "、"),
                 per_store(stores, lambda S: "%s %s" % (
                     S["store"], pct(S["m"]["ads_net"] / S["m"]["ing"] if S["m"]["ing"] else 0, 1)), "、")),
             total=False, fmt_code=PCT2)
    return sh, ws, {"state": state, "a1": a1, "block": block, "head": head,
                    "line": line, "gap": gap, "ncol": ncol_idx}



def _bridge_tail(stores, ctx, sh, ws, H, n):
    """③ 的后半段：桥 D / E / F 与常见误算。游标沿用 _bridge_head 的 state。"""
    state = H["state"]
    block, head, line, gap = H["block"], H["head"], H["line"], H["gap"]
    ncol_idx = H["ncol"]
    M = main_store(stores)
    mm = M["m"]
    V = lambda f: [f(S) for S in stores]

    def banner(text, color=C_WARN, font=None):
        r = state["r"]
        c = ws.cell(row=r, column=2, value=text)
        c.font = font or F(10, True, C_RED)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol_idx)
        c.alignment = wrap()
        ws.row_dimensions[r].height = est_height(text, 100, 16, 32)
        for k in range(1, ncol_idx + 1):
            ws.cell(row=r, column=k).fill = fill(color)
            ws.cell(row=r, column=k).border = BOX
        state["r"] = r + 1

    # ───────── 桥 D ─────────
    gap()
    block("桥 D　从“到账净额”到“经营利润”：为什么到账净额不等于利润",
          "销售报表的 Total（到账净额）只是订单层面进账的钱，广告费和仓储费是月底另外从同一个 MercadoPago 余额里扣走的，"
          "不在这个数里。%s要得到利润，必须做下面这个桥。"
          % ("同时它扣掉了代扣代缴税金——那是预缴税款不是费用。" if site.WITHHOLDING else
             "同时它把退货处理费并在佣金列里扣掉了，而本报表按账单单列退货处理费（D3），所以先加回 D2 再扣 D3，避免扣两次。"))
    head()
    d1 = line("D1", "销售报表 “Total” 到账净额", V(lambda S: S["m"]["total"]),
              "桥 A 的结论。这是订单层面真正进你账户的钱，不是利润。", kind="s")
    line("D1b", "＋ 加回部分退款被重复扣的佣金",
         V(lambda S: S["cn"]["partial_total"]),
         "销售报表对部分退款订单把原佣金重复扣了一次，Total 因此少记。逐单用 "
         "MercadoPago 现金流水验证过：正常订单里 93%%–97%% 的 Total 与实际到账"
         "分毫不差，而这些单**每一笔**都正好差一个原佣金。明细见【%s】3.3。"
         % SHEET_NAMES[9])
    if site.WITHHOLDING:
        line("D2", "＋ 加回平台代扣代缴税金", V(lambda S: S["m"]["tax_ret"]),
             "预缴给税局的钱，能抵扣就不是成本，所以算利润时要加回来（若不能抵扣，见【%s】敏感性行）" % SHEET_NAMES[1])
        line("D3", "− 退货处理费", V(lambda S: -S["m"]["dev_net"]),
             "Cargo por devolución，在退款交易里单独扣，不在 Total 内")
    else:
        line("D2", "＋ 加回佣金列中的退货处理费等（非税）", V(lambda S: S["m"]["tax_ret"]),
             "销售报表佣金列 − 账单纯佣金（＝桥 A 的 A5）。不是税：智利平台不代扣。加回后由下一行按账单口径扣退货处理费，"
             "两行之差就是桥 E 里的时间/关联差")
        line("D3", "− 退货处理费", V(lambda S: -S["m"]["dev_net"]),
             "Cargo por devolución（账单口径，按订单关联跨账期取数）")
    line("D4", "− 广告费（Product Ads ＋ Display Ads ＋ 粉丝推广，净）", V(lambda S: -S["m"]["ads_net"]),
         "★ 完全不在销售报表里。账单 Descontado=No，月底自动扣款")
    line("D5", "− Full 仓相关费用（仓储/揽收/仓位/罚金等）", V(lambda S: -S["m"]["full_bill_total"]),
         "★ 同样不在销售报表里。%s"
         % ("已与 Reporte_Cargos_Full 逐项核对，差异 %s"
            % fmt(mm["cargos_full_report_total"] - mm["full_bill_total"])
            if M["files"].get("cargos_full") else
            "本期未下载到 Full 费用明细报表，无法交叉核对" if site.HAS_FULL else
            "%s没有 Full 仓，恒为 0" % site.COUNTRY_CN))
    line("D6", "− 官方店页面维护费", V(lambda S: -S["m"]["subscription_net"]), "Mi página 月费")
    # 新增的 "misc" 大类必须在桥 D 里有一行，否则 D7 的合计会和损益表对不上。
    line("D6a", "− 其他账单费用（信誉保障计划等）", V(lambda S: -S["m"]["misc_net"]),
         "已登记、但不属于广告 / Full / 页面维护的账单级费用")
    d6b = line("D6b", "− 未登记科目（订单级＋账单级）",
               V(lambda S: -(S["m"]["unc_order"] + S["m"]["unc_bill"])),
               "字典外的费用。为 0 说明本期所有科目都已登记；不为 0 时见【%s】。" % SHEET_NAMES[9])
    d7 = line("D7", "＝ 经营贡献毛利（未扣商品成本）",
              lambda c, j: "=SUM(%s%d:%s%d)" % (c, d1, c, d6b),
              "与【%s】的同名行完全一致" % SHEET_NAMES[1], kind="t")
    d8 = line("D8", "− 商品采购成本", [0.0] * n,
              "⚠ 平台数据不含此项。在【%s】填入单位成本后即可得出真实净利润。" % SHEET_NAMES[6], kind="s")
    for j in range(n):
        c = ws.cell(row=d8, column=sh.c0 + j)
        c.fill = fill(C_INPUT)
        c.font = F(10, True, "0000FF")
    line("D9", "＝ 税前净利润", lambda c, j: "=%s%d-%s%d" % (c, d7, c, d8), "", kind="chk")

    gap()
    contrib_m = ctx["contrib"][M["store"]]
    banner("一句话：到账净额是「现金」，经营贡献毛利才是「利润」。%s 本期到账 %s，"
           "但%s扣掉广告和仓储之后，真正的经营贡献是 %s；两者差 %s%s——这就是只看到账额会误判的地方。"
           % (M["store"], fmt(mm["total"]), "加回代扣税、再" if site.WITHHOLDING else "",
              fmt(contrib_m), fmt(abs(contrib_m - mm["total"])),
              "，方向还相反" if contrib_m > mm["total"] else ""))

    if not site.WITHHOLDING:
        _bridge_tax_no_withholding(stores, ctx, ws, H, n, banner)
    else:
        # ───────── 桥 E ─────────
        gap()
        block("桥 E　“平台代扣代缴税金”到底是什么，为什么算利润时要加回来",
              "这是本报表唯一一项“从现金里扣了、但不算费用”的支出。下面是它在原始数据里的准确形态与验证过程。")
        head()
        e1 = line("E1", "本期被代扣金额", V(lambda S: S["m"]["tax_ret"]),
                  "MercadoPago 结算流水里有独立字段 TAXES_AMOUNT，与手续费 FEE_AMOUNT 分开列示，"
                  "说明平台自己也不把它当手续费")
        mp = next((S for S in stores if S.get("mp") and S["mp"].get("rate_n")), None)
        rate_note = "本期各店实测：%s。" % per_store(
            stores, lambda S: "%s %s" % (S["store"], pct(S["m"]["tax_rate_gross"], 3)), "、")
        if mp:
            rate_note += ("%s 的 %d 笔结算流水中有 %d 笔费率完全一致，标准差 %.4f —— 费率高度统一，是法定预扣不是浮动费用。"
                          % (mp["store"], mp["mp"]["rate_n"], mp["mp"]["rate_mode_n"], mp["mp"]["rate_std"]))
        line("E2", "占含税成交额比例",
             lambda c, j: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, H["a1"], c, e1, c, H["a1"]),
             rate_note, total=False, fmt_code=PCT3)
        line("E3", "占不含税价比例（÷1.16 后）",
             lambda c, j: "=IF(%s%d=0,0,%s%d/(%s%d/(1+%s)))" % (c, H["a1"], c, e1, c, H["a1"], site.IVA_RATE),
             "本期各店实测 %s。按墨西哥数字平台代扣制度，最可能的构成是 IVA 预扣 8%% ＋ ISR 预扣 2.5%%，"
             "但准确拆分必须以 MercadoPago 每月出具的 Constancia de retenciones de IVA e ISR 为准 —— "
             "该凭证不在本次下载的文件里。" % per_store(
                 stores, lambda S: "%s %s" % (S["store"], pct(
                     S["m"]["tax_ret"] / S["m"]["net_revenue_exvat"] if S["m"]["net_revenue_exvat"] else 0, 2)), "、"),
             total=False, fmt_code=PCT3)
        line("E4", "为什么算利润时要加回", [None] * n,
             "预扣税是“预先替你缴给税局的税”，性质是预付税款（资产），不是经营成本。"
             "你在做 IVA/ISR 申报时可以拿它抵减应缴税额。若当成费用扣掉，等于同一笔税被计两次。",
             kind="s", total=False)

        # ───────── 桥 F 含税口径提示 ─────────
        gap()
        banner("⚠ 但要注意：本报表的收入是含税口径", C_WARN, F(11, True, C_RED))
        head()
        f1 = line("F1", "含税商品收入（报表中的 GMV）", V(lambda S: S["m"]["ing"]),
                  "买家实付金额，含 %s IVA" % pct(site.IVA_RATE, 0))
        f2 = line("F2", "其中 IVA 成分（× 16/116）",
                  lambda c, j: "=%s%d*%s/(1+%s)" % (c, f1, site.IVA_RATE, site.IVA_RATE),
                  "这部分钱在报表里被算进了收入，但它不属于你，最终要交给税局")
        line("F3", "不含税收入（参考）",
             lambda c, j: "=%s%d-%s%d" % (c, f1, c, f2),
             "若要编制正式的不含税损益表，应以此为收入起点", kind="t")
        f4 = line("F4", "已通过代扣预缴的部分", V(lambda S: S["m"]["tax_ret"]), "")
        line("F5", "预缴占 IVA 成分比例",
             lambda c, j: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, f2, c, f4, c, f2),
             "本期 %s。也就是说 IVA 还没缴完，剩余部分要在申报时结清（可用采购与费用的进项 IVA 抵扣）。"
             "★ 结论：加回代扣税只是把“不该当费用的”剔出去；要得到真正的税后净利，"
             "还需要用采购发票做完整的 IVA 销项—进项计算，这已超出平台数据能覆盖的范围，请交给你的墨西哥会计师。"
             % per_store(stores, lambda S: "%s %s" % (
                 S["store"], pct(S["m"]["tax_vs_iva"], 1)), "、"),
             total=False, fmt_code=PCT, kind="chk")

    # ───────── 常见误算 ─────────
    gap()
    block("常见误算示范（请避免）")
    tot = lambda f: sum(f(S) for S in stores)
    mis = [
        ("❌ 把“到账净额”当成利润",
         "到账净额里没有广告费(%s)、Full 仓费用(%s)和页面费(%s)，这些是月底另外扣的；"
         "同时它多扣了代扣税(%s)。见桥 D。"
         % (fmt(tot(lambda S: S["m"]["ads_net"])), fmt(tot(lambda S: S["m"]["full_bill_total"])),
            fmt(tot(lambda S: S["m"]["subscription_net"])), fmt(tot(lambda S: S["m"]["tax_ret"])))),
        ("❌ 把销售报表 Total 减去账单全部费用",
         "会把佣金和配送费重复扣一次（本期 %s），利润被严重低估。"
         % per_store(stores, lambda S: "%s 约 %s" % (S["store"], fmt(S["m"]["bill_by_channel"][CH_SI], 0)), "、")),
        ("❌ 把“Cargo por venta e impuestos”整笔当作平台佣金",
         "其中 %s 是代扣税，属预缴税款而非费用，会让佣金率虚高约 %s。"
         % (per_store(stores, lambda S: "%s %s" % (S["store"], fmt(S["m"]["tax_ret"])), " / "),
            pct(mm["tax_rate_gross"], 1))),
        ("❌ 用账单月份直接当作销售月份",
         "本期账单里混着上月订单的费用，本期订单也有一部分费用要到下月账单才出现。见桥 C。"),
        ("❌ 把代扣税当作平台费用扣掉",
         "它是预付税款不是成本，可用于抵减申报税额。当费用扣会让同一笔税被计两次，本期共 %s。"
         % fmt(tot(lambda S: S["m"]["tax_ret"]))),
        ("❌ 把含税收入直接当作利润起点",
         "GMV %s 里含 %s 的 IVA，那不是你的钱。本报表为便于与平台数据勾稽采用含税口径，正式申报需换算成不含税。"
         % (fmt(tot(lambda S: S["m"]["ing"])), fmt(tot(lambda S: S["m"]["iva_in_gmv"])))),
        ("❌ 只看销售报表算利润",
         "广告费、仓储费、揽收费等 %s（全部店铺合计）完全不在销售报表里。"
         % fmt(tot(lambda S: S["m"]["other_total"]))),
        ("❌ 退款只看销售报表的 Anulaciones",
         "平台退还的佣金/运费走贷记单（本期 %s），退货处理费又在账单里单独收，三处都要看。"
         % fmt(tot(lambda S: S["m"]["pay_nc"]))),
        ("❌ 把 Cargos_Full 报表与账单相加",
         "Full 费用报表是账单里 Full 科目的明细拆分，两者金额相同（本期差 %s）。相加即每笔仓储费重复计一次。"
         % fmt(tot(lambda S: S["m"]["cargos_full_report_total"] - S["m"]["full_bill_total"]))),
    ]
    if not site.WITHHOLDING:
        # 智利：把说"代扣税"的几条换成智利真正容易错的地方，其余照旧。
        mis[0] = ("❌ 把“到账净额”当成利润",
                  "到账净额里没有广告费(%s)、Full 仓费用(%s)和页面费(%s)，这些是月底另外扣的。见桥 D。"
                  % (fmt(tot(lambda S: S["m"]["ads_net"])), fmt(tot(lambda S: S["m"]["full_bill_total"])),
                     fmt(tot(lambda S: S["m"]["subscription_net"]))))
        mis[2] = ("❌ 把“Cargo por venta e impuestos”当成“佣金＋代扣税”",
                  "智利平台不代扣税：这一列在正常订单上就是纯佣金，退货订单上还并进了退货处理费"
                  "（本期差额 %s）。照墨西哥的做法把差额当税加回而不扣退货处理费，利润会被高估。见桥 E。"
                  % per_store(stores, lambda S: fmt(S["m"]["tax_ret"]), " / "))
    if not site.WITHHOLDING and any(cycle_day(S) != 1 for S in stores):
        mis[3] = ("❌ 拿一张账单当一个月",
                  "智利账单不按自然月出，%s：“%s”账单装的是 %s 的费用。直接拿一张账单当本月会错开好几天；"
                  "本报表按每一行的计费日把两张账单拆回自然月，订单级费用另按订单号跨账期关联。见桥 B、桥 C。"
                  % (("%s 按 %s" % (M["store"], cycle_cn(cycle_day(M)))), MONTH_FULL_ES[ctx["p0"].month],
                     "%s–%s" % (M["bill_cycles"][0][2].strftime("%m/%d"),
                                M["bill_cycles"][0][3].strftime("%m/%d"))))
    if not site.WITHHOLDING:
        mis[4] = ("❌ 以为 IVA 平台已经代缴了",
                  "智利平台不代扣（MercadoPago 流水 TAXES_AMOUNT 为 0），GMV 里的 IVA（本期 %s）全额由卖家自行申报。"
                  % fmt(tot(lambda S: S["m"]["iva_in_gmv"])))
    if not site.HAS_FULL:
        mis = [x for x in mis if "Cargos_Full" not in x[0]]
    for lab, note in mis:
        r = state["r"]
        ws.cell(row=r, column=2, value=lab).font = F(10, True, C_RED)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol_idx - 1)
        c = ws.cell(row=r, column=ncol_idx, value=note)
        c.font = F(9)
        c.alignment = wrap()
        for k in range(1, ncol_idx + 1):
            ws.cell(row=r, column=k).border = BOX
        ws.row_dimensions[r].height = est_height(note, 56)
        state["r"] = r + 1
    ws.freeze_panes = "%s5" % sh.scol(0)


def _bridge_tax_no_withholding(stores, ctx, ws, H, n, banner):
    """智利版桥 E / F。墨西哥的桥 E 讲"代扣税是什么、为什么加回"，桥 F 讲含税口径；
    智利没有代扣，这里改讲"佣金列里那块不是佣金的钱是什么"，并把含税口径的提示
    和需要会计判断的地方写清楚。行号与墨西哥不同，智利报表不和墨西哥的跨表对齐。"""
    block, head, line, gap = H["block"], H["head"], H["line"], H["gap"]
    V = lambda f: [f(S) for S in stores]

    gap()
    block("桥 E　智利没有代扣税：佣金列里不是佣金的那部分是什么",
          "销售报表的“Cargo por venta e impuestos”名字里有“impuestos”，墨西哥这一列确实含平台代扣的税；"
          "智利不是。下面三条证据都来自本期原始数据，⑩ 页的“平台未代扣税（智利）”校验每期自动复核。")
    head()
    line("E1", "销售报表佣金列 − 账单纯佣金", V(lambda S: S["m"]["tax_ret"]),
         "＝ 桥 A 的 A5 ＝ 桥 D 的 D2")
    line("E2", "其中：账单上的退货处理费 (Cargo por devolución)", V(lambda S: S["m"].get("vcol_dev", 0.0)),
         "按订单号关联到本期订单、跨账期取数。平台把它并进了销售报表的佣金列，损益表按账单单列")
    line("E3", "剩余：时间 / 关联差", V(lambda S: S["m"].get("vcol_rest", 0.0)),
         "退货处理费已进销售报表、账单要到下一期才出；多件包裹的佣金记在母行、账单记在子行等。"
         "本期占 GMV %s（校验上限 %s）。"
         % (per_store(stores, lambda S: pct(S["m"].get("vcol_rest", 0.0) / S["m"]["ing"]
                                             if S["m"]["ing"] else 0, 3), "、"),
            pct(CL_VCOL_REST_MAX, 1)), kind="s")
    line("E4", "佣金列与账单佣金逐单相等的订单数", V(lambda S: S["m"].get("com_match_n", 0)),
         "可比订单（佣金列不为 0、能按订单号挂上账单佣金）共 %s 单。墨西哥这两者逐单差约 9%%（代扣税），"
         "智利逐单相等，比值是 1.000 而不是 1.19 —— 既不是代扣，也不是在佣金上另加 IVA。"
         % per_store(stores, lambda S: "%d" % S["m"].get("com_match_of", 0), "、"), fmt_code=INT)
    line("E5", "MercadoPago 结算流水的税金字段 TAXES_AMOUNT",
         V(lambda S: (S.get("mp") or {}).get("tax")),
         "本月结算流水里平台代扣的税金合计。为 0 表示平台没有在打款时扣税；空白表示本店没有流水。")

    gap()
    banner("⚠ 但要注意：本报表的收入是含税口径（智利 IVA %s），而且平台不替你缴" % pct(site.IVA_RATE, 0),
           C_WARN, F(11, True, C_RED))
    head()
    f1 = line("F1", "含税商品收入（报表中的 GMV）", V(lambda S: S["m"]["ing"]),
              "买家实付金额，含 %s IVA" % pct(site.IVA_RATE, 0))
    f2 = line("F2", "其中 IVA 成分（× 19/119）",
              lambda c, j: "=%s%d*%s/(1+%s)" % (c, f1, site.IVA_RATE, site.IVA_RATE),
              "销项 IVA（débito fiscal）。平台不代扣，全额由卖家自行向 SII 申报缴纳")
    line("F3", "不含税收入（参考）",
         lambda c, j: "=%s%d-%s%d" % (c, f1, c, f2),
         "若要编制正式的不含税损益表，应以此为收入起点", kind="t")
    line("F4", "需要会计判断的地方", [None] * n,
         "⚠ 以下不是平台数据能回答的，请交智利会计师：① 销项 IVA 的申报（含开了 factura 的 B2B 订单、"
         "退货与取消的冲减）；② 平台费用（佣金、配送、广告%s）的发票里含不含 IVA、能否作为进项抵扣 —— "
         "平台导出的账单只给一个“Valor del cargo”，没有单列 IVA；③ 退货处理费、信誉保障计划扣款等"
         "科目的税务性质。本报表的经营贡献毛利是含税口径，没有扣任何税。" % ("、Full" if site.HAS_FULL else ""),
         kind="chk", total=False)


def sheet_bridge(wb, stores, ctx):
    sh, ws, H = _bridge_head(wb, stores, ctx)
    _bridge_tail(stores, ctx, sh, ws, H, len(stores))
