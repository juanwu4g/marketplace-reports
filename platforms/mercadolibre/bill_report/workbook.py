# -*- coding: utf-8 -*-
"""The ten family sheets for one store (docs/REPORT_FAMILY.md), built with
core/report_family. Every sheet ties to the ledger: its totals end in rows
"残差（应为 0）" that are formulas against ② (so they stay 0 after a person
edits nothing, and show a number if anything is changed by hand).
"""
import collections

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from core import report_family as RF
from core import report_kit as K
from core import sku_costs as SC
from core.report_kit import F, fill, wrap
from platforms.mercadolibre.bill_report import catalog as C
from platforms.mercadolibre.bill_report import mapping as M
from platforms.mercadolibre.bill_report import status as ST

SHEETS = K.SHEET_NAMES
PL_SHEET = SHEETS[1]
TEXT9 = F(9)
GREY9 = F(9, False, K.C_GREY)
BOLD = F(10, True)


def q(sheet):
    return "'%s'" % sheet


def fmtm(x, dec=2):
    if x is None:
        return "—"
    return ("{:,.%df}" % dec).format(x)


def pct(x, dec=1):
    return ("{:.%df}%%" % dec).format(100.0 * x)


# ------------------------------------------------------------------ ② rows

def pl_rows(S):
    p = S["prof"]
    R = RF.PLRow
    rows = [
        R("gmv_billed", "商品销售额（本期计费的销售）", "Total de la venta · filas Cargo por venta", "d",
          "佣金行计费日在本月的销售，按账单上的“Total de la venta”（数量×单价，含 %s）；一笔销售一行，"
          "冲销后重新计费的按次计" % p.vat_name),
        R("rev_cancel", "减：发货前取消", "Anulación del cargo por venta", "d",
          "本月计费日的佣金冲销 → 冲销比例 × 订单金额（估算，见①五）；按证据归为发货前取消"),
        R("rev_return", "减：退货退款", "Anulación del cargo por venta", "d", "同上，证据显示已发货/退货"),
        R("rev_partial", "减：部分退款", "Anulación parcial", "d", "佣金只冲销一部分：按比例估算退款额"),
        R("rev_unproven", "减：其他冲销（状态未证实）", "Anulación del cargo por venta", "d",
          "有冲销、但找不到取消或退货的证据（见⑩）"),
        R("net_gmv", "净商品销售额", "", "f", "＝ 计费销售额 − 各类冲销",
          ("gmv_billed", "rev_cancel", "rev_return", "rev_partial", "rev_unproven")),
        R("ship_income", "加：买家支付的运费", "Envío a cargo del cliente", "d",
          "买家付给卖家的运费（随销售行列出）；平台的全额运费在下面“平台运费”里扣"),
        R("ship_income_rev", "减：冲销订单的买家运费", "Envío a cargo del cliente × 冲销比例", "d", ""),
        R("revenue", "营业收入合计", "", "sum", "", ("net_gmv", "ship_income", "ship_income_rev")),
    ]
    for key, lab, orig, note in C.PL_FEE_LINES:
        rows.append(R(key, lab, orig, "d", note))
    rows += [
        R("fees_total", "平台费用合计", "Σ Valor del cargo（本期）", "sum",
          "＝ 本期全部账单行（含贷记单）之和的相反数；与账本的勾稽见③桥 A", RF.Span("fee_commission", "fee_other")),
        R("contrib", "经营贡献（未扣商品成本）", "", "f", "＝ 营业收入 − 平台费用", ("revenue", "fees_total")),
        R("cogs", "减：商品采购成本", "", "in", "⚠ 账单不含商品成本：请录入（或配置 SKU 成本表）"),
        R("profit", "扣除商品成本后的贡献", "", "f",
          "未扣：MercadoPago 资金费用、卖家自担优惠、税金（见备查项与①六）", ("contrib", "-cogs")),
    ]
    return rows


# ------------------------------------------------------------------ build

def build(S, ctx, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    fam = RF.Family([S], money=K.money_format(S["prof"].decimals))
    ctx["money"] = fam.money
    sheet_readme(wb, fam, S, ctx)
    at = sheet_pl(wb, fam, S, ctx)
    ctx["pl_at"] = at
    sheet_bridge(wb, fam, S, ctx)
    sheet_fees(wb, fam, S, ctx)
    sheet_returns(wb, fam, S, ctx)
    sheet_qc(wb, fam, S, ctx)
    sheet_sku(wb, fam, S, ctx)
    sheet_trend(wb, fam, S, ctx)
    sheet_inventory(wb, fam, S, ctx)
    sheet_gaps(wb, fam, S, ctx)
    wb.save(out_path)
    return out_path


def plref(ctx, key):
    """='② 合并损益表'!C12 — the store's cell of a P&L row."""
    return "%s!C%d" % (q(PL_SHEET), ctx["pl_at"][key])


def write_cells(ws, r, c0, values, fmt=None, font=None, fill_color=None, border=True):
    for j, v in enumerate(values):
        c = ws.cell(row=r, column=c0 + j, value=v)
        if fmt:
            c.number_format = fmt
        if font is not None:
            c.font = font
        if fill_color:
            c.fill = fill(fill_color)
        if border:
            c.border = K.BOX


def section_head(ws, r, text, ncols):
    return RF.section(ws, r, text, ncols)


def resid_row(fam, ws, r, ncols, label, cells, label_col=1):
    """A '残差（应为 0）' row: cells = {col index: formula}."""
    ws.cell(row=r, column=label_col, value=label).font = BOLD
    for j in range(1, ncols + 1):
        ws.cell(row=r, column=j).fill = fill(K.C_TOTAL)
        ws.cell(row=r, column=j).border = K.BOX
    for col, f in cells.items():
        c = ws.cell(row=r, column=col, value=f)
        c.number_format = fam.money
        c.font = BOLD
    return r + 1


# ------------------------------------------------------------------ ①

def readme_sections(S, ctx):
    """①'s sections for one store; the combined workbook reuses the policy ones."""
    p, m, pl = S["prof"], S["m"], S["m"]["pl"]
    dec = p.decimals
    fac = [f for f in S["inv_facts"] if f["doc"] == "FAC"]
    cut_lines = []
    for f in sorted(fac, key=lambda f: f["first"] or S["p0"]):
        rows = S["P"][S["P"]["file"] == f["name"]]
        if len(rows):
            cut_lines.append("%s：计费日 %s → %s 的 %d 行（本文件全部计费日 %s → %s）" % (
                f["name"], rows["charge_date"].min(), rows["charge_date"].max(), len(rows), f["first"], f["last"]))
    ncs = S["P"][S["P"]["doc"] == "NC"]
    if len(ncs):
        cut_lines.append("贷记单：%s 的 %d 行（计费日在本月）" % (
            "、".join(sorted(set(ncs["file"]))), len(ncs)))
    used = [f for f in sorted(fac, key=lambda f: f["first"] or S["p0"])
            if len(S["P"][S["P"]["file"] == f["name"]])]
    mon = "%d 月" % S["p0"].month
    if len(used) > 1:
        period_text = ("%s的账单不是自然月（本店账单周期按文件里的计费日：%s）。按会计月（计费日 Fecha del cargo）取数："
                       "%s ＝ %s ＋ 计费日在本月的贷记单行。各张账单其余日期的行不计入本月。" % (
                           p.country_cn, "；".join("%s %s → %s（平台更新于 %s）" % (
                               f["bill_month"], f["first"], f["last"], f["updated"] or "?") for f in used),
                           mon, " ＋ ".join("%s 里 %s → %s 的行" % (
                               f["name"], S["P"].loc[S["P"]["file"] == f["name"], "charge_date"].min(),
                               S["P"].loc[S["P"]["file"] == f["name"], "charge_date"].max()) for f in used)))
    else:
        period_text = ("%s的账单按自然月出：%s ＝ %s ＋ 计费日在本月的贷记单行。之后的账单只用来看本月销售在月后的"
                       "冲销（期后事项，不计入本月）。" % (p.country_cn, mon, "、".join(f["name"] for f in used) or "（无）"))
    cost = S.get("cost")
    sections = [
        ("一、这份报表是什么", [
            ("基准", "以 MercadoLibre 的账单（Reporte_Facturacion ＋ Reporte_Notas_Credito）为唯一金额基准的财务报表。"
                     "账单的每一行（含贷记单）只计一次（按 Número del cargo），本月计费日的行合起来就是“账本”；"
                     "② 的平台费用 ＝ 账本，销售额也取自账单的销售行。其他文件只用来找 SKU、判断状态、把店铺级"
                     "费用落到 SKU，不提供金额。"),
            ("两份报表的关系", "每月与订单基准报表（%s_%s.xlsx）一起生成，两者各自独立计算、互不引用。会计入账以本报表（账单基准）为准；订单基准报表作经营分析参考。" % (S["store"], ctx["month"])),
            ("店铺 / 国家 / 币种", "%s ／ %s ／ %s（%s）" % (S["store"], p.country_cn, p.currency, p.currency_cn)),
            ("下载文件夹", S["folder"]),
        ]),
        ("二、期间：%s，按计费日" % ctx["period_cn"], [("会计月", period_text)] +
         [("本月取用的账单行", x) for x in cut_lines]),
        ("三、数据来源及其角色", [(name, "%s　｜　%s" % (role, facts)) for name, role, facts in S["files"]] +
         ([SC.source_line(ctx["sku_costs"])] if ctx.get("sku_costs") is not None else [])),
        ("四、本期要点", [
            ("计费销售", "%d 笔、%s 件，计费 GMV %s；冲销 %s（取消 %s、退货 %s、部分退款 %s、未证实 %s），"
                         "净商品销售额 %s。" % (
                             m["n_cohort"], fmtm(m["units_billed"], 0), fmtm(pl["gmv_billed"], dec),
                             fmtm(pl["net_gmv"] - pl["gmv_billed"], dec), fmtm(pl["rev_cancel"], dec),
                             fmtm(pl["rev_return"], dec), fmtm(pl["rev_partial"], dec),
                             fmtm(pl["rev_unproven"], dec), fmtm(pl["net_gmv"], dec))),
            ("平台费用", "%s，占净商品销售额 %s；其中佣金 %s、运费 %s、广告 %s、Full %s。" % (
                fmtm(-pl["fees_total"], dec), pct(-pl["fees_total"] / pl["net_gmv"]) if pl["net_gmv"] else "—",
                fmtm(-pl["fee_commission"], dec), fmtm(-pl["fee_shipping"], dec), fmtm(-pl["fee_ads"], dec),
                fmtm(-(pl["fee_full_storage"] + pl["fee_full_ops"]), dec))),
            ("经营贡献", "%s（未扣商品成本），占净商品销售额 %s。" % (
                fmtm(pl["contrib"], dec), pct(pl["contrib"] / pl["net_gmv"]) if pl["net_gmv"] else "—")),
            ("SKU 映射", "%s 的计费 GMV 已映射到 SKU（T1 %s、T1b %s、T2 %s、T3 %s；未匹配 %d 笔）。" % (
                pct(m["mapped_share"], 2), fmtm(m["coverage"][M.T1][1], dec), fmtm(m["coverage"][M.T1B][1], dec),
                fmtm(m["coverage"][M.T2U][1] + m["coverage"][M.T2M][1], dec), fmtm(m["coverage"][M.T3][1], dec),
                m["coverage"][M.NONE][0])),
            ("期后冲销", "本月计费、在 %s 之后才冲销的销售 %d 笔、%s：按计费日口径不计入本月，列在②备查项与⑤三-E。" % (
                S["p1"], m["post_n"], fmtm(-m["post_amount"], dec))),
        ] + ([SC.point_line([cost], pl["contrib"] - cost.cogs)] if cost is not None else [])),
        ("五、会计处理（口径）", [
            ("%s" % p.vat_name,
             "与正式报表同一口径：金额按 MercadoLibre 账单原值、含 %s（%s %s）。商品销售额、买家运费、平台费用"
             "都含税；所含 %s 只在②备查项列示（销项 ＝ 营业收入 × %s/(1+%s)；进项 ＝ 平台费用 × 同一比例），"
             "不参与损益计算。怎么申报、能抵多少，请交%s会计师按 %s 规定确认。" % (
                 p.vat_name, p.vat_name, pct(p.vat_rate, 0), p.vat_name, pct(p.vat_rate, 0),
                 pct(p.vat_rate, 0), p.country_cn, p.tax_auth)),
            ("代扣税", ("墨西哥平台通过 MercadoPago 代扣 IVA/ISR（retenciones）。代扣不在账单里，所以不在本基准内：②备查项"
                       "单列一行；有 MercadoPago 结算文件时附上其 TAXES_AMOUNT 合计作参考。它是预缴税款而非费用。")
             if p.withholding else ("%s平台不代扣税，账单里也没有代扣行。" % p.country_cn)),
            ("销售额", "一笔销售在它的“Cargo por venta”（佣金）行的计费日计入；金额取该行的 Total de la venta"
                       "（＝ 数量 × 单价，已逐行验证）。账单没有单独的销售收入行，佣金行就是销售的凭证。"),
            ("取消与退款（估算）", "账单不写退款金额，只写佣金冲销（Anulación del cargo por venta）。冲销比例 ＝ 冲销额 ÷ "
                                "原佣金（按 Cargo que bonifica/anula 找原行；找不到时按同一刊登的佣金率推算）；冲销销售额 "
                                "＝ 冲销比例 × 订单金额。比例 ≥ 99% 视为全额冲销（件数同时冲回），否则为部分退款（件数不冲回）。"
                                "MercadoPago 的实际退款只作旁证列在⑤，不替代账单口径 —— 两者的差主要是运费与时间差。"),
            ("取消还是退货", "按证据分类（⑤）：销售报表状态与发货/送达日期、ML 退货审核、账单的退货处理费、仓库质检、"
                          "MercadoPago 的状态与运单状态。没有任何证据时：冲销距下单 ≤ 3 天视为发货前取消（标“推定”），"
                          "否则列为“状态未证实”。"),
            ("期间归属", "一切按计费日：本月的冲销即使对应上月的销售也计入本月；本月销售在下月才冲销的不计入本月"
                       "（②备查项“期后冲销”）。"),
            ("买家运费", "账单销售行的“Envío a cargo del cliente”是买家付给卖家的运费，计入收入；平台向卖家收取的是全额运费"
                       "（Cargo por envíos），在费用里扣。冲销的订单按冲销比例冲回买家运费。"),
            ("费用归 SKU（政策）",
             "① 随订单的费用（佣金、联盟佣金、运费、退货处理费及其冲销）按订单号直接归该单的 SKU。"
             "② Full 揽收、取回、长龄库存费按 Full 费用明细的费用号（或账单行的刊登号）直接归 SKU。"
             "③ Full 仓储日租：账单某日的仓储费 ＝ 前一天的库存仓储费，按平台“仓储费报表”当天各 SKU 的金额比例归属"
             "（报表缺那一天时用最近一天的结构）。"
             "④ 其余店铺级费用（Product Ads、粉丝推广、Mi página、Full 超量/购买空间、入仓违约、信誉计划扣款）"
             "账单不指明商品，按各 SKU 净商品销售额（负数按 0）的占比分摊，⑦里单独成列，标“按销售额分摊”，"
             "与直接费用分开。"),
            ("商品成本", SC.policy_line("本期计费件数 − 发货前取消件数 − 退回后重新上架的件数（ML 审核“Apto para la "
                                      "venta”或质检“重新上架”）；其他退货、部分退款保留成本")[1]
             if ctx.get("sku_costs") is not None else
             "账单不含商品成本。②“商品采购成本”是黄色录入行；配置 SKU 成本表（config.json → reports.sku_costs）"
             "后按“计费件数 − 发货前取消 − 退回后重新上架件数”× 单位成本自动填入。"),
        ]),
        ("六、不在本基准内", [
            ("MercadoPago 资金费用", "分期手续费、提现、资金冻结与释放、MercadoPago 自身的费用都不在 MercadoLibre 账单里。"),
            ("卖家自担的优惠", "卖家自己出的折扣券、促销让利若已体现在单价里则已含在销售额；单独结算的部分账单不显示。"),
            ("平台补贴", "平台给卖家的运费补贴、现金返还（cashback）不在账单里（MercadoPago 有）。"),
            ("代扣税", "见五。" if p.withholding else "不适用（本国平台不代扣）。"),
            ("库存估值", "⑨只有数量；库存成本、跌价不在本报表。"),
        ]),
        ("七、各页说明", [
            (SHEETS[1], "损益：计费销售 → 冲销 → 净销售额 → 平台费用 → 经营贡献；行位置固定。下方费用率与备查项。"),
            (SHEETS[2], "勾稽：账本 ＝ ②费用、销售 → 净销售额、费用归属、SKU 映射覆盖、销售状态、每张账单 ↔ 付款报表、"
                        "贷记单 ↔ 付款报表、仓储费归属。每座桥以“残差（应为 0）”结束。"),
            (SHEETS[3], "账单里出现的每个科目：原文、中文、归入的②行、归 SKU 的方式、结算方式、本期金额。"),
            (SHEETS[4], "冲销汇总、本月计费销售的状态、逐单清单（取消 / 退货 / 部分退款 / 未证实 / 期后）、按 SKU 汇总。"),
            (SHEETS[5], "退货逐单：ML 审核结果、去向、仓库质检、退货处理费、MercadoPago 退款。"),
            (SHEETS[6], "表一每个 SKU 一行：件数、销售额、冲销、直接费用、分摊费用、贡献；表二每笔销售 → SKU 的映射。"),
            (SHEETS[7], "本月按周（按计费日）的销售、冲销、费用与贡献；销售额前 10 的 SKU 按周。"),
            (SHEETS[8], "Full 库存快照与本月净销量（只看数量，不进损益）。"),
            (SHEETS[9], "全部校验（error / warn / skipped）、未匹配的销售、待决事项。"),
        ]),
        ("八、局限", [
            ("退款金额是估算", "按佣金冲销比例推算；部分退款的比例来自佣金，若平台的佣金冲销不与退款同比例，会有偏差。"),
            ("状态证据", "销售报表只覆盖下载时往前约 2–6 个月，且会漏掉部分已取消的订单；没有 MercadoPago 文件的店铺，"
                        "取消/退货的判断更多依赖时滞推定。"),
            ("账单还会变", "下载之后平台还可能对本月的销售开出冲销（期后）；本月已截止的账单行不会再变。"),
        ]),
    ]
    return sections


# indices of readme_sections() that hold policy text, the same for every store
POLICY_SECTIONS = (4, 5, 7)


def sheet_readme(wb, fam, S, ctx):
    sections = readme_sections(S, ctx)
    bad = RF.failed_errors(ctx["checks"])
    parts = [x for x in (("本期有 %d 项 error 级校验未通过（见⑩）：相关数字在修好之前不能信" % len(bad)) if bad else "",
                         provisional_text([S])) if x]
    warning = "；".join(parts) + "。" if parts else None
    fam.readme_sheet(wb, SHEETS[0], "账单基准财务报表 · %s · %s" % (ctx["period_cn"], S["store"]),
                     "账单基准（会计入账口径）　生成于 %s　基准：MercadoLibre 账单 ＋ 贷记单（计费日）" % ctx["built"],
                     sections, warning=warning)


def provisional_text(Ss):
    """The 暂定 banner when a bill the month uses has not closed yet, else ""."""
    open_ = ["%s %s" % (S["store"], "、".join(S["provisional"])) if len(Ss) > 1 else "、".join(S["provisional"])
             for S in Ss if S.get("provisional")]
    if not open_:
        return ""
    return ("暂定：%s 尚未关账，本月落在其中的费用以后还可能变 —— 关账后在窗口点「只出报表」重出" % "；".join(open_))


# ------------------------------------------------------------------ ②

def sheet_pl(wb, fam, S, ctx):
    """② for one store or several (fam.stores; S is the first, for country facts)."""
    p = S["prof"]
    rows = pl_rows(S)
    sh, r, at, note_col = fam.pl_sheet(
        wb, SHEETS[1], rows, value=lambda X, key: X["m"]["pl"][key],
        title="合并损益表（%s · 账单基准）" % ctx["period_cn"],
        subtitle="口径：计费日在 %s → %s 的全部账单行（含贷记单）；金额含 %s，按 MercadoLibre 账单原值。单位：%s。"
                 "灰底=公式行，浅蓝=小计，黄底=需你录入。行位置每期固定。" % (S["p0"], S["p1"], p.vat_name, p.currency))
    costs = [X.get("cost") for X in fam.stores]
    if any(c is not None for c in costs):
        for i, cost in enumerate(costs):
            if cost is not None:
                sh.ws.cell(row=at["cogs"], column=sh.c0 + i).value = cost.cogs
        sh.ws.cell(row=at["cogs"], column=note_col).value = fam.per_store(
            lambda X: SC.pl_note(X["cost"]) if X.get("cost") is not None else "未配置成本")
    net = lambda X: X["m"]["pl"]["net_gmv"]
    share = lambda keys: (lambda X: -sum(X["m"]["pl"][k] for k in keys))
    r = fam.ratio_block(sh, r + 1, "费用率（占净商品销售额）", [
        ("销售佣金（净）", share(["fee_commission", "fee_affiliate"]), net),
        ("平台运费（净，未减买家运费）", share(["fee_shipping"]), net),
        ("退货处理费", share(["fee_return"]), net),
        ("广告", share(["fee_ads"]), net),
        ("Full 仓储与操作", share(["fee_full_storage", "fee_full_ops"]), net),
        ("平台费用合计", share(C.PL_FEE_KEYS), net),
        ("经营贡献率", lambda X: X["m"]["pl"]["contrib"], net)], note_col)
    vr = pct(p.vat_rate, 0)
    memo = [
        ("营业收入所含 %s（销项，参考）" % p.vat_name, "%s débito fiscal" % p.vat_name, lambda X: X["m"]["vat_out"],
         "＝ 营业收入 × %s/(1+%s)。仅供参考，不参与上表；B2B 发票、退货冲减的申报口径请会计确认" % (vr, vr)),
        ("平台费用所含 %s（进项，参考）" % p.vat_name, "%s acreditable" % p.vat_name, lambda X: X["m"]["vat_in"],
         "＝ 平台费用 × %s/(1+%s)；能否抵扣取决于平台开给本店的发票（factura / boleta）" % (vr, vr)),
    ]
    if p.withholding:
        memo.append(("平台代扣代缴税金（IVA/ISR）——不在本基准内", "Retenciones (MercadoPago)",
                     lambda X: X["m"]["mp_taxes"] or 0.0,
                     fam.per_store(lambda X: (
                         "账单不含代扣税。本行是 MercadoPago 结算流水里本月计费销售的 TAXES_AMOUNT 合计，仅作参考；"
                         "它是预缴税款而非费用") if X["m"]["mp_taxes"] is not None else
                         "账单不含代扣税；本店文件夹没有 MercadoPago 结算文件，无法给出参考值（显示 0）")))
    else:
        memo.append(("平台代扣税", "Retenciones", lambda X: 0.0, "%s平台不代扣；账单无此项" % p.country_cn))
    memo += [
        ("期后冲销：本月计费、月后才冲销的销售额（不计入本月）", "Anulación posterior al cierre",
         lambda X: -X["m"]["post_amount"], fam.per_store(lambda X: "%d 笔；对应佣金冲销 %s" % (
             X["m"]["post_n"], fmtm(-X["m"]["post_commission"], p.decimals))) + "。按计费日口径属于下月，明细见⑤三-E"),
        ("平台给的费用折扣（已含在费用净额中）", "Valor del descuento", lambda X: X["m"]["fee_discount"],
         "佣金、运费行上的折扣；上表的费用已是折后金额"),
        ("本期费用：从销售款中直接扣除", "Descontado de la operación = Si",
         lambda X: -X["m"]["deducted"].get("Si", 0.0), "随每笔收款在 MercadoPago 扣掉，不进月结账单"),
        ("本期费用：月结账单支付", "Descontado de la operación = No",
         lambda X: -X["m"]["deducted"].get("No", 0.0), "进月结账单，由“账单付款”支付（③桥 F）"),
        ("本期费用：冲销及其他", "Descontado de la operación = No aplica",
         lambda X: -X["m"]["deducted"].get("No aplica", 0.0) - X["m"]["deducted"].get("（空）", 0.0),
         "冲销行与被冲销的原行"),
        ("本期账单行中已由账单付款结清的金额", "Reporte_Pagos_Facturas", lambda X: X["m"]["paid_in_period"],
         "付款报表“Detalle de Pagos del mes”逐费用号对上的金额"),
        ("本期冲销对应的 MercadoPago 实际退款（旁证）", "settlement REFUND / collection",
         lambda X: X["m"]["mp_refund_in"], "与②冲销（估算）对照用；含运费、时间差，不替代账单口径"),
    ]
    fam.memo_block(sh, r + 1, "备查项（不参与上表计算）", memo, note_col)
    return at


# ------------------------------------------------------------------ ③

def sheet_bridge(wb, fam, S, ctx):
    p, m, pl = S["prof"], S["m"], S["m"]["pl"]
    Ln = RF.Line
    P_ = S["P"]
    A = []
    for name in sorted(set(P_["file"]), key=lambda n: (("Notas" in n), P_.loc[P_["file"] == n, "charge_date"].min())):
        rows = P_[P_["file"] == name]
        A.append(Ln("%s（计费日 %s → %s，%d 行）" % (name, rows["charge_date"].min(), rows["charge_date"].max(),
                                                len(rows)), (lambda X, rows=rows: float(rows["amount"].sum())),
                    "账单原符号：费用为正、冲销为负"))
    A += [Ln("账本合计（本月计费日的全部账单行）", lambda X: X["m"]["ledger"], "", "subtotal"),
          Ln("② 平台费用合计（取反）", lambda X: -X["m"]["pl"]["fees_total"], ""),
          Ln("残差（应为 0）", lambda X: round(X["m"]["ledger"] + X["m"]["pl"]["fees_total"], 6), "", "residual")]
    rev = lambda c: (lambda X: X["m"]["pl"]["rev_" + c])
    revs5 = sum(e["amount"] for e in S["revs"] if e["in_period"])
    B = [Ln("计费销售额（本月计费日的 Cargo por venta 行）", lambda X: X["m"]["pl"]["gmv_billed"],
            "%d 笔销售" % m["n_cohort"])]
    B += [Ln("减：%s" % ST.CLASS_CN[c], rev(c), "") for c in ST.REVERSAL_CLASSES]
    B += [Ln("净商品销售额（②）", lambda X: X["m"]["pl"]["net_gmv"], "", "subtotal"),
          Ln("⑦ 表一 各 SKU 净商品销售额合计", lambda X: X["m"]["sku_sum"]["net_gmv"], ""),
          Ln("残差（应为 0）：② − ⑦", lambda X: round(X["m"]["pl"]["net_gmv"] - X["m"]["sku_sum"]["net_gmv"], 6),
             "", "residual"),
          Ln("残差（应为 0）：② 冲销 − ⑤ 逐单冲销合计",
             lambda X: round(sum(X["m"]["pl"]["rev_" + c] for c in ST.REVERSAL_CLASSES) - revs5, 6), "",
             "residual")]
    att = S["attributed"]
    Cc = []
    for key, lab, _o, _n in C.PL_FEE_LINES:
        d, a = att.get((key, "direct"), 0.0), att.get((key, "alloc"), 0.0)
        if d or a or pl[key]:
            Cc.append(Ln("%s：直接 %s ＋ 分摊 %s" % (lab.replace("减：", ""), fmtm(d, p.decimals),
                                                   fmtm(a, p.decimals)), (lambda X, d=d, a=a: d + a), ""))
    Cc += [Ln("直接归 SKU 的费用合计", lambda X: sum(v for (l, h), v in att.items() if h == "direct"), "", "subtotal"),
           Ln("按销售额分摊的费用合计", lambda X: sum(v for (l, h), v in att.items() if h == "alloc"), "", "subtotal"),
           Ln("② 平台费用合计", lambda X: X["m"]["pl"]["fees_total"], ""),
           Ln("残差（应为 0）：直接 ＋ 分摊 − ②", lambda X: round(sum(att.values()) - X["m"]["pl"]["fees_total"], 6),
              "", "residual"),
           Ln("残差（应为 0）：⑦ 直接列 ＋ 分摊列 − ②", lambda X: round(
               X["m"]["sku_sum"]["direct"] + X["m"]["sku_sum"]["alloc"] - X["m"]["pl"]["fees_total"], 6), "",
              "residual")]
    cov = m["coverage"]
    D = [Ln("%s（%d 笔）" % (M.TIER_CN[t], cov[t][0]), (lambda X, t=t: X["m"]["coverage"][t][1]),
            "可信度：%s" % M.CONFIDENCE[t]) for t in M.TIER_ORDER]
    D += [Ln("合计", lambda X: sum(v[1] for v in X["m"]["coverage"].values()), "", "subtotal"),
          Ln("② 计费销售额", lambda X: X["m"]["pl"]["gmv_billed"], "覆盖率 %s（目标 ≥ 99%%）" % pct(m["mapped_share"], 2)),
          Ln("残差（应为 0）", lambda X: round(sum(v[1] for v in X["m"]["coverage"].values())
                                          - X["m"]["pl"]["gmv_billed"], 6), "", "residual")]
    cs = m["cohort_status"]
    E = [Ln("%s（%d 笔）" % (ST.CLASS_CN[c], cs[c][0]), (lambda X, c=c: X["m"]["cohort_status"][c][1]),
            "按本月计费的销售计（含期后冲销的证据）") for c in ST.CLASS_CN]
    E += [Ln("合计", lambda X: sum(v[1] for v in X["m"]["cohort_status"].values()), "", "subtotal"),
          Ln("残差（应为 0）：− ② 计费销售额", lambda X: round(sum(v[1] for v in X["m"]["cohort_status"].values())
                                                       - X["m"]["pl"]["gmv_billed"], 6), "", "residual")]
    Fb = []
    for d in S["docs"]:
        tag = "" if d["in_period"] else "（本期外，仅列示）"
        Fb += [Ln("%s%s：月结费用（Descontado = No）" % (d["name"], tag), (lambda X, d=d: d["invoiced"]),
                  "整张账单，非仅本月计费日"),
               Ln("　减：其后被冲销的部分（冲销行指向这些费用）", (lambda X, d=d: d["annulled"]), ""),
               Ln("　＝ 应付", (lambda X, d=d: d["expected"]), "", "subtotal"),
               Ln("　已付（付款报表逐费用号对上）", (lambda X, d=d: d["paid"]), ""),
               (Ln("　残差（应为 0）", (lambda X, d=d: round(d["expected"] - d["paid"], 6)), "已付清", "residual")
                if d["has_pay"] and abs(d["expected"] - d["paid"]) <= p.tol else
                Ln("　待付余额（账单未到期或未付清，不是差错）", (lambda X, d=d: round(d["expected"] - d["paid"], 6)),
                   "付款报表还没有这张账单的付款" if not d["has_pay"] else "部分已付", "pending"))]
    nm = S["nc_match"]
    G = [Ln("贷记单文件合计（全部，取正）", lambda X: X["m"]["nc_file_total"], ""),
         Ln("其中与付款报表对上的（%d 张）" % len(nm["matched"]), lambda X: sum(a for _, a, _b, _h in nm["matched"]),
            "按号码，其次按金额，待开票的按合计"),
         Ln("　付款报表上的对应金额", lambda X: sum(b for _, _a, b, _h in nm["matched"]), ""),
         Ln("　残差（应为 0）", lambda X: round(sum(a - b for _, a, b, _h in nm["matched"]), 6), "", "residual"),
         Ln("只在贷记单文件里的（%d 张）" % len(nm["nc_only"]), lambda X: sum(a for _, a in nm["nc_only"]),
            "；".join("%s %s" % (n, fmtm(a, p.decimals)) for n, a in nm["nc_only"][:10]) or "无", "pending"
            if nm["nc_only"] else None),
         Ln("只在付款报表里的（%d 张）" % len(nm["pagos_only"]), lambda X: sum(a for _, a in nm["pagos_only"]),
            "；".join("%s %s" % (n, fmtm(a, p.decimals)) for n, a in nm["pagos_only"][:10]) or "无", "pending"
            if nm["pagos_only"] else None)]
    ss = S["storage_stats"]
    H = [Ln("本期 Full 仓储日租（账单）", lambda X: ss["exact"] + ss["nearest"] + ss["none"], ""),
         Ln("　按当天的 SKU 仓储明细归属", lambda X: ss["exact"], "账单 d 日 ＝ 明细 d−1 日"),
         Ln("　按最近一天的 SKU 结构归属", lambda X: ss["nearest"], "明细未覆盖的日子（多为月初）"),
         Ln("　无明细，按销售额分摊", lambda X: ss["none"], ""),
         Ln("参考：明细当天合计 − 账单当天合计（同日对比）", lambda X: round(ss["report_exact"] - ss["exact"], 6),
            "平台两份报表的尾差，不影响归属（按比例分配账单金额）")]
    sections = [
        ("桥 A　账本 → ② 平台费用（%s）" % ("智利：9 月账单尾 ＋ 10 月账单头 ＋ 贷记单" if p.site == "cl" else
                                          "本月账单 ＋ 贷记单"), A),
        ("桥 B　计费销售 → 净商品销售额（②、⑤、⑦ 三处一致）", B),
        ("桥 C　平台费用的归属：直接归 SKU ＋ 按销售额分摊 ＝ ②", Cc),
        ("桥 D　SKU 映射覆盖（按计费销售额）", D),
        ("桥 E　本月计费销售的状态（⑤ 二）", E),
        ("桥 F　每张账单 ↔ 账单付款报表（Reporte_Pagos_Facturas）", Fb or [Ln("没有账单付款报表", lambda X: 0.0, "")]),
        ("桥 G　贷记单文件 ↔ 付款报表中的贷记单", G),
    ]
    if ss["exact"] + ss["nearest"] + ss["none"]:
        sections.append(("桥 H　Full 仓储日租按 SKU 明细归属", H))
    fam.bridge_sheet(wb, SHEETS[2], "平账勾稽表（%s · 账单基准）" % ctx["period_cn"],
                     "每座桥最后是“残差（应为 0）”；琥珀色 = 等待中的单据（账单未付、贷记单未开），不是错误。单位：%s" % p.currency,
                     sections)


# ------------------------------------------------------------------ ④

FEE_COLS = [RF.Col("科目（账单原文 Detalle）", 52, font=TEXT9, align=wrap()), RF.Col("中文", 28, font=TEXT9),
            RF.Col("归入②行", 22, font=TEXT9), RF.Col("归 SKU 方式", 30, font=TEXT9, align=wrap()),
            RF.Col("结算方式（Descontado）", 22, font=TEXT9), RF.Col("本期行数", 9, K.INT),
            RF.Col("本期金额（账单原符号）", 16, RF.MONEY), RF.Col("其中直接归 SKU", 15, RF.MONEY),
            RF.Col("其中按销售额分摊", 15, RF.MONEY), RF.Col("文件中本期外行数", 10, K.INT),
            RF.Col("本期外金额", 15, RF.MONEY)]


def sheet_fees(wb, fam, S, ctx):
    L, P_ = S["L"], S["P"]
    ws = fam.table_sheet(wb, SHEETS[3], "平台费用明细（%s · 账单基准）" % ctx["period_cn"],
                         "账单里出现的每个科目一行（含贷记单）；金额为账单原符号（费用为正、冲销为负）。"
                         "②的费用行 ＝ 本表本期金额按“归入②行”相加后取反。", FEE_COLS)
    r = 5
    dets = list(C.CATALOG.keys())
    seen = list(dict.fromkeys(list(L["detalle"])))
    order = [d for d in dets if d in seen] + [d for d in seen if d not in C.CATALOG]
    first = r
    settle = {"Si": "从销售款直接扣", "No": "月结账单支付", "No aplica": "冲销/不适用", "": "（空）"}
    for d in order:
        e, known = C.entry(d)
        rows_in = P_[P_["detalle"] == d]
        rows_out = L[(L["detalle"] == d) & (~L["in_period"])]
        how = "、".join(sorted(set(settle.get(x, x) for x in rows_in["deducted"]))) or "—"
        cells = fam.write_row(ws, r, FEE_COLS, [
            d, e.cn if known else "⚠ 未登记科目", C.FEE_CN[e.line], C.ATTRIB_CN[e.attrib], how,
            len(rows_in), float(rows_in["amount"].sum()), S["by_det"].get((d, "direct"), 0.0),
            S["by_det"].get((d, "alloc"), 0.0), len(rows_out), float(rows_out["amount"].sum())],
            fill_color=None if known else K.C_BAD)
        r += 1
    last = r - 1
    ws.cell(row=r, column=1, value="合计").font = BOLD
    for col in (6, 7, 8, 9, 10, 11):
        L_ = get_column_letter(col)
        c = ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (L_, first, L_, last))
        c.number_format = K.INT if col in (6, 10) else fam.money
        c.font = BOLD
        c.fill = fill(K.C_TOTAL)
        c.border = K.BOX
    tot = r
    r += 1
    r = resid_row(fam, ws, r, 11, "残差（应为 0）：本期金额 ＋ ②平台费用合计",
                  {7: "=ROUND(G%d+%s,6)" % (tot, plref(ctx, "fees_total")),
                   8: "=ROUND(H%d+I%d-G%d,6)" % (tot, tot, tot)})
    ws.cell(row=r - 1, column=10, value="H＋I−G：直接＋分摊＝本期").font = GREY9
    r += 1
    r = section_head(ws, r, "按②行汇总（与②逐行勾稽）", 11)
    hdr = ["②行", "本期账单金额（原符号）", "② 对应值", "残差（应为 0）"]
    K.header_cells(ws, r, hdr)
    r += 1
    for key, lab, _o, _n in C.PL_FEE_LINES:
        amt = float(P_.loc[P_["line"] == key, "amount"].sum())
        ws.cell(row=r, column=1, value=lab).border = K.BOX
        fam.money_cell(ws.cell(row=r, column=2), amt)
        fam.money_cell(ws.cell(row=r, column=3), "=%s" % plref(ctx, key))
        c = fam.money_cell(ws.cell(row=r, column=4), "=ROUND(B%d+C%d,6)" % (r, r))
        c.fill = fill(K.C_TOTAL)
        r += 1
    ws.freeze_panes = "B5"


# ------------------------------------------------------------------ ⑤

REV_COLS = [RF.Col("订单号（Número de venta）", 19), RF.Col("包裹号", 18), RF.Col("下单日", 11),
            RF.Col("冲销计费日", 11), RF.Col("SKU", 18), RF.Col("商品", 36, font=TEXT9),
            RF.Col("件数", 6, K.INT), RF.Col("订单金额", 13, RF.MONEY), RF.Col("冲销比例", 8, K.PCT),
            RF.Col("冲销销售额（估算，计入②）", 15, RF.MONEY), RF.Col("佣金冲销", 12, RF.MONEY),
            RF.Col("运费冲销（本期）", 12, RF.MONEY), RF.Col("退货处理费（本期，净）", 12, RF.MONEY),
            RF.Col("MercadoPago 实际退款（旁证）", 14, RF.MONEY), RF.Col("证据可信度", 9),
            RF.Col("状态证据", 70, font=TEXT9, align=wrap("top")),
            RF.Col("冲销比例算法", 40, font=GREY9, align=wrap("top"))]


def _rev_values(S, e):
    s = S["sales"].get(e["sale_no"], {})
    stt = S["status"][e["sale_no"]]
    fees = S["fees_by_sale"][e["sale_no"]]
    ship_rev = fees.get(("Anulación del cargo por envíos de Mercado Libre", "in" if e["in_period"] else "all"), 0.0)
    ret_fee = fees.get(("Cargo por devolución", "in" if e["in_period"] else "all"), 0.0) + \
        fees.get(("Anulación del cargo por devolución", "in" if e["in_period"] else "all"), 0.0)
    sku = s.get("sku", "")
    return [e["sale_no"], s.get("pack_no", ""), str(e["sale_date"] or ""), str(e["charge_date"] or ""), sku,
            S["sku_titles"].get(sku) or s.get("title", ""), e["qty"], e["total"], e["share"], e["amount"],
            e["commission_rev"], ship_rev, ret_fee, stt["ev"].mp_refund, stt["conf"], stt["basis"], e["method"]]


def sheet_returns(wb, fam, S, ctx):
    p, pl, m = S["prof"], S["m"]["pl"], S["m"]
    ws = wb.create_sheet(SHEETS[4])
    K.title_rows(ws, "退货退款分析（%s · 账单基准）" % ctx["period_cn"],
                 "金额来自账单：冲销销售额 ＝ 佣金冲销比例 × 订单金额（估算，②符号）；佣金/运费/退货费列为账单原符号"
                 "（费用为正、冲销为负）。MercadoPago 的实际退款只作旁证。"
                 "分类规则见①五。单位：%s" % p.currency)
    RF.set_widths(ws, [c.width for c in REV_COLS])
    r = 4
    # 一、汇总
    r = section_head(ws, r, "一、本期冲销汇总（计费日在本月，计入②）", len(REV_COLS))
    hdr = ["类别", "笔数", "冲销销售额（估算）", "佣金冲销", "MercadoPago 实际退款（旁证）", "② 对应值", "残差（应为 0）"]
    K.header_cells(ws, r, hdr)
    r += 1
    first = r
    keymap = {ST.CANCEL: "rev_cancel", ST.RETURN: "rev_return", ST.PARTIAL: "rev_partial",
              ST.UNPROVEN: "rev_unproven"}
    for c in ST.REVERSAL_CLASSES:
        ev = [e for e in S["revs"] if e["in_period"] and e["cls"] == c]
        mp = sum(S["status"][sn]["ev"].mp_refund for sn in set(e["sale_no"] for e in ev))
        write_cells(ws, r, 1, [ST.CLASS_CN[c]], font=F(10))
        write_cells(ws, r, 2, [len(ev)], fmt=K.INT)
        write_cells(ws, r, 3, [sum(e["amount"] for e in ev), sum(e["commission_rev"] for e in ev), mp,
                               "=%s" % plref(ctx, keymap[c])], fmt=fam.money)
        write_cells(ws, r, 7, ["=ROUND(C%d-F%d,6)" % (r, r)], fmt=fam.money, font=BOLD, fill_color=K.C_TOTAL)
        r += 1
    write_cells(ws, r, 1, ["合计"], font=BOLD, fill_color=K.C_TOTAL)
    write_cells(ws, r, 2, ["=SUM(B%d:B%d)" % (first, r - 1)], fmt=K.INT, font=BOLD, fill_color=K.C_TOTAL)
    write_cells(ws, r, 3, ["=SUM(%s%d:%s%d)" % (x, first, x, r - 1) for x in "CDEFG"], fmt=fam.money, font=BOLD,
                fill_color=K.C_TOTAL)
    r += 2
    # 二、状态
    r = section_head(ws, r, "二、本月计费销售的状态（每笔销售一行计；用到下载时的全部证据，含期后冲销）", len(REV_COLS))
    K.header_cells(ws, r, ["状态", "笔数", "计费销售额", "占比", "说明"])
    r += 1
    first = r
    notes = {ST.COMPLETED: "已送达（或纠纷结案平台把钱给了卖家），账单无冲销",
             ST.CANCEL: "佣金全额冲销，证据显示未发货", ST.RETURN: "佣金全额冲销，证据显示已发货/退货",
             ST.PARTIAL: "佣金部分冲销", ST.UNPROVEN: "佣金冲销，但无取消/退货证据",
             ST.OPEN: "在途、退货/纠纷处理中、或有退款证据但账单尚未冲销（见⑩）"}
    gmv = pl["gmv_billed"] or 1.0
    for c, (n, amt) in m["cohort_status"].items():
        write_cells(ws, r, 1, [ST.CLASS_CN[c]])
        write_cells(ws, r, 2, [n], fmt=K.INT)
        write_cells(ws, r, 3, [amt], fmt=fam.money)
        write_cells(ws, r, 4, [amt / gmv], fmt=K.PCT)
        write_cells(ws, r, 5, [notes[c]], font=GREY9)
        r += 1
    write_cells(ws, r, 1, ["合计"], font=BOLD, fill_color=K.C_TOTAL)
    write_cells(ws, r, 2, ["=SUM(B%d:B%d)" % (first, r - 1)], fmt=K.INT, font=BOLD, fill_color=K.C_TOTAL)
    write_cells(ws, r, 3, ["=SUM(C%d:C%d)" % (first, r - 1)], fmt=fam.money, font=BOLD, fill_color=K.C_TOTAL)
    tot = r
    r += 1
    r = resid_row(fam, ws, r, 5, "残差（应为 0）：合计 − ②计费销售额",
                  {3: "=ROUND(C%d-%s,6)" % (tot, plref(ctx, "gmv_billed"))})
    r += 1
    # 三–六 逐单
    lists = [
        ("A　发货前取消（本期冲销，计入②）", ST.CANCEL, "in"),
        ("B　退货退款（本期冲销，计入②）", ST.RETURN, "in"),
        ("C　部分退款（本期冲销，计入②）", ST.PARTIAL, "in"),
        ("D　已冲销·状态未证实（本期冲销，计入②）", ST.UNPROVEN, "in"),
        ("E　期后冲销：本月计费、月后才冲销（不计入本月，仅列示）", None, "post"),
    ]
    for title, cls, when in lists:
        if when == "in":
            ev = [e for e in S["revs"] if e["in_period"] and e["cls"] == cls]
        else:
            ev = [e for e in S["revs"] if e["post"] and S["sales"].get(e["sale_no"], {}).get("billed")]
        r = section_head(ws, r, "三、逐单 %s —— %d 笔" % (title, len(ev)), len(REV_COLS))
        K.header_cells(ws, r, [c.header for c in REV_COLS])
        ws.row_dimensions[r].height = 30
        r += 1
        first = r
        for e in sorted(ev, key=lambda e: (e["charge_date"] or S["p0"], e["sale_no"])):
            fam.write_row(ws, r, REV_COLS, _rev_values(S, e))
            r += 1
        write_cells(ws, r, 1, ["小计"], font=BOLD, fill_color=K.C_TOTAL)
        for col in (7, 8, 10, 11, 12, 13, 14):
            L_ = get_column_letter(col)
            write_cells(ws, r, col, ["=SUM(%s%d:%s%d)" % (L_, first, L_, r - 1) if r > first else 0],
                        fmt=K.INT if col == 7 else fam.money, font=BOLD, fill_color=K.C_TOTAL)
        sub = r
        r += 1
        if when == "in":
            r = resid_row(fam, ws, r, len(REV_COLS), "残差（应为 0）：小计 − ②",
                          {10: "=ROUND(J%d-%s,6)" % (sub, plref(ctx, keymap[cls]))})
        r += 1
    # 七、按 SKU
    r = section_head(ws, r, "四、按 SKU 汇总（本期冲销）", len(REV_COLS))
    K.header_cells(ws, r, ["SKU", "商品", "取消笔数", "取消金额", "退货笔数", "退货金额", "部分退款笔数",
                           "部分退款金额", "未证实笔数", "未证实金额", "冲销合计"])
    r += 1
    agg = {}
    for e in S["revs"]:
        if not e["in_period"]:
            continue
        sku = S["sales"].get(e["sale_no"], {}).get("sku", M.UNMAPPED_PREFIX + "无订单")
        a = agg.setdefault(sku, dict((c, [0, 0.0]) for c in ST.REVERSAL_CLASSES))
        a[e["cls"]][0] += 1
        a[e["cls"]][1] += e["amount"]
    first = r
    for sku, a in sorted(agg.items(), key=lambda kv: sum(v[1] for v in kv[1].values())):
        vals = [sku, S["sku_titles"].get(sku, "")]
        for c in ST.REVERSAL_CLASSES:
            vals += [a[c][0], a[c][1]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = K.BOX
            if j > 2:
                c.number_format = K.INT if j % 2 == 1 else fam.money
            if j == 2:
                c.font = TEXT9
        fam.money_cell(ws.cell(row=r, column=11), "=D%d+F%d+H%d+J%d" % (r, r, r, r))
        r += 1
    write_cells(ws, r, 1, ["合计"], font=BOLD, fill_color=K.C_TOTAL)
    for col in range(3, 12):
        L_ = get_column_letter(col)
        write_cells(ws, r, col, ["=SUM(%s%d:%s%d)" % (L_, first, L_, r - 1) if r > first else 0],
                    fmt=K.INT if col in (3, 5, 7, 9) else fam.money, font=BOLD, fill_color=K.C_TOTAL)
    tot = r
    r += 1
    resid_row(fam, ws, r, 11, "残差（应为 0）：冲销合计 − ②各冲销行",
              {11: "=ROUND(K%d-(%s),6)" % (tot, "+".join(plref(ctx, keymap[c]) for c in ST.REVERSAL_CLASSES))})
    ws.freeze_panes = "B4"


# ------------------------------------------------------------------ ⑥

QC_COLS = [RF.Col("订单号", 19), RF.Col("SKU", 18), RF.Col("商品", 34, font=TEXT9), RF.Col("下单日", 11),
           RF.Col("状态分类", 14), RF.Col("销售报表状态", 34, font=TEXT9, align=wrap()),
           RF.Col("ML 审核", 8), RF.Col("审核日", 16, font=TEXT9), RF.Col("审核结果", 18, font=TEXT9),
           RF.Col("去向", 12, font=TEXT9), RF.Col("原因", 24, font=TEXT9), RF.Col("款项归", 12, font=TEXT9),
           RF.Col("仓库质检（Returns）", 30, font=TEXT9, align=wrap()), RF.Col("重新上架？", 9),
           RF.Col("退货处理费（本期，净，②符号）", 14, RF.MONEY), RF.Col("本期冲销销售额", 14, RF.MONEY),
           RF.Col("MercadoPago 实际退款", 14, RF.MONEY), RF.Col("计入本期？", 12)]


def qc_rows(S):
    """⑥'s rows for one store, in display order."""
    rev_in = {}
    for e in S["revs"]:
        if e["in_period"]:
            rev_in[e["sale_no"]] = rev_in.get(e["sale_no"], 0.0) + e["amount"]
    sale_nos = set()
    for sn, f in S["fees_by_sale"].items():
        if f.get(("fee_return", "in"), 0.0):
            sale_nos.add(sn)
    for e in S["revs"]:
        if (e["in_period"] or e["post"]) and e["cls"] == ST.RETURN:
            sale_nos.add(e["sale_no"])
    rows = []
    for sn in sale_nos:
        s = S["sales"].get(sn, {})
        stt = S["status"].get(sn)
        if stt is None:
            continue
        v = stt["ev"].ventas
        t = stt["ev"].triage
        fee = -S["fees_by_sale"][sn].get(("fee_return", "in"), 0.0)
        in_p = bool(fee) or sn in rev_in
        rows.append([sn, s.get("sku", ""), S["sku_titles"].get(s.get("sku", ""), s.get("title", "")),
                     str(s.get("sale_date") or ""), ST.CLASS_CN[stt["cls"]],
                     v.estado if v is not None else "（销售报表无此单）",
                     v.reviewed if v is not None else "", v.review_date if v is not None else "",
                     v.result if v is not None else "", v.destination if v is not None else "",
                     v.result_reason if v is not None else "", v.money_to if v is not None else "",
                     ("%s｜%s｜%s" % (t.product_status, t.review_result, t.money_status)) if t is not None else "",
                     "是" if stt["back_to_stock"] else "", fee, rev_in.get(sn, 0.0), stt["ev"].mp_refund,
                     "是" if in_p else "否（期后）"])
    nos = -S["P"].loc[(S["P"]["line"] == "fee_return") & (S["P"]["sale_no"] == ""), "amount"].sum()
    if nos:
        rows.append(["（无订单号）", "", "", "", "", "", "", "", "", "", "", "", "", "", float(nos), 0.0, 0.0, "是"])
    return sorted(rows, key=lambda x: (x[-1] != "是", x[3], x[0]))


QC_SUBTITLE = ("每笔有退货迹象的销售一行：本期有退货处理费、本期被归为退货退款的冲销、或本月计费且在月后退货的。"
               "审核与质检来自销售报表与 Returns 文件，退货处理费与冲销来自账单。单位：%s")


def sheet_qc(wb, fam, S, ctx):
    p = S["prof"]
    ws = fam.table_sheet(wb, SHEETS[5], "退货质检明细（%s · 账单基准）" % ctx["period_cn"],
                         QC_SUBTITLE % p.currency, QC_COLS)
    r = 5
    first = r
    for vals in qc_rows(S):
        fam.write_row(ws, r, QC_COLS, vals)
        r += 1
    write_cells(ws, r, 1, ["合计"], font=BOLD, fill_color=K.C_TOTAL)
    for col in (15, 16, 17):
        L_ = get_column_letter(col)
        write_cells(ws, r, col, ["=SUM(%s%d:%s%d)" % (L_, first, L_, r - 1) if r > first else 0], fmt=fam.money,
                    font=BOLD, fill_color=K.C_TOTAL)
    tot = r
    r += 1
    resid_row(fam, ws, r, len(QC_COLS), "残差（应为 0）：退货处理费合计 − ②退货处理费",
              {15: "=ROUND(O%d-%s,6)" % (tot, plref(ctx, "fee_return"))})
    if not S["prof"].has_full:
        ws.cell(row=r + 2, column=1, value="%s没有 Full 仓：没有仓库质检文件，审核信息只来自销售报表。" %
                S["prof"].country_cn).font = GREY9
    ws.freeze_panes = "B5"


# ------------------------------------------------------------------ ⑦

def sku_cols(with_cost):
    mny = lambda h, w=13, **k: RF.Col(h, w, RF.MONEY, **k)
    calc = dict(fill=K.C_CALC)
    cols = [RF.Col("SKU", 20), RF.Col("商品", 36, font=TEXT9), RF.Col("映射级别（笔数）", 16, font=TEXT9),
            RF.Col("计费件数", 8, K.INT), RF.Col("冲销件数", 8, K.INT),
            RF.Col("净件数", 8, K.INT, formula="=D{r}-E{r}", **calc),
            mny("计费销售额"), mny("发货前取消"), mny("退货退款"), mny("部分退款"), mny("其他冲销"),
            mny("净商品销售额", formula="=G{r}+H{r}+I{r}+J{r}+K{r}", **calc), mny("买家运费（净）"),
            mny("直接：佣金"), mny("直接：联盟佣金"), mny("直接：平台运费"), mny("直接：退货处理费"),
            mny("直接：Full 仓储"), mny("直接：Full 揽收/取回"),
            mny("直接费用小计", formula="=SUM(N{r}:S{r})", **calc),
            mny("分摊：广告"), mny("分摊：Full 仓储类"), mny("分摊：Full 揽收/违约"), mny("分摊：店铺服务"),
            mny("分摊：其他"), mny("分摊费用小计", formula="=SUM(U{r}:Y{r})", **calc),
            mny("经营贡献（未扣成本）", 15, formula="=L{r}+M{r}+T{r}+Z{r}", **calc),
            RF.Col("贡献率", 8, K.PCT, formula="=IF(L{r}=0,0,AA{r}/L{r})", **calc)]
    if with_cost:
        cols += [RF.Col("单位成本", 11, RF.MONEY), RF.Col("成本件数", 8, K.INT),
                 mny("商品成本", formula="=-AC{r}*AD{r}", **calc),
                 mny("扣成本后贡献", formula="=AA{r}+AE{r}", **calc),
                 RF.Col("成本来源", 34, font=GREY9)]
    return cols


# ② key(s) each ⑦ money column must sum to
SKU_TIES = {"G": ("gmv_billed",), "H": ("rev_cancel",), "I": ("rev_return",), "J": ("rev_partial",),
            "K": ("rev_unproven",), "L": ("net_gmv",), "M": ("ship_income", "ship_income_rev"),
            "N": ("fee_commission",), "O": ("fee_affiliate",), "P": ("fee_shipping",), "Q": ("fee_return",),
            "AA": ("contrib",)}

MAP_COLS = [RF.Col("订单号（Número de venta）", 19), RF.Col("包裹号", 18), RF.Col("付款号（Pago）", 14),
            RF.Col("刊登号", 15), RF.Col("账单上的标题", 40, font=TEXT9), RF.Col("件数", 6, K.INT),
            RF.Col("订单金额", 13, RF.MONEY), RF.Col("佣金计费日", 11), RF.Col("SKU", 20),
            RF.Col("映射级别", 26, font=TEXT9), RF.Col("可信度", 7), RF.Col("依据", 40, font=GREY9),
            RF.Col("备注", 40, font=GREY9), RF.Col("状态分类", 14), RF.Col("本行用途", 16, font=GREY9)]


def sku_rows(S, cost):
    """⑦ 表一 for one store: [(SKU, values)] in display order (formula columns None)."""
    sk = S["sku"]
    out = []

    def sort_key(k):
        return (k.startswith("（"), k.startswith(M.UNMAPPED_PREFIX), -sk[k]["net_gmv"], k)
    for k in sorted(sk, key=sort_key):
        x = sk[k]
        tiers = S["sku_tiers"].get(k)
        tier_txt = "、".join("%s×%d" % (t, n) for t, n in sorted(tiers.items(), key=lambda kv: M.TIER_ORDER.index(
            kv[0]))) if tiers else "—"
        vals = [k, S["sku_titles"].get(k, ""), tier_txt, x["units"], x["units_rev"], None,
                x["gmv"], x["rev_cancel"], x["rev_return"], x["rev_partial"], x["rev_unproven"], None,
                x["ship_net"],
                x["d:fee_commission"], x["d:fee_affiliate"], x["d:fee_shipping"], x["d:fee_return"],
                x["d:fee_full_storage"], x["d:fee_full_ops"], None,
                x["a:fee_ads"], x["a:fee_full_storage"], x["a:fee_full_ops"], x["a:fee_store"], x["a:fee_other"],
                None, None, None]
        if cost is not None:
            line = cost.line(k)
            vals += [line.unit_cost if line is not None and line.unit_cost is not None else 0.0,
                     (line.cost_units if line is not None else 0), None, None, SC.status_text(line)]
        out.append((k, vals))
    return out


def sheet_sku(wb, fam, S, ctx):
    p = S["prof"]
    cost = S.get("cost")
    cols = sku_cols(cost is not None)
    ws = wb.create_sheet(SHEETS[6])
    K.title_rows(ws, "SKU 盈利分析（%s · 账单基准）" % ctx["period_cn"],
                 "表一：每个 SKU 一行。直接费用按订单/明细归属；“分摊”列是店铺级费用按净商品销售额占比分摊（政策见①五），"
                 "两者分开列。单位：%s" % p.currency,
                 "表二（下方）：每笔销售 → SKU 的映射，含级别与可信度。")
    RF.set_widths(ws, [c.width for c in cols])
    r = 4
    r = section_head(ws, r, "表一　SKU 盈利", len(cols))
    K.header_cells(ws, r, [c.header for c in cols])
    ws.row_dimensions[r].height = 42
    r += 1
    first = r
    for k, vals in sku_rows(S, cost):
        cells = fam.write_row(ws, r, cols, vals)
        if cost is not None and (cost.line(k) is None or cost.line(k).unit_cost is None):
            cells[28].fill = fill(K.C_INPUT)
        if k.startswith(M.UNMAPPED_PREFIX) or k.startswith("（"):
            cells[0].font = F(10, False, K.C_RED)
        r += 1
    last = r - 1
    write_cells(ws, r, 1, ["合计"], font=BOLD, fill_color=K.C_TOTAL)
    for j in range(4, len(cols) + 1):
        if j == 28 or (cost is not None and j in (29, 33)):
            continue
        L_ = get_column_letter(j)
        write_cells(ws, r, j, ["=SUM(%s%d:%s%d)" % (L_, first, L_, last)],
                    fmt=K.INT if j in (4, 5, 6, 30) else fam.money, font=BOLD, fill_color=K.C_TOTAL)
    tot = r
    ctx["sku_total_row"] = tot
    r += 1
    # ② values and residuals, column by column
    ws.cell(row=r, column=1, value="② 对应值").font = BOLD
    ties = dict(SKU_TIES)
    ties_extra = {"R+V": ("fee_full_storage",), "S+W": ("fee_full_ops",), "U": ("fee_ads",), "X": ("fee_store",),
                  "Y": ("fee_other",)}
    for col, keys in ties.items():
        c = ws.cell(row=r, column=_col_idx(col), value="=" + "+".join(plref(ctx, k) for k in keys))
        c.number_format = fam.money
        c.border = K.BOX
    r += 1
    cells = dict((_col_idx(col), "=ROUND(%s%d-%s%d,6)" % (col, tot, col, r - 1)) for col in ties)
    r = resid_row(fam, ws, r, len(cols), "残差（应为 0）：合计 − ②", cells)
    # the Full and allocated columns tie to ② lines in pairs
    ws.cell(row=r, column=1, value="残差（应为 0）：Full/分摊列 − ②").font = BOLD
    for colspec, keys in ties_extra.items():
        parts = colspec.split("+")
        f = "=ROUND(%s-(%s),6)" % ("+".join("%s%d" % (c, tot) for c in parts),
                                  "+".join(plref(ctx, k) for k in keys))
        c = ws.cell(row=r, column=_col_idx(parts[-1]), value=f)
        c.number_format = fam.money
        c.font = BOLD
        c.fill = fill(K.C_TOTAL)
        c.border = K.BOX
    ws.cell(row=r, column=3, value="R+V＝Full仓储类；S+W＝Full揽收/违约；U 广告；X 店铺服务；Y 其他").font = GREY9
    r += 2
    if cost is not None:
        ws.cell(row=r, column=1, value="商品成本：%s" % SC.pl_note(cost)).font = GREY9
        r += 2
    # 表二
    r = section_head(ws, r, "表二　账单销售 → SKU 映射（本月计费的销售，及本期冲销所涉的其他销售）", len(MAP_COLS))
    K.header_cells(ws, r, [c.header for c in MAP_COLS])
    ws.row_dimensions[r].height = 30
    r += 1
    for vals, unmapped in map_rows(S):
        fam.write_row(ws, r, MAP_COLS, vals, fill_color=K.C_BAD if unmapped else None)
        r += 1
    ws.freeze_panes = "B6"


def map_rows(S):
    """⑦ 表二 for one store: (values, unmapped?)."""
    listed = set()
    rows = []
    for ev in S["cohort"]:
        rows.append((ev, "本月计费（GMV）"))
        listed.add(ev["sale_no"])
    for e in S["revs"]:
        if e["in_period"] and e["sale_no"] not in listed and e["sale_no"] in S["sales"]:
            rows.append((S["sales"][e["sale_no"]], "仅本期冲销（上月销售）"))
            listed.add(e["sale_no"])
    out = []
    for s, use in rows:
        mp = S["map"][s["sale_no"]]
        out.append(([s["sale_no"], s["pack_no"], s["pago"], s["listing"], s["title"], s["qty"], s["total"],
                     str(s["charge_date"] or ""), s["sku"], M.TIER_CN[mp.tier], M.CONFIDENCE[mp.tier], mp.source,
                     mp.note, ST.CLASS_CN[S["status"][s["sale_no"]]["cls"]], use], mp.tier == M.NONE))
    return out


def _col_idx(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


# ------------------------------------------------------------------ ⑧

TREND_KEYS = ["fee_commission", "fee_affiliate", "fee_shipping", "fee_return", "fee_ads", "fee_full_storage",
              "fee_full_ops", "fee_store", "fee_other"]


def sheet_trend(wb, fam, S, ctx):
    p = S["prof"]
    cols = ([RF.Col("周（计费日）", 14), RF.Col("计费笔数", 8, K.INT), RF.Col("件数", 8, K.INT),
             RF.Col("计费销售额", 14, RF.MONEY), RF.Col("冲销", 13, RF.MONEY),
             RF.Col("净商品销售额", 14, RF.MONEY, formula="=D{r}+E{r}", fill=K.C_CALC),
             RF.Col("买家运费（净）", 12, RF.MONEY)] +
            [RF.Col(C.FEE_CN[k], 12, RF.MONEY) for k in TREND_KEYS] +
            [RF.Col("平台费用合计", 14, RF.MONEY, formula="=SUM(H{r}:P{r})", fill=K.C_CALC),
             RF.Col("经营贡献", 14, RF.MONEY, formula="=F{r}+G{r}+Q{r}", fill=K.C_CALC),
             RF.Col("费用率", 8, K.PCT, formula="=IF(F{r}=0,0,-Q{r}/F{r})", fill=K.C_CALC)])
    ws = fam.table_sheet(wb, SHEETS[7], "月度趋势（%s 按周 · 账单基准）" % ctx["period_cn"],
                         "一个会计月内按周（周一起，按计费日）。冲销按冲销的计费日、费用按各自的计费日落周。单位：%s" % p.currency,
                         cols)
    r = 5
    first = r
    for (a, b, label), x in zip(S["weeks"], S["weekly"]):
        fam.write_row(ws, r, cols, [label, x["n"], x["units"], x["gmv"], x["rev"], None, x["ship"]] +
                      [x[k] for k in TREND_KEYS] + [None, None, None])
        r += 1
    write_cells(ws, r, 1, ["合计"], font=BOLD, fill_color=K.C_TOTAL)
    for j in range(2, 19):
        L_ = get_column_letter(j)
        write_cells(ws, r, j, ["=SUM(%s%d:%s%d)" % (L_, first, L_, r - 1)],
                    fmt=K.INT if j in (2, 3) else fam.money, font=BOLD, fill_color=K.C_TOTAL)
    tot = r
    r += 1
    ties = {"D": ("gmv_billed",), "E": ("rev_cancel", "rev_return", "rev_partial", "rev_unproven"),
            "F": ("net_gmv",), "G": ("ship_income", "ship_income_rev"), "Q": ("fees_total",), "R": ("contrib",)}
    for k, col in zip(TREND_KEYS, "HIJKLMNOP"):
        ties[col] = (k,)
    ws.cell(row=r, column=1, value="② 对应值").font = BOLD
    for col, keys in ties.items():
        fam.money_cell(ws.cell(row=r, column=_col_idx(col)), "=" + "+".join(plref(ctx, k) for k in keys))
    r += 1
    r = resid_row(fam, ws, r, len(cols), "残差（应为 0）",
                  dict((_col_idx(c), "=ROUND(%s%d-%s%d,6)" % (c, tot, c, r - 1)) for c in ties))
    r += 1
    top = S["top_weekly"]
    if top:
        r = section_head(ws, r, "净商品销售额前 10 的 SKU，按周", len(cols))
        hdr = ["SKU", "商品"] + [w[2] for w in S["weeks"]] + ["本月合计"]
        K.header_cells(ws, r, hdr)
        r += 1
        for sku, vals in top:
            ws.cell(row=r, column=1, value=sku).border = K.BOX
            c = ws.cell(row=r, column=2, value=S["sku_titles"].get(sku, ""))
            c.font = TEXT9
            c.border = K.BOX
            for j, v in enumerate(vals, 3):
                fam.money_cell(ws.cell(row=r, column=j), v)
            fam.money_cell(ws.cell(row=r, column=3 + len(vals)), "=SUM(%s%d:%s%d)" % (
                get_column_letter(3), r, get_column_letter(2 + len(vals)), r))
            r += 1
    ws.freeze_panes = "B5"


# ------------------------------------------------------------------ ⑨

INV_SUBTITLE = ("库存是下载时的快照（%s），不是月末余额；对账列来自 conciliation（%s）。只看数量，不进损益。"
                "判断：可售 > 120 天 = 积压；< 14 天 = 补货；有销量无库存 = 断货。")


def inventory_cols(S):
    return [RF.Col("SKU", 20), RF.Col("商品", 40, font=TEXT9), RF.Col("可售（Buena calidad）", 11, K.INT),
            RF.Col("待促销（Para impulsar）", 11, K.INT), RF.Col("未上架（Para poner）", 11, K.INT),
            RF.Col("待处理（Para evitar descarte）", 11, K.INT),
            RF.Col("库存合计（快照）", 11, K.INT, formula="=SUM(C{r}:F{r})", fill=K.C_CALC),
            RF.Col("平台近 30 天销量", 10, K.INT), RF.Col("本月净件数（账单，⑦）", 11, K.INT),
            RF.Col("日均（本月）", 9, "0.00", formula="=I{r}/%d" % S["p1"].day, fill=K.C_CALC),
            RF.Col("可售天数", 9, K.INT, formula="=IF(J{r}<=0,IF(G{r}>0,999,0),G{r}/J{r})", fill=K.C_CALC),
            RF.Col("对账期初", 9, K.INT), RF.Col("入库", 8, K.INT), RF.Col("售出", 8, K.INT),
            RF.Col("买家退回", 8, K.INT), RF.Col("对账期末", 9, K.INT), RF.Col("判断", 14)]


def sheet_inventory(wb, fam, S, ctx):
    p = S["prof"]
    if not p.has_full:
        RF.not_applicable_sheet(wb, SHEETS[8], "库存与动销分析",
                                "不适用：%s没有 Full 仓，平台不提供库存报表（库存页 404）。本月净销量见⑦。" % p.country_cn)
        return
    cols = inventory_cols(S)
    ws = fam.table_sheet(wb, SHEETS[8], "库存与动销分析（Full 快照 vs %s 账单净销量）" % ctx["period_cn"],
                         INV_SUBTITLE % (S["stock_snapshot"] or "无快照文件", S["conciliation_range"] or "无对账文件"),
                         cols)
    r = 5
    first = r
    for vals, flag in inventory_rows(S):
        fam.write_row(ws, r, cols, vals, fill_color=K.C_WARN if flag else None)
        r += 1
    write_cells(ws, r, 1, ["合计"], font=BOLD, fill_color=K.C_TOTAL)
    for j in (3, 4, 5, 6, 7, 8, 9):
        L_ = get_column_letter(j)
        write_cells(ws, r, j, ["=SUM(%s%d:%s%d)" % (L_, first, L_, r - 1)], fmt=K.INT, font=BOLD,
                    fill_color=K.C_TOTAL)
    tot = r
    r += 1
    unm = unmapped_net_units(S)
    ws.cell(row=r, column=1, value="⑦ 净件数合计（不含未匹配 %s 件）" % fmtm(unm, 0)).font = BOLD
    c = ws.cell(row=r, column=9, value="=%s!F%d-(%.6f)" % (q(SHEETS[6]), ctx["sku_total_row"], unm))
    c.number_format = K.INT
    r += 1
    resid_row(fam, ws, r, len(cols), "残差（应为 0）", {9: "=ROUND(I%d-I%d,6)" % (tot, r - 1)})
    ws.freeze_panes = "B5"


def unmapped_net_units(S):
    """⑦'s net units that are not a real SKU (unmapped sales, pseudo rows)."""
    skus = [k for k in S["sku"] if not k.startswith("（") and not k.startswith(M.UNMAPPED_PREFIX)]
    return S["m"]["sku_sum"]["units"] - S["m"]["sku_sum"]["units_rev"] - sum(S["sku"][k]["units_net"] for k in skus)


def inventory_rows(S):
    """⑨'s rows for one store (formula columns None): [(values, flagged?)]."""
    stock = S["stock"]
    by = {}
    if stock is not None and len(stock):
        for r_ in stock.itertuples(index=False):
            x = by.setdefault(r_.sku, {"title": r_.title, "sold30": 0.0})
            x[r_.state] = x.get(r_.state, 0.0) + r_.units
            x["sold30"] = max(x["sold30"], r_.sold30)
    conc = {}
    if S["conciliation"] is not None and len(S["conciliation"]):
        for r_ in S["conciliation"].to_dict("records"):
            x = conc.setdefault(r_["sku"], collections.Counter())
            for k in ("start", "in", "sold", "returned", "end"):
                x[k] += r_[k]
    skus = [k for k in S["sku"] if not k.startswith("（") and not k.startswith(M.UNMAPPED_PREFIX)]
    allk = sorted(set(skus) | set(by) | set(conc), key=lambda k: (-S["sku"][k]["units_net"] if k in S["sku"] else 0, k))
    out = []
    for k in allk:
        x = by.get(k, {})
        c = conc.get(k, {})
        units = S["sku"][k]["units_net"] if k in S["sku"] else 0.0
        stock_total = sum(x.get(s, 0.0) for s in ("Buena calidad", "Para impulsar ventas", "Para poner en venta",
                                                    "Para evitar descarte"))
        daily = units / float(S["p1"].day)
        cover = (stock_total / daily) if daily > 0 else (999 if stock_total > 0 else 0)
        verdict = ("断货" if stock_total <= 0 and units > 0 else "积压" if cover > 120 else
                   "补货" if cover < 14 and units > 0 else "无库存无销量" if stock_total <= 0 else "正常")
        out.append(([
            k, x.get("title") or S["sku_titles"].get(k, ""), x.get("Buena calidad", 0.0),
            x.get("Para impulsar ventas", 0.0), x.get("Para poner en venta", 0.0),
            x.get("Para evitar descarte", 0.0), None, x.get("sold30", 0.0), units, None, None,
            c.get("start", 0.0), c.get("in", 0.0), c.get("sold", 0.0), c.get("returned", 0.0), c.get("end", 0.0),
            verdict], verdict in ("积压", "补货", "断货")))
    return out


# ------------------------------------------------------------------ ⑩

def sheet_gaps(wb, fam, S, ctx):
    p = S["prof"]
    unm = [s for s in S["cohort"] if S["map"][s["sale_no"]].tier == M.NONE]
    notes = []
    if unm:
        notes.append(("未匹配 SKU 的销售（%d 笔，%s）" % (len(unm), fmtm(sum(s["total"] for s in unm), p.decimals)), [
            (s["sale_no"], "刊登 %s｜%s｜%s｜%s" % (s["listing"], s["title"], fmtm(s["total"], p.decimals),
                                                 S["map"][s["sale_no"]].note)) for s in unm[:200]]))
    low = [s for s in S["cohort"] if S["map"][s["sale_no"]].tier in (M.T2M, M.T3)]
    if low:
        notes.append(("低可信度映射（%d 笔）——请抽查" % len(low), [
            (s["sale_no"], "%s → %s（%s；%s）" % (s["listing"], s["sku"], M.TIER_CN[S["map"][s["sale_no"]].tier],
                                                 S["map"][s["sale_no"]].note)) for s in low[:100]]))
    flagged = [(sn, S["status"][sn]) for sn in S["issues"]["refund_not_reversed"] + S["issues"]["cancel_not_reversed"]]
    if flagged:
        notes.append(("本月计费、有退款/取消证据但账单尚未冲销的销售（%d 笔）" % len(flagged), [
            (sn, st["basis"][:300]) for sn, st in flagged[:100]]))
    notes.append(("待你决定 / 会计判断", [
        ("冲销金额的口径", "本报表按“佣金冲销比例 × 订单金额”估算退款（账单口径）；MercadoPago 实际退款列为旁证。"
                        "若财务希望以实际退款为准，需要每家店都有 MercadoPago 文件。"),
        ("期后冲销", "本月计费、下月才冲销的销售按计费日留在本月 GMV。若希望按“销售所属月”回冲，需要等下月账单。"),
        ("%s" % p.vat_name, "金额含 %s；申报口径（B2B 发票、退货冲减、平台发票进项）交%s会计师确认。" % (
            p.vat_name, p.country_cn)),
        ("店铺级费用分摊", "广告等按净商品销售额分摊是一个可以改的政策；若要按“广告带来的销售”分摊，需要广告报表。"),
    ]))
    fam.gaps_sheet(wb, SHEETS[9], "数据缺口与待办（%s · 账单基准）" % ctx["period_cn"], ctx["checks"], notes=notes)

