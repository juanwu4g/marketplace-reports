# -*- coding: utf-8 -*-
"""MercadoLibre report — ⑧ 月度趋势: the months in the Ventas window, commission by
month (billed or estimated) and the notes under the table."""

import pandas as pd

from core.report_family import Col, Family, MONEY, header_row
from core.report_kit import (C_AMBER, C_NAVY, C_RED, C_SUB, F, INT, PCT, est_height, fill,
    title_rows, wrap)

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (FEE_BY_ES, SHEET_NAMES, TAX_RATE_NOMINAL_GROSS,
    V_UNITS, fmt, fmt0, money, num, pct)


# ══════════════════════════════════════════════════════════════════════
# ⑧ 月度趋势
# ══════════════════════════════════════════════════════════════════════

# 账单关联覆盖率低于这个值就不信账单，改用税率推算 —— 半数以上订单挂不上时，
# 用挂上的那点钱当整月佣金会把代扣税撑得离谱。
TREND_LINK_MIN = 0.5
TREND_LINK_MIN_NO_WITHHOLDING = 0.95
SRC_BILLED = "账单实测"
SRC_RATED = "按本店实测税率推算"


def monthly_commission(S):
    """按**销售月**把 Ventas 的『佣金及税金』拆成 纯佣金 / 代扣代缴税金。

    返回 {ym: (纯佣金, 代扣税, 来源)}。

    账单只覆盖最近两三个计费期，销售报表却是滚动 6~13 个月，所以早期月份
    根本没有账单可挂。这些月份推的是**代扣税**而不是佣金 —— 代扣税是法定
    比例（各店实测 9.0x%，逐月几乎不动），佣金率却随品类在 11%~17% 之间
    跳动，推佣金等于编数字。推完再用 佣金 ＝ Ventas 合并值 − 代扣税 反求，
    两者之和因此永远等于平台给出的合并值，一分不多一分不少。
    """
    V = S["ventas_all"]
    V = V[V["ym"] != "NaT"]
    B = S["billing_all"]
    com = B[(~B["is_rev"]) & B["parent_detalle"].map(
        lambda d: FEE_BY_ES.get(str(d), (None, None, None))[1] == "commission").astype(bool)]
    ks_all = com["k_sale"].astype(str).str.strip()
    kp_all = com["k_pack"].astype(str).str.strip()

    out, measured, facts = {}, [], {}
    for ym, A in V.groupby("ym"):
        gmv = float(num(A, site.V_REVENUE).sum())
        vcom = float(-num(A, site.V_COMMISSION).sum())
        facts[ym] = (gmv, vcom)
        oids = set(A["oid"])
        ids = oids | set(A["pid"])
        ids.discard(""); ids.discard("nan")
        billed, hit = 0.0, set()
        for ks, kp, amt in zip(ks_all, kp_all, com["amount"]):
            k = ks if ks in ids else (kp if kp in ids else None)
            if k:
                billed += float(amt)
                hit.add(k)
        cov = len(hit & oids) / max(1, A["oid"].nunique())
        # 智利的账单不按自然月切，月末那几天的订单常常还在下一期没出的账单里：
        # 半数挂上就当实测，没挂上那部分的佣金会整笔被算成"佣金列口径差"
        # （DAB_CL 2026-10 只挂上一半，口径差虚高到 GMV 的 5%）。所以智利要求几乎
        # 全部挂上（95%）才算实测，否则按实测比例推。
        if cov >= (TREND_LINK_MIN if site.WITHHOLDING else TREND_LINK_MIN_NO_WITHHOLDING):
            out[ym] = (billed, vcom - billed, SRC_BILLED)
            if gmv > 0:
                measured.append((gmv, vcom - billed, billed))
        else:
            out[ym] = (None, None, None)      # 第二轮按税率补

    tot_gmv = sum(g for g, _, _ in measured)
    # 智利没有法定代扣率可作兜底：一个账单月都没有时按 0 推（整列算佣金）。
    rate = (sum(t for _, t, _ in measured) / tot_gmv) if tot_gmv > 0 else \
        (TAX_RATE_NOMINAL_GROSS if site.WITHHOLDING else 0.0)
    src = (SRC_RATED if site.WITHHOLDING else "按本店实测比例推算") + "（%.2f%%）" % (rate * 100)
    # 推算月份的哨兵：代扣税率一旦推高了，多算的税会**整笔落到佣金列**上，
    # 表现为佣金率异常偏低。只查偏低这一个方向 —— 佣金率偏高多半是品类结构
    # 变了（佣金按品类 13%~18%），本来就该逐月波动，当成异常只会天天误报。
    # 墨西哥站未完成税务登记的账号按更高比例代扣，新店早期几个月正是这种情况。
    crs = [b / g for g, _, b in measured if g > 0]
    floor = (min(crs) - 0.03) if crs else None
    for ym in out:
        if out[ym][0] is not None:
            continue
        gmv, vcom = facts[ym]
        tax = gmv * rate
        pure = vcom - tax
        note = src
        if site.WITHHOLDING and floor is not None and gmv > 0 and pure / gmv < floor:
            note = ("%s ⚠ 倒推出的佣金率只有 %s，远低于实测月份的 %s —— "
                    "说明当月实际代扣税率高于 %s（账号税务登记状态多半和现在不同），"
                    "本行两列的分摊不可用，合计仍准确"
                    % (src, pct(pure / gmv, 1), pct(min(crs), 1), pct(rate, 2)))
        out[ym] = (pure, tax, note)
    return out


def trend_table(S):
    V = S["ventas_all"].copy()
    V = V[V["ym"] != "NaT"]
    split = monthly_commission(S)
    g = V.groupby("ym").apply(lambda x: pd.Series({
        "订单数": len(x), "件数": num(x, V_UNITS).sum(),
        "GMV": num(x, site.V_REVENUE).sum(),
        "运费": -num(x, site.V_SHIP_COST).sum(),
        "退款": -num(x, site.V_REFUND).sum(),
        "券": x["coupon"].sum(),
        "净额": num(x, site.V_TOTAL).sum(),
        "异常": x["abnormal"].sum(),
    })).reset_index()
    g["佣金"] = g["ym"].map(lambda y: split[y][0])
    g["代扣税"] = g["ym"].map(lambda y: split[y][1])
    g["拆分来源"] = g["ym"].map(lambda y: split[y][2])
    return g


def trend_cols():
    """⑧ 的列（core/report_family.Col）。有两列的表头随站点变（代扣税 / 退货处理费等）。"""
    tax = site.TAX_SHORT.rstrip("等")
    f10 = F(10)
    return [
        Col("店铺", 16, font=f10), Col("月份", 10, font=f10),
        Col("订单数", 9, INT, font=f10), Col("销售件数", 10, INT, font=f10),
        Col("销售额 GMV", 15, MONEY, font=f10),
        Col("客单价", 11, MONEY, font=f10, formula="=IF(C{r}=0,0,E{r}/C{r})"),
        Col("平台佣金", 14, MONEY, font=f10), Col(site.TAXCOL_CN, 15, MONEY, font=f10),
        Col("平台配送费", 14, MONEY, font=f10),
        Col("佣金率", 9, PCT, font=f10, formula="=IF($E{r}=0,0,G{r}/$E{r})"),
        Col(tax + "率", 10, PCT, font=f10, formula="=IF($E{r}=0,0,H{r}/$E{r})"),
        Col("配送费率", 10, PCT, font=f10, formula="=IF($E{r}=0,0,I{r}/$E{r})"),
        Col("退款金额", 13, MONEY, font=f10),
        Col("退款率", 9, PCT, font=f10, formula="=IF(E{r}=0,0,M{r}/E{r})"),
        Col("卖家优惠券", 12, MONEY, font=f10), Col("售后订单数", 10, INT, font=f10),
        Col("售后率", 9, PCT, font=f10, formula="=IF(C{r}=0,0,P{r}/C{r})"),
        Col("到账净额", 15, MONEY, font=f10),
        Col("到账率", 9, PCT, font=f10, formula="=IF(E{r}=0,0,R{r}/E{r})"),
        Col("佣金/%s来源" % tax, 52, font=f10),
    ]


def sheet_trend(wb, stores, ctx):
    fam = Family(stores, money=site.MNY)
    ws = wb.create_sheet(SHEET_NAMES[7])
    title_rows(ws, "月度经营趋势",
               "数据源：%s（销售报表为滚动窗口导出，窗口两端的月份可能不完整）。"
               "本期为 %s，灰斜体行表示不完整月份，仅供参考。单位：%s"
               % (site.VENTAS_SHEET, ctx["period"], site.CUR))
    ws["A3"] = ("平台三项扣费各占一列，不合并。佣金与代扣代缴税金在平台报表里是一个数，"
                "本表按账单明细逐单还原；账单只覆盖最近两三个计费期，更早的月份改按本店"
                "实测代扣税率推算代扣税、再倒推佣金（两者之和始终等于平台给出的合并值）。"
                "每行末列注明这一行到底是实测还是推算。" if site.WITHHOLDING else
                "平台三项扣费各占一列，不合并。智利平台不代扣税；销售报表的佣金列在退货订单上"
                "并进了退货处理费，本表按账单明细逐单拆成“平台佣金”和“%s”两列。账单只覆盖"
                "最近两三个计费期，更早的月份按本店实测比例推算（两者之和始终等于平台给出的合并值）。"
                "每行末列注明这一行到底是实测还是推算。" % site.TAXCOL_CN)
    ws["A3"].font = F(9, False, C_NAVY)
    cols = trend_cols()
    header_row(ws, 4, [c.header for c in cols], [c.width for c in cols])
    NC = len(cols)

    r = 5
    for S in stores:
        for _, x in trend_table(S).iterrows():
            ym = str(x["ym"])
            cells = fam.write_row(ws, r, cols, (
                S["store"], ym, int(x["订单数"]), int(x["件数"]), money(x["GMV"]), None,
                money(x["佣金"]), money(x["代扣税"]), money(x["运费"]), None, None, None,
                money(x["退款"]), None, money(x["券"]), int(x["异常"]), None, money(x["净额"]), None,
                x["拆分来源"]))
            sc = cells[-1]
            if "⚠" in str(x["拆分来源"]):
                sc.font = F(9, True, C_RED)       # 推算与实测明显打架，别当成可用数字
            elif not str(x["拆分来源"]).startswith(SRC_BILLED):
                sc.font = F(9, False, C_AMBER)
            if ym in ctx["partial_months"]:       # 不完整月份：灰斜体
                for c in cells:
                    c.font = F(10, False, "808080", True)
            if ym == ctx["period"]:               # 本期：浅蓝加粗
                for c in cells:
                    c.fill = fill(C_SUB)
                    c.font = F(10, True)
            r += 1
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="趋势解读（按规则自动生成）").font = F(12, True, C_NAVY)
    r += 1
    for t in trend_notes(stores, ctx):
        c = ws.cell(row=r, column=1, value="• " + t)
        c.font = F(10)
        c.alignment = wrap()
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        ws.row_dimensions[r].height = est_height(t, 120, 15, 15)
        r += 1
    ws.freeze_panes = "C5"


def trend_notes(stores, ctx):
    """确定性的趋势结论。规则固定，只有数字随月份变化。"""
    out = []
    for S in stores:
        g = trend_table(S)
        g = g[~g["ym"].isin(ctx["partial_months"])]
        if len(g) < 2:
            out.append("%s 的销售报表窗口内只有 %d 个完整月份，暂时无法判断趋势；"
                       "要看真实走势需要按月累积导出并保存历史。" % (S["store"], len(g)))
            continue
        cur = g[g["ym"] == ctx["period"]]
        if not len(cur):
            continue
        cur = cur.iloc[0]
        peak = g.loc[g["GMV"].idxmax()]
        if peak["ym"] != ctx["period"] and peak["GMV"] > 0:
            chg = (cur["GMV"] - peak["GMV"]) / peak["GMV"]
            out.append("%s 的 GMV 从 %s 的峰值 %s 变化到本期 %s（%s），订单数由 %s 变为 %s。"
                       % (S["store"], peak["ym"], fmt(peak["GMV"]), fmt(cur["GMV"]),
                          pct(chg, 1), fmt0(peak["订单数"]), fmt0(cur["订单数"])))
        gm = g["GMV"].replace(0, float("nan"))
        # 三项费率分开看：佣金率随品类结构走，配送费率随客单价走，代扣税率是法定的，
        # 混在一起看不出是哪一项在动。
        billed = g[g["拆分来源"] == SRC_BILLED]
        cr = (g["佣金"] / gm).dropna()
        if len(cr):
            out.append("%s 的佣金率区间 %s–%s，本期 %s（账单实测的月份有 %d 个，"
                       "其余月份的佣金是用合并值减推算%s倒推的）。佣金率变动说明品类结构变了。"
                       % (S["store"], pct(cr.min(), 1), pct(cr.max(), 1),
                          pct(cur["佣金"] / cur["GMV"] if cur["GMV"] else 0, 1), len(billed),
                          site.TAX_SHORT))
        sr = (g["运费"] / gm).dropna()
        if len(sr):
            out.append("%s 的配送费率区间 %s–%s，本期 %s。每单配送费基本固定，"
                       "这个比率主要随客单价走 —— 客单价低则费率高。"
                       % (S["store"], pct(sr.min(), 1), pct(sr.max(), 1),
                          pct(cur["运费"] / cur["GMV"] if cur["GMV"] else 0, 1)))
        if len(billed) and not site.WITHHOLDING:
            tr = (billed["代扣税"] / billed["GMV"].replace(0, float("nan"))).dropna()
            if len(tr):
                out.append("%s 的“%s”占 GMV（账单实测月份）区间 %s–%s。智利平台不代扣税，"
                           "这一列只是并进佣金列的退货处理费，应接近 0；明显变大说明平台开始代扣"
                           "或佣金列混进了新的扣款，需要人工确认。"
                           % (S["store"], site.TAXCOL_CN, pct(tr.min(), 2), pct(tr.max(), 2)))
        elif len(billed):
            tr = (billed["代扣税"] / billed["GMV"].replace(0, float("nan"))).dropna()
            if len(tr):
                out.append("%s 的代扣代缴税率（账单实测月份）区间 %s–%s，法定名义值 %s。"
                           "越界说明税制或账号税务状态变了，需要人工确认。"
                           % (S["store"], pct(tr.min(), 2), pct(tr.max(), 2),
                              pct(TAX_RATE_NOMINAL_GROSS, 2)))
        rr = (g["退款"] / gm).dropna()
        if len(rr):
            worst = g.loc[rr.idxmax()]
            out.append("%s 的退款率峰值出现在 %s（%s），本期 %s；售后订单率本期 %s。"
                       % (S["store"], worst["ym"], pct(rr.max(), 2),
                          pct(cur["退款"] / cur["GMV"] if cur["GMV"] else 0, 2),
                          pct(cur["异常"] / cur["订单数"] if cur["订单数"] else 0, 2)))
        out.append("%s 的到账率（Ventas Total ÷ GMV）本期 %s，即每 100 元 GMV 实际到账约 %s 元，"
                   "其余是佣金、%s、配送费与退款。注意这个比率还没有扣广告费和仓储费。"
                   % (S["store"], pct(cur["净额"] / cur["GMV"] if cur["GMV"] else 0, 1),
                      fmt(cur["净额"] / cur["GMV"] * 100 if cur["GMV"] else 0, 0), site.TAX_SHORT))
    if len(stores) > 1:
        # key 不能省：两家店的比率完全相等时（最常见的是两家当月都没有销售，
        # 比率都是 0），元组比较会接着去比第二个元素 —— 两个 dict，直接
        # TypeError，整份汇总报表跟着没了。只按比率排，相等时保持传入顺序。
        rates = sorted((((S["m"]["com_imp"] + S["m"]["env_cost"]) / S["m"]["ing"] if S["m"]["ing"] else 0, S)
                        for S in stores), key=lambda t: t[0])
        lo, hi = rates[0], rates[-1]
        if hi[0] - lo[0] > 0.02:
            aov = lambda S: S["m"]["ing"] / S["m"]["rows"] if S["m"]["rows"] else 0
            out.append("横向比较：%s 的三项扣费合计占 GMV %s，明显高于 %s 的 %s，"
                       "主要因为客单价差异（%s vs %s）而每单固定配送费相近，配送费占比被拉高。"
                       % (hi[1]["store"], pct(hi[0], 1), lo[1]["store"], pct(lo[0], 1),
                          fmt(aov(hi[1]), 0), fmt(aov(lo[1]), 0)))
    return out
