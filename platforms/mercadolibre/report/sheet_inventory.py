# -*- coding: utf-8 -*-
"""MercadoLibre report — ⑨ 库存与动销分析: stock snapshot against the sales window."""

import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from core.report_family import Col, Family, header_row
from core.report_kit import (BOX, C_AMBER, C_BAD, C_BAND, C_GREEN, C_GREY, C_NAVY, C_RED, C_SUB,
    C_WARN, F, INT, fill, header_cells, title_rows)

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import SHEET_NAMES, V_SKU, V_UNITS, money, num, pct
from platforms.mercadolibre.report.loading import cell_int
from platforms.mercadolibre.report.sheet_common import no_full_text


# ══════════════════════════════════════════════════════════════════════
# ⑨ 库存与动销分析
# ══════════════════════════════════════════════════════════════════════

# 平台的"近 30 天销量"只给到某个时刻为止，我们只能按整天切窗口，边界上差
# 一两件属于正常。超过这个容差才算真的对不上。
STOCK_TOL_ABS = 2
STOCK_TOL_PCT = 0.05
STOCK_WINDOW_DAYS = 30
# 可售天数的判定线
# 90 天没到平台加收长龄费的门槛，报出来只是噪音；120 天正是 Full 开始按
# "库龄 4 个月以上"加收的那条线 —— 越过它，积压才真的开始花钱。
DAYS_URGENT, DAYS_PILED = 14, 120


def sales_window(S):
    """与平台"近 30 天"对齐的自算销量：{SKU: (件数, 金额)}。

    口径是**剔除异常订单后的全部销量**。实测过四种口径（全部 / 仅 Full /
    仅正常 / Full＋正常），"仅正常"最贴平台：603 个 SKU 里 563 个差不超过
    1 件；再加 Full 过滤反而更差，说明平台这个数并不只算 Full 渠道。
    """
    d = S.get("stock_snap")
    V = S["ventas_all"]
    if d is None or pd.isna(d) or not len(V):
        return {}, None, None
    w0 = d - pd.Timedelta(days=STOCK_WINDOW_DAYS)
    W = V[(V["fecha"] >= w0) & (V["fecha"] < d + pd.Timedelta(days=1))]
    W = W[~W["abnormal"]] if "abnormal" in W.columns else W
    out = {}
    if len(W):
        k = W[V_SKU].astype(str).str.strip()
        u = W.groupby(k)[V_UNITS].apply(lambda x: float(pd.to_numeric(x, errors="coerce").fillna(0).sum()))
        a = W.groupby(k).apply(lambda x: float(num(x, site.V_REVENUE).sum()))
        for sku in u.index:
            out[sku] = (float(u[sku]), float(a.get(sku, 0.0)))
    return out, w0, d


def stock_verdict(row, ours, tol):
    """一行 SKU 的核对结论 + 动销判定 + 需要处理的事。

    核对和动销分开说：核对回答"平台和我们的数对不对得上"，动销回答
    "这批货该补还是该清"。两件事混成一列的话，一个对不上的滞销 SKU
    只能显示其中一半。
    """
    plat = row["sales30_u"]
    diff = ours - plat
    if row["_no_sales_row"] and plat > 0:
        chk = "⚠ 销售报表里查无此 SKU"
    elif row["_no_stock_row"]:
        chk = "⚠ 库存报表里查无此 SKU（非 Full 或已下架）"
    elif abs(diff) <= tol:
        chk = "一致"
    elif diff > 0:
        chk = "⚠ 平台少记 %d 件" % round(diff)
    else:
        chk = "⚠ 平台多记 %d 件" % round(-diff)

    sell, total = row["sellable"], row["total_u"]
    days = (sell / (plat / STOCK_WINDOW_DAYS)) if plat > 0 else None
    acts = []
    if row["discard"] > 0:
        acts.append("面临丢弃 %d 件" % round(row["discard"]))
    if row["unsellable"] > 0:
        acts.append("不可售 %d 件" % round(row["unsellable"]))
    if row["aged"] > 0:
        acts.append("长龄 %d 件（加收仓储费）" % round(row["aged"]))
    if sell <= 0 and plat > 0:
        acts.append("已断货，近 30 天仍有 %d 件成交" % round(plat))
    elif plat <= 0 and sell > 0:
        acts.append("零动销，压 %d 件" % round(sell))
    elif days is not None and days < DAYS_URGENT:
        acts.append("可售仅 %.0f 天，需补货" % days)
    elif days is not None and days > DAYS_PILED:
        acts.append("可售 %.0f 天，库存积压" % days)
    if row["to_list"] > 0:
        acts.append("未上架 %d 件" % round(row["to_list"]))
    if row["in_transit"] > 0 and total > 0 and row["in_transit"] == total:
        # 整批刚发出去还没入仓，"零动销/积压"是它的必然状态，不是问题。
        # 归到"正常"，免得每批新到货都进待办清单把真问题淹掉。
        acts = [a for a in acts if "零动销" not in a and "积压" not in a
                and "断货" not in a]
        if not acts:
            return chk, days, "正常（%d 件在途，尚未入仓）" % round(row["in_transit"])
        acts.insert(0, "全部在途，尚未入仓")
    return chk, days, ("；".join(acts) if acts else "正常")


# 表一 · 各店 Full 仓库存结构（core/report_family.Col）。最后一列周转天数 = 平均库存
# ÷ 日均销量；销量为 0 时不写 0 —— 0 天会被读成"马上就卖完"，恰好与事实相反。
INV_TABLE1_COLS = (
    [Col("店铺", 18, font=F(10)), Col("快照日", 12, font=F(10))]
    + [Col(h, w, INT, font=F(10)) for h, w in (
        ("SKU 数", 8), ("在途待入库", 12), ("在转运", 10), ("买家退回", 11), ("可售", 10),
        ("不可售", 10), ("丢失/审核/取消", 14), ("库存合计", 11), ("好品质", 11),
        ("待促动销", 11), ("待上架", 10), ("待避免丢弃", 12), ("长龄单位", 11),
        ("平台30天销量", 13), ("自算30天销量", 13), ("销量差异", 11), ("近30天平均库存", 14))]
    + [Col("库存周转天数", 13, "0.0", font=F(10),
           formula='=IF(P{r}=0,"无动销",S{r}/(P{r}/%d))' % STOCK_WINDOW_DAYS)])
# 表三 · 需要处理的清单
INV_TABLE3_COLS = (
    [Col("店铺", None, font=F(9)), Col("SKU", None, font=F(9)), Col("商品名称", None, font=F(9)),
     Col("需要处理的事", None, font=F(9), align=Alignment(wrap_text=True, vertical="top")),
     Col("核对结论", None, font=F(9))]
    + [Col(h, None, INT, font=F(9)) for h in (
        "可售库存", "平台30天销量", "自算30天销量", "可售天数", "不可售+待丢弃")])


def sheet_inventory(wb, stores, ctx):
    fam = Family(stores, money=site.MNY)
    ws = wb.create_sheet(SHEET_NAMES[8])
    title_rows(ws, "库存与动销分析")
    if not site.HAS_FULL:
        ws["A2"] = no_full_text()
        ws["A2"].font = F(10, False, C_GREY)
        ws["A4"] = ("库存在卖家自己的仓库里，平台不提供库存快照。各 SKU 的销量见【%s】，"
                    "库存与周转请用自有仓库的进销存数据。" % SHEET_NAMES[6])
        ws["A4"].font = F(10, False, C_GREY)
        return
    have = [S for S in stores if len(S.get("stock", []))]
    none = [S["store"] for S in stores if not len(S.get("stock", []))]
    snaps = "；".join("%s 快照于 %s" % (S["store"],
                                    S["stock_snap"].strftime("%Y-%m-%d")
                                    if pd.notna(S.get("stock_snap")) else "日期未知")
                     for S in have)
    ws["A2"] = ("数据源：stock_general_full 报表的 Resumen 页（Full 仓库存）。%s%s"
                % (snaps, ("。%s 没有库存报表（需单独权限）。" % "、".join(none)) if none else ""))
    ws["A2"].font = F(9, False, C_GREY)
    ws["A3"] = ("⚠ 这是**下载当时的快照**，不是月末余额，与损益表不是同一个时点，"
                "两者不能相加。本页的“近 30 天”也是以快照日为终点，"
                "与会计月（%s）不重合。" % ctx["period"])
    ws["A3"].font = F(9, True, C_RED)
    ws["A4"] = ("销量核对：把平台给的“近 30 天销量”与我们从销售报表同窗口自算的件数逐 SKU 对比。"
                "自算口径为剔除异常订单后的全部销量。平台只给到快照那一刻，我们只能按整天切，"
                "所以差 %d 件以内（或 %s 以内）算一致，超出才标记。"
                % (STOCK_TOL_ABS, pct(STOCK_TOL_PCT, 0)))
    ws["A4"].font = F(9, False, C_NAVY)

    if not have:
        ws["A6"] = "本期没有任何店铺拉到库存报表，无法做库存分析。"
        ws["A6"].font = F(11, False, C_GREY)
        return

    # ---------------- 表一 · 各店库存结构 ----------------
    r = 6
    ws.cell(row=r, column=1, value="表一 · 各店 Full 仓库存结构").font = F(12, True, C_NAVY)
    r += 1
    t1 = INV_TABLE1_COLS
    header_row(ws, r, [c.header for c in t1], [c.width for c in t1])
    r += 1
    t1_start = r
    per_store = {}
    for S in have:
        ST = S["stock"]
        sold, w0, wd = sales_window(S)
        ours = sum(v[0] for v in sold.values())   # 含库存报表里没有的 SKU，差异才是全的
        per_store[S["store"]] = (sold, w0, wd)
        vals = [S["store"],
                S["stock_snap"].strftime("%Y-%m-%d") if pd.notna(S.get("stock_snap")) else "—",
                int(ST["sku"].nunique()),
                ST["in_transit"].sum(), ST["transfer"].sum(), ST["returned"].sum(),
                ST["sellable"].sum(), ST["unsellable"].sum(),
                ST["lost"].sum() + ST["in_review"].sum() + ST["cancelled_u"].sum(),
                ST["total_u"].sum(), ST["good"].sum(), ST["boost"].sum(),
                ST["to_list"].sum(), ST["discard"].sum(), ST["aged"].sum(),
                ST["sales30_u"].sum(), ours, ours - ST["sales30_u"].sum(),
                ST["avg_stock"].sum()]
        fam.write_row(ws, r, t1, [cell_int(v) for v in vals] + [None])
        r += 1
    if len(have) > 1:
        ws.cell(row=r, column=1, value="合计").font = F(10, True)
        for col in range(3, 20):
            cl = get_column_letter(col)
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, t1_start, cl, r - 1)
                    ).number_format = INT
        ws.cell(row=r, column=20,
                value="=IF(P%d=0,\"无动销\",S%d/(P%d/%d))" % (r, r, r, STOCK_WINDOW_DAYS)
                ).number_format = "0.0"
        for c in range(1, len(t1) + 1):
            ws.cell(row=r, column=c).fill = fill(C_SUB)
            ws.cell(row=r, column=c).font = F(10, True)
            ws.cell(row=r, column=c).border = BOX
        r += 1
    r += 2

    # ---------------- 表二 · 逐 SKU ----------------
    ws.cell(row=r, column=1, value="表二 · 逐 SKU 库存、动销与销量核对").font = F(12, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "“自算30天销量”来自销售报表同窗口、剔除异常订单后的件数。"
        "单位采购成本从【%s】自动带过来（在那一页填一次即可），用来估算可售库存占用的资金。"
        % SHEET_NAMES[6])).font = F(9, False, C_GREY)
    r += 1
    t2 = ["店铺", "SKU", "商品名称", "刊登状态", "在途待入库", "可售", "不可售",
          "丢失/审核/取消", "库存合计", "长龄单位", "好品质", "待促动销", "待上架",
          "待避免丢弃", "平台30天销量", "自算30天销量", "销量差异", "核对结论",
          "平台30天销售额", "本期会计月销量", "近30天平均库存", "可售天数",
          "平台售罄预估", "单位采购成本", "可售库存金额", "需要处理的事"]
    # 列宽由表一定下，本表不另设（表一之后的列用默认宽度）。
    header_cells(ws, r, t2, color=C_BAND)
    ws.row_dimensions[r].height = 32
    r += 1

    problems = []
    for S in have:
        ST = S["stock"].copy()
        sold, w0, wd = per_store[S["store"]]
        month_u = {}
        A = S["ventas"]
        if len(A):
            ok = A[~A["abnormal"]] if "abnormal" in A.columns else A
            month_u = ok.groupby(ok[V_SKU].astype(str).str.strip())[V_UNITS].apply(
                lambda x: float(pd.to_numeric(x, errors="coerce").fillna(0).sum())).to_dict()
        # 平台按**刊登**出一行，同一个 SKU 可能挂在好几个刊登上，先按 SKU 合并
        agg = ST.groupby("sku", as_index=False).agg(
            {c: "sum" for c in ("in_transit", "transfer", "returned", "sellable",
                                "unsellable", "lost", "in_review", "cancelled_u",
                                "total_u", "aged", "good", "boost", "to_list",
                                "discard", "avg_stock", "sales30_u", "sales30_amt")})
        first = ST.drop_duplicates("sku").set_index("sku")
        agg["title"] = agg["sku"].map(first["title"])
        agg["estado"] = agg["sku"].map(first["estado"])
        agg["runout"] = agg["sku"].map(first["runout"])
        agg["_no_stock_row"] = False
        # 销售报表里有、库存报表里没有的 SKU 也要露面，否则漏掉的正是最该查的
        extra = sorted(set(sold) - set(agg["sku"]) - {""})
        for sku in extra:
            row = {c: 0.0 for c in agg.columns if c not in ("sku", "title", "estado", "runout")}
            row.update({"sku": sku, "title": "（库存报表中无此 SKU）", "estado": "—",
                        "runout": "—", "_no_stock_row": True})
            agg = pd.concat([agg, pd.DataFrame([row])], ignore_index=True)
        agg = agg[agg["sku"].astype(str).str.strip() != ""]
        agg = agg.sort_values(["sales30_amt", "total_u"], ascending=[False, False])

        start = r
        for _, x in agg.iterrows():
            sku = str(x["sku"])
            ours = sold.get(sku, (0.0, 0.0))[0]
            x = dict(x)
            x["_no_sales_row"] = sku not in sold
            tol = max(STOCK_TOL_ABS, x["sales30_u"] * STOCK_TOL_PCT)
            chk, days, todo = stock_verdict(x, ours, tol)
            vals = [S["store"], sku, str(x["title"])[:45], x["estado"],
                    x["in_transit"], x["sellable"], x["unsellable"],
                    x["lost"] + x["in_review"] + x["cancelled_u"], x["total_u"],
                    x["aged"], x["good"], x["boost"], x["to_list"], x["discard"],
                    x["sales30_u"], ours, ours - x["sales30_u"]]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=cell_int(v))
                if i >= 5:
                    c.number_format = INT
            cc = ws.cell(row=r, column=18, value=chk)
            if chk != "一致":
                cc.font = F(9, True, C_AMBER)
            ws.cell(row=r, column=19, value=money(x["sales30_amt"])).number_format = site.MNY
            ws.cell(row=r, column=20, value=int(month_u.get(sku, 0.0))).number_format = INT
            ws.cell(row=r, column=21, value=int(x["avg_stock"])).number_format = INT
            ws.cell(row=r, column=22,
                    value="=IF(O%d=0,\"无动销\",F%d/(O%d/%d))" % (r, r, r, STOCK_WINDOW_DAYS)
                    ).number_format = "0.0"
            ws.cell(row=r, column=23, value=x["runout"])
            # 成本从 ⑦ 页表五带过来：那里是全页唯一的录入位，不要让人填第二遍。
            # 按【店铺＋SKU】两个键匹配 —— 汇总报表里同一个 SKU 会出现在好几家店。
            # 只认表五那一段，不能整列求和：⑦ 页的表一、表四同样有 M 列，会重复计。
            ref = ctx.get("sku_cost_ref")
            ws.cell(row=r, column=24, value=(
                "=SUMIFS('%s'!$M$%d:$M$%d,'%s'!$A$%d:$A$%d,$A%d,'%s'!$B$%d:$B$%d,$B%d)"
                % (ref[0], ref[1], ref[2], ref[0], ref[1], ref[2], r,
                   ref[0], ref[1], ref[2], r) if ref else 0)
                ).number_format = site.MNY
            ws.cell(row=r, column=25, value="=F%d*X%d" % (r, r)).number_format = site.MNY
            tc = ws.cell(row=r, column=26, value=todo)
            tc.alignment = Alignment(wrap_text=True, vertical="top")
            if not todo.startswith("正常"):
                tc.font = F(9, False, C_RED)
            if not todo.startswith("正常") or chk != "一致":
                problems.append((S["store"], sku, str(x["title"])[:40], todo, chk,
                                 x["sellable"], x["sales30_u"], ours,
                                 (0 if days is None else round(days)),
                                 x["discard"] + x["unsellable"]))
            for c in range(1, len(t2) + 1):
                ws.cell(row=r, column=c).border = BOX
                if ws.cell(row=r, column=c).font.color is None:
                    ws.cell(row=r, column=c).font = F(9)
            if x["_no_stock_row"] or x["_no_sales_row"]:
                for c in range(1, len(t2) + 1):
                    ws.cell(row=r, column=c).fill = fill(C_WARN)
            r += 1

        ws.cell(row=r, column=2, value="%s 合计" % S["store"]).font = F(10, True)
        for col in list(range(5, 18)) + [19, 20, 21, 25]:
            cl = get_column_letter(col)
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)
                    ).number_format = site.MNY if col in (19, 25) else INT
        for c in range(1, len(t2) + 1):
            ws.cell(row=r, column=c).fill = fill(C_SUB)
            ws.cell(row=r, column=c).font = F(10, True)
            ws.cell(row=r, column=c).border = BOX
        r += 2

    # ---------------- 表三 · 需要处理的清单 ----------------
    r += 1
    ws.cell(row=r, column=1, value="表三 · 需要处理的清单（按占用资金排序）").font = F(12, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "只列出表二里需要处理、或与平台对不上的 SKU。顺序：断货 → 面临丢弃 → "
        "数量对不上 → 其余。断货是在丢确定的销量，面临丢弃是已经确定要亏的货，"
        "数量对不上则说明 SKU 映射或渠道口径有问题，会影响后面所有按 SKU 的分析。"
        )).font = F(9, False, C_GREY)
    r += 1
    t3 = INV_TABLE3_COLS
    header_cells(ws, r, [c.header for c in t3], color=C_BAND)
    r += 1
    if not problems:
        ws.cell(row=r, column=1, value="本期库存没有需要处理的项。").font = F(10, False, C_GREEN)
        r += 1
    else:
        # 断货 > 面临丢弃 > 其他；同级按可售库存量
        rank = lambda p: (0 if "断货" in p[3] else 1 if "丢弃" in p[3] else
                          2 if "⚠" in p[4] else 3, -p[5])
        for p in sorted(problems, key=rank):
            fam.write_row(ws, r, t3, list(p[:5]) + [int(v) for v in p[5:]],
                          fill_color=C_BAD if "断货" in p[3] else None)     # 断货标红
            r += 1
    ws.freeze_panes = "C8"
