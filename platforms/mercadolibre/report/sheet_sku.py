# -*- coding: utf-8 -*-
"""MercadoLibre report — ⑦ SKU盈利分析, with the per-order splits it needs
(packages, commission vs withheld tax, return reviews)."""

import re

import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from core.report_kit import (BOX, C_AMBER, C_BAND, C_CALC, C_GREEN, C_GREY, C_INPUT, C_NAVY,
    C_RED, C_SUB, C_WARN, F, INT, PCT, fill)

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (FEE_BY_ES, RETURN_REASON_ZH, RETURN_RESULT_ZH,
    SHEET_NAMES, SPANISH_MONTHS, STATES_BACK_ON_SALE, STATUS_ZH, V_ORDER_ID, V_SKU, V_STATUS,
    V_TITLE, V_UNITS, fmt, idkey, money, num, parse_spanish_date)


# ════════════════════════════════════════════════════════════════════════
# ⑦ SKU 盈利分析
# ════════════════════════════════════════════════════════════════════════

PKG_RE = re.compile(r"^Paquete de (\d+) productos", re.I)


def split_packages(A):
    """把组合订单母行的金额按 件数×单价 摊到它的子行上。

    MercadoLibre 的导出把一单多件写成：母行 Estado="Paquete de N productos"
    带走**全部金额**却没有 SKU；紧随其后的 N 行有 SKU 和件数，金额列全为空。
    父子之间没有任何共同 ID —— 子行各有自己的 # de venta，母行的
    Orden de compra 还是空的 —— 所以只能靠**相邻位置**关联。

    拆分基数用 件数×单价。实测 Σ(件数×单价) 与母行收入分毫不差
    （BOCINA_TA02 的 538.92、580.00，TANKE_EE 的 397.20、418.60 等均如此），
    所以这是精确拆分而非估算。四舍五入的尾差记到份额最大的那一行，保证拆完
    各子行之和仍等于母行原值。

    **只给 ⑦ SKU 盈利分析用。** 损益表、各座桥和全部校验仍读未拆分的原始明细，
    这样既拿到按 SKU 的归属，又不会扰动已经通过的勾稽（桥A 残差 0.00）。
    """
    A = A.copy().reset_index(drop=True)
    money_cols = [c for c in (site.V_REVENUE, site.V_COMMISSION, site.V_SHIP_INCOME,
                              site.V_SHIP_COST, site.V_REFUND, "coupon",
                              "com_pure", "tax_ret") if c in A.columns]
    if V_STATUS not in A.columns:
        return A, 0
    st = A[V_STATUS].astype(str)
    drop, n_split = [], 0

    for i in range(len(A)):
        m = PKG_RE.match(st.iloc[i].strip())
        if not m:
            continue
        n = int(m.group(1))
        kids = list(range(i + 1, min(i + 1 + n, len(A))))
        # 子行数量对不上就原样保留 —— 宁可留一行空 SKU，也不要把金额摊到
        # 不相干的订单上。
        if len(kids) != n:
            continue

        units = pd.to_numeric(A.loc[kids, V_UNITS], errors="coerce").fillna(0.0)
        if site.V_UNIT_PRICE in A.columns:
            price = pd.to_numeric(A.loc[kids, site.V_UNIT_PRICE], errors="coerce").fillna(0.0)
        else:
            price = pd.Series(0.0, index=kids)
        basis = units * price
        if float(basis.sum()) <= 0:          # 没有单价就退回按件数分
            basis = units.copy()
        if float(basis.sum()) <= 0:          # 件数也没有就均分
            basis = pd.Series(1.0, index=kids)
        share = basis / float(basis.sum())
        big = share.idxmax()

        for c in money_cols:
            amt = pd.to_numeric(pd.Series([A.loc[i, c]]), errors="coerce").fillna(0.0).iloc[0]
            alloc = (share * float(amt)).round(2)
            alloc.loc[big] += round(float(amt) - float(alloc.sum()), 2)
            cur = pd.to_numeric(A.loc[kids, c], errors="coerce").fillna(0.0)
            A.loc[kids, c] = cur + alloc
        drop.append(i)
        n_split += 1

    if drop:
        A = A.drop(index=drop).reset_index(drop=True)

    # ---- 第二遍：换货订单 ----
    # 换货写成两行且**共享同一个 # de venta**：一行带金额没有 SKU
    # （Estado="Venta con solicitud de cambio"），另一行带 SKU 没有金额。
    # 这比包裹可靠 —— 有共同 ID，不用靠相邻位置猜。
    if V_ORDER_ID in A.columns and V_SKU in A.columns:
        sk = A[V_SKU].astype(str).str.strip()
        has_sku = (sk != "") & (sk.str.lower() != "nan")
        rev = pd.to_numeric(A[site.V_REVENUE], errors="coerce").fillna(0.0)             if site.V_REVENUE in A.columns else pd.Series(0.0, index=A.index)
        drop2 = []
        for oid, idx in A.groupby(A[V_ORDER_ID].astype(str)).groups.items():
            idx = list(idx)
            if len(idx) < 2:
                continue
            payers = [i for i in idx if not has_sku[i] and abs(rev[i]) > 0.01]
            takers = [i for i in idx if has_sku[i]]
            if len(payers) != 1 or not takers:
                continue
            p = payers[0]
            units = pd.to_numeric(A.loc[takers, V_UNITS], errors="coerce").fillna(0.0)
            basis = units if float(units.sum()) > 0 else pd.Series(1.0, index=takers)
            share = basis / float(basis.sum())
            big = share.idxmax()
            for c in money_cols:
                amt = pd.to_numeric(pd.Series([A.loc[p, c]]), errors="coerce").fillna(0.0).iloc[0]
                alloc = (share * float(amt)).round(2)
                alloc.loc[big] += round(float(amt) - float(alloc.sum()), 2)
                cur = pd.to_numeric(A.loc[takers, c], errors="coerce").fillna(0.0)
                A.loc[takers, c] = cur + alloc
            drop2.append(p)
            n_split += 1
        if drop2:
            A = A.drop(index=drop2).reset_index(drop=True)

    return A, n_split


def sku_grid():
    """⑦ 页四张表共用的一套列。主表用 A~P，三张明细多出来的项放到 Q 之后。

    四张表共用同一套列，是为了让人可以直接把鼠标从主表拖到任意一张明细上
    求和 —— 同一列永远是同一个项目。明细上没有的主表列留空（不是 0，
    0 会被误读成"这项是零"），主表上没有的明细列在主表里同样留空。
    """
    return [
        # (列名, 宽度)
        ("店铺", 16), ("SKU", 16), ("商品名称", 38),
        ("订单数", 9), ("销售件数", 10),
        ("销售额 GMV", 14), ("运费收入", 12),
        ("平台佣金", 13), (site.TAXCOL_CN, 15), ("平台配送费", 13),
        ("退款金额", 13), ("卖家优惠券", 12),
        ("单位采购成本\n（只在表五填）", 14), ("采购成本合计", 14),
        ("净贡献毛利", 14), ("贡献率", 10),
        # ↓ 以下仅明细表使用，主表留空
        ("订单号", 20), ("销售日期", 13), ("订单状态", 30),
        ("资金归属", 14), ("库存处置", 24),
        ("确定可回收日期", 15), ("质检结论", 16), ("质检原因 / 商品状态", 30),
        ("去向", 12), ("判定来源", 16), ("原始状态 (Estado)", 40),
        # ↓ 仅表五使用
        ("其中：正常销售件数", 14), ("其中：损失件数(不可再售)", 15), ("损失件数占比", 12),
    ]


SKU_NCOL = 30
SKU_MAIN_NCOL = 16
# 净贡献毛利 ＝ 销售额 ＋ 运费收入 − 佣金 − 代扣税 − 配送费 − 退款 − 优惠券 − 采购成本
SKU_NET = "=F{r}+G{r}-H{r}-I{r}-J{r}-K{r}-L{r}-N{r}"
SKU_RATE = "=IF(F{r}+G{r}=0,0,O{r}/(F{r}+G{r}))"

# 取消类订单：商品从未发出，库存没有减少。
CANCEL_RE = re.compile(r"cancel", re.I)
# Returns 报表的质检判定"重新上架"——库存实际回到手上
RESULT_BACK_ON_SALE = "Producto quedó nuevamente para la venta"
# Ventas 报表自带的质检三列（Resultado / Destino / Motivo del resultado）。
# 它和 Returns 报表是两个不同的来源：这三列就在订单行上，不用关联，覆盖面更广；
# Returns 报表则是 Full 仓的质检台账，两者都用，前者优先。
V_REVIEW_DATE = "Fecha de revisión"
V_REVIEW_RESULT = "Resultado"
V_REVIEW_DEST = "Destino"
V_REVIEW_WHY = "Motivo del resultado"
VENTAS_RESULT_ZH = {
    "Apto para la venta": ("可再售", True),
    "No apto para la venta": ("不可再售", False),
    "Regresado al comprador": ("已退回买家", False),
}
VENTAS_DEST_ZH = {
    "Vendedor": "退回卖家",
    "Comprador": "留在买家处",
    "Mercado Libre": "留在平台仓",
}
VENTAS_WHY_ZH = {
    "está en buenas condiciones": "状况良好",
    "están en buenas condiciones": "多件均状况良好",
    "no está en buenas condiciones": "状况不佳",
    "no funciona": "无法正常工作",
    "presenta irregularidades": "存在异常",
    "tiene el sello de la caja dañado": "包装封条破损",
    "tiene la caja original abierta": "原包装已拆封",
    "tiene la caja original dañada": "原包装破损",
    "tiene marcas de uso": "有使用痕迹",
}


def parse_spanish_daymonth(value, after=None):
    """'8 de septiembre' —— 平台的质检日**不带年份**，只能靠销售日推年。

    质检一定发生在销售之后，所以取销售当年；若这样算出来的日期早于销售日
    （典型是 12 月下单、次年 1 月质检），就进位到下一年。没有销售日可参照时
    返回 NaT，宁可留空也不要编一个年份出来。
    """
    if isinstance(value, pd.Timestamp):
        return value
    if not isinstance(value, str):
        return pd.NaT
    m = re.match(r"\s*(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)\s*$", value.strip(), re.IGNORECASE)
    if not m:
        return parse_spanish_date(value)      # 带年份的照旧走原解析
    month = SPANISH_MONTHS.get(m.group(2).lower())
    if not month or after is None or pd.isna(after):
        return pd.NaT
    for y in (after.year, after.year + 1):
        try:
            d = pd.Timestamp(y, month, int(m.group(1)))
        except ValueError:
            return pd.NaT
        if d >= after.normalize():
            return d
    return pd.NaT


def split_commission(S):
    """把 Ventas 的『佣金及税金』一列拆成 纯佣金 与 代扣代缴税金 两列。

    平台在 Ventas 报表里把这两项**合并**给出（Cargo por venta e impuestos），
    在账单明细（Facturación）里只给纯佣金（Cargo por venta），代扣税从来没有
    单独的科目。所以逐单的拆法只能是：

        某单代扣税 ＝ 该单 Ventas 佣金及税 − 该单账单纯佣金

    账单行通过 # de venta / Orden de compra 关联到订单，可以跨账期
    （8 月的单，佣金可能记在 9 月账单上）。关联不上的少数订单用
    店内已知部分推出的比例补足，并保证两列之和仍等于损益表的
    平台佣金 + 平台代扣代缴税金 —— 页面因此与损益表分毫不差。

    返回 (com_pure, tax_ret, 说明文字)，两个 Series 与 S["ventas"] 同索引。
    """
    A = S["ventas"]
    v_com = -num(A, site.V_COMMISSION)                    # Ventas 口径，正数=费用
    if not len(A):
        return v_com, v_com, "本期没有订单"

    ids = set(A["oid"]) | set(A["pid"])
    ids.discard(""); ids.discard("nan")
    L = S["billing_linked"]
    com = L[(~L["is_rev"]) & L["parent_detalle"].map(
        lambda d: FEE_BY_ES.get(str(d), (None, None, None))[1] == "commission").astype(bool)]

    # 账单行 → 订单键。同一行的 k_sale / k_pack 只认能对上 Ventas 的那个，
    # 两个都对不上就放弃这行（它的钱由后面的残差分摊兜住）。
    by_key = {}
    for ks, kp, amt in zip(com["k_sale"].astype(str).str.strip(),
                           com["k_pack"].astype(str).str.strip(),
                           com["amount"]):
        k = ks if ks in ids else (kp if kp in ids else None)
        if k:
            by_key[k] = by_key.get(k, 0.0) + float(amt)

    pure = pd.Series(0.0, index=A.index)
    hit = pd.Series(False, index=A.index)
    for i, oid in A["oid"].items():
        if oid in by_key:
            pure[i] = by_key[oid]
            hit[i] = True
    # 只挂在包裹号上的账单行：按包裹内各行的 Ventas 佣金比例分摊
    rest = A.index[~hit]
    for pid, idx in A.loc[rest].groupby(A.loc[rest, "pid"]).groups.items():
        if pid not in by_key:
            continue
        idx = list(idx)
        w = v_com[idx].abs()
        w = w / w.sum() if float(w.sum()) > 0 else pd.Series(1.0 / len(idx), index=idx)
        pure[idx] = w * by_key[pid]
        hit[idx] = True

    # 残差：账单里有、但一行都没挂上的佣金。按 Ventas 佣金比例摊到没对上的行，
    # 全都对上了就摊回全部行。不摊会让代扣税凭空变大，比估算更糟。
    resid = float(S["m"]["com_gross"]) - float(pure.sum())
    if abs(resid) > 0.005:
        tgt = A.index[~hit] if (~hit).any() else A.index
        w = v_com[tgt].abs()
        w = w / w.sum() if float(w.sum()) > 0 else pd.Series(1.0 / len(tgt), index=tgt)
        pure[tgt] = pure[tgt] + w * resid

    n_hit = int(hit.sum())
    note = ("%d/%d 单的佣金直接取自账单明细" % (n_hit, len(A))
            + ("；其余 %d 单账单尚未出，按店内实际佣金率推算" % (len(A) - n_hit)
               if n_hit < len(A) else "")
            + ("；另有 %s %s 账单佣金未能挂到具体订单，已按比例摊回" % (fmt(resid), site.CUR)
               if abs(resid) > 0.005 and n_hit == len(A) else ""))
    return pure, v_com - pure, note


def returns_review(S):
    """订单号 → (质检日, 处置结果, 商品状态)。取全部期间的质检记录，
    因为本期订单的退货常常下个月才质检。

    这是 Ventas 自带质检列的**补充来源**：Returns 报表只覆盖 Full 仓，
    且它的订单号与 Ventas 不一定有交集，所以只在 Ventas 那三列为空时才用。
    """
    R = S.get("returns_all")
    out = {}
    if R is None or not len(R) or "Número de orden" not in R.columns:
        return out
    blank = pd.Series([None] * len(R), index=R.index)
    for oid, fr, res, est in zip(idkey(R["Número de orden"]), R["fr"],
                                 R["Resultado de la revisión"] if "Resultado de la revisión" in R.columns else blank,
                                 R["Estado del producto"] if "Estado del producto" in R.columns else blank):
        out[oid] = (fr, str(res) if pd.notna(res) else "",
                    str(est) if pd.notna(est) else "")
    return out


def _txt(x, col):
    v = x.get(col)
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    v = str(v).strip()
    return "" if v in ("nan", "-") else v


def review_of(x, fallback):
    """一行异常订单的质检结论，以及它是从哪来的。

    先看 Ventas 自带的 Resultado/Destino/Motivo（就在订单行上，最可靠）；
    没有才回退到 Returns 报表按订单号关联的记录；两边都没有就是"未质检"。

    返回 (日期, 结论中文, 原因中文, 去向中文, 来源, 可回收?)，
    可回收? 为 None 表示平台还没给结论 —— 与"判定为不可回收"必须区分开。
    """
    res = _txt(x, V_REVIEW_RESULT)
    if res:
        zh, ok = VENTAS_RESULT_ZH.get(res, (res, None))
        d = parse_spanish_daymonth(_txt(x, V_REVIEW_DATE), x.get("fecha"))
        why = _txt(x, V_REVIEW_WHY)
        return (d, zh, VENTAS_WHY_ZH.get(why, why),
                VENTAS_DEST_ZH.get(_txt(x, V_REVIEW_DEST), _txt(x, V_REVIEW_DEST)),
                "销售报表质检列", ok)
    if fallback:
        fr, r2, est = fallback
        return (fr, RETURN_RESULT_ZH.get(r2, (r2, "", ""))[0],
                RETURN_REASON_ZH.get(est, (est, ""))[0], "",
                "Returns 质检报表", (r2 == RESULT_BACK_ON_SALE) if r2 else None)
    return (pd.NaT, "", "", "", "", None)


def classify_abnormal(estado, rv):
    """异常订单归到哪一张明细表。

    分三类是因为三者的**库存后果完全不同**，混在一起算 SKU 盈利必然是错的：
      cancel  —— 从未发货，库存没动，不该有任何采购成本，毛利应为零
      back    —— 退货已质检判定可再售，货回到手上，同样不消耗采购成本
      other   —— 判定不可再售、或平台还没给结论，以及仲裁/投诉/换货/未送达：
                 库存价值已损失或去向未定，要计采购成本

    注意"还没给结论"归 other 而不是 back —— 在平台说可以之前，
    默认这批货收不回来才是保守且安全的。
    """
    if CANCEL_RE.search(estado or ""):
        return "cancel"
    if rv and rv[5] is True:
        return "back"
    if rv and rv[5] is None and estado in STATES_BACK_ON_SALE:
        return "back"      # 老状态文案已经写明"重新上架"，即便没有质检行
    return "other"


def sheet_sku(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[6])
    ws["A1"] = "SKU 盈利分析（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("【使用说明】单位采购成本（到岸成本，含头程运费与关税）只在最下方"
                "【表五】的黄色 M 列填一次，表一与表四会自动引用同一个值，不要分头填。单位：%s" % site.CUR)
    ws["A2"].font = F(9, True, C_RED)
    ws["A3"] = ("本页共五张表。**要看某个 SKU 到底赚不赚钱，直接看最下方的【表五】** —— "
                "它是最终口径。表一只算正常订单，表二至表四是三类异常订单的逐单明细。"
                "五张表**列完全对齐**，同一列永远是同一个项目，可以直接拖着跨表求和；"
                "明细独有的项放在 Q 列之后，不适用的表留空。")
    ws["A3"].font = F(9, True, C_NAVY)
    n_par = sum(len(S["cn"]["partial"]) for S in stores)
    if n_par:
        ws["A5"] = ("⚠ 本页的“退款金额”取自销售报表原始数据。其中 %d 笔**部分退款**订单，"
                    "平台在那一列里把原佣金重复扣了一次，所以对应 SKU 的净贡献被低估了"
                    "相应金额（合计 %s）。本页保持与销售报表一一对应不作修正；"
                    "利润口径的修正在【%s】，逐单明细见【%s】3.3。"
                    % (n_par, fmt(sum(S["cn"]["partial_total"] for S in stores)),
                       SHEET_NAMES[1], SHEET_NAMES[9]))
        ws["A5"].font = F(9, False, C_AMBER)
    ws["A4"] = (("注 1：佣金与代扣代缴税金已拆成两列 —— 平台在 Ventas 报表里把它们合并给出，"
                 if site.WITHHOLDING else
                 "注 1：佣金与并进佣金列的退货处理费等已拆成两列（智利平台不代扣税）—— "
                 "平台在 Ventas 报表里把它们合并给出，")
                + "本页用账单明细的纯佣金逐单还原（见表末说明）。每一列只放一项费用，不再合并。"
                "注 2：组合订单（Paquete de N productos）的母行金额已按 件数×单价 拆到各子 SKU。"
                "注 3：广告费、Full 仓储/揽收费在平台账单里没有商品编号，无法按 SKU 拆分，"
                "故本页是“订单级贡献”，未分摊店铺级共同费用。")
    ws["A4"].font = F(9, False, C_GREY)

    grid = sku_grid()
    HDR = 6

    def header(r, main):
        for i, (t, w) in enumerate(grid, 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_NAVY if (main and i <= SKU_MAIN_NCOL) else C_BAND)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
            if main:
                ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[r].height = 34

    header(HDR, True)

    def band(r, color, bold=False):
        for c in range(1, SKU_NCOL + 1):
            ws.cell(row=r, column=c).fill = fill(color)
            ws.cell(row=r, column=c).border = BOX
            if bold:
                ws.cell(row=r, column=c).font = F(10, True)

    def total_row(r, start, label, extra=()):
        """合计行：五张表列位置一一对应，所以整段直接求和即可。"""
        ws.cell(row=r, column=2, value=label).font = F(10, True)
        for col in tuple((4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15)) + tuple(extra):
            cl = get_column_letter(col)
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)
                    ).number_format = INT if col in (4, 5) or col in extra else site.MNY
        ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
        band(r, C_SUB, bold=True)

    def money_cells(r, vals):
        for col, v in vals:
            ws.cell(row=r, column=col, value=money(v)).number_format = site.MNY

    def sku_agg(D):
        """一组订单行 → 该 SKU 的各项金额。五张表都用它，口径不会分叉。"""
        return {
            "订单数": len(D), "件数": float(num(D, V_UNITS).sum()),
            "GMV": float(num(D, site.V_REVENUE).sum()),
            "运费收入": float(num(D, site.V_SHIP_INCOME).sum()),
            "佣金": float(num(D, "com_pure").sum()),
            "代扣税": float(num(D, "tax_ret").sum()),
            "配送费": float(-num(D, site.V_SHIP_COST).sum()),
            "退款": float(-num(D, site.V_REFUND).sum()),
            "券": float(D["coupon"].sum()) if "coupon" in D.columns else 0.0,
        }

    buckets = {"cancel": [], "back": [], "other": []}   # 跨店收集，明细按类分表
    split_notes = []
    other_cost_cell = {}      # 店铺 → 表一"其他异常小计"行号，稍后回指表四
    frames = {}               # 店铺 → (正常订单, 全部异常订单)
    cost_ref_rows = []        # 表一里需要回指表五取成本的行

    r = HDR + 1
    for S in stores:
        pure, tax, note = split_commission(S)
        split_notes.append("%s：%s" % (S["store"], note))
        V = S["ventas"].copy()
        V["com_pure"] = pure
        V["tax_ret"] = tax
        A, n_split = split_packages(V)          # 组合单的金额（含新加的两列）摊到子行
        S["_pkg_split"] = n_split
        fb = returns_review(S)

        ab = A[A["abnormal"]] if "abnormal" in A.columns else A.iloc[0:0]
        ok = A[~A["abnormal"]] if "abnormal" in A.columns else A
        kinds = {}
        for i, x in ab.iterrows():
            rv = review_of(x, fb.get(str(x.get("oid", ""))))
            k = classify_abnormal(str(x.get(V_STATUS, "")), rv)
            buckets[k].append((S["store"], x, rv))
            kinds[i] = k
        ab = ab.copy()
        ab["_kind"] = pd.Series(kinds) if kinds else ""
        frames[S["store"]] = (ok, ab)

        g = ok.groupby(ok[V_SKU].astype(str)).apply(lambda x: pd.Series(dict(
            sku_agg(x), 名称=(str(x[V_TITLE].iloc[0])[:45] if V_TITLE in x.columns else "")
        ))).reset_index()
        if len(g):
            g.columns = ["SKU"] + list(g.columns[1:])
            g = g.sort_values(["GMV", "SKU"], ascending=[False, True])
        start = r
        for _, x in g.iterrows():
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value=str(x["SKU"]))
            ws.cell(row=r, column=3, value=x["名称"])
            ws.cell(row=r, column=4, value=int(x["订单数"])).number_format = INT
            ws.cell(row=r, column=5, value=int(x["件数"])).number_format = INT
            money_cells(r, [(6, x["GMV"]), (7, x["运费收入"]), (8, x["佣金"]),
                            (9, x["代扣税"]), (10, x["配送费"]), (11, x["退款"]),
                            (12, x["券"])])
            cost_ref_rows.append(r)          # M 列稍后回指表五
            ws.cell(row=r, column=14, value="=E%d*M%d" % (r, r)).number_format = site.MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = site.MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            for c in range(1, SKU_NCOL + 1):
                ws.cell(row=r, column=c).border = BOX
            for c in list(range(1, 13)) + list(range(14, 17)):
                ws.cell(row=r, column=c).font = F(10)
            r += 1

        # 三类异常各留一行小计。留在表一里，本店合计才与损益表的 GMV 对得上。
        # 取消与可回收的采购成本恒为 0；其他异常的成本在表四逐单算，这里回指过去，
        # 不然表四有成本、本店合计却看不见。
        for kind, label, why in (
            ("cancel", "取消订单小计", "从未发货，库存未减少，不计采购成本；明细见表二"),
            ("back", "退货可回收小计", "已质检判定可再售，库存已收回，不计采购成本；明细见表三"),
            ("other", "其他异常订单小计", "退货不可再售、仲裁、投诉、换货、尚未质检等；采购成本取自表四"),
        ):
            sub = [x for st, x, _ in buckets[kind] if st == S["store"]]
            if not sub:
                continue
            a = sku_agg(pd.DataFrame(sub))
            ws.cell(row=r, column=2, value=label).font = F(10, True, C_AMBER)
            ws.cell(row=r, column=3, value=why).font = F(9, False, C_GREY)
            ws.cell(row=r, column=4, value=a["订单数"]).number_format = INT
            ws.cell(row=r, column=5, value=int(a["件数"])).number_format = INT
            money_cells(r, [(6, a["GMV"]), (7, a["运费收入"]), (8, a["佣金"]),
                            (9, a["代扣税"]), (10, a["配送费"]), (11, a["退款"]),
                            (12, a["券"])])
            ws.cell(row=r, column=14, value=0).number_format = site.MNY
            if kind == "other":
                other_cost_cell[S["store"]] = r
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = site.MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            band(r, C_WARN)
            r += 1

        total_row(r, start, "%s 合计" % S["store"])
        r += 2

    # ──────────────────────────────────────────────────────────────
    # 表二~表四：三张逐单明细。列与表一完全一致，Q 列之后才是明细独有的项。
    # ──────────────────────────────────────────────────────────────
    nz = lambda v: 0.0 if pd.isna(v) else float(v)
    detail_cost_rows = []

    def detail(title, intro, kind, cost_ref):
        """返回明细数据行的 (首行, 末行)，没有数据时返回 None。"""
        nonlocal r
        rows = buckets[kind]
        r += 1
        ws.cell(row=r, column=1, value=title).font = F(12, True, C_NAVY)
        r += 1
        ws.cell(row=r, column=1, value=intro).font = F(9, False, C_GREY)
        r += 1
        header(r, False)
        r += 1
        if not rows:
            ws.cell(row=r, column=1, value="本期没有这一类订单。").font = F(10, False, C_GREEN)
            r += 2
            return None
        start = r
        for store, x, rv in rows:
            estado = str(x.get(V_STATUS, ""))
            zh, who, inv = STATUS_ZH.get(estado, (estado, "待确认", "待确认"))
            rd, rres, rwhy, rdest, rsrc, _ = rv
            ws.cell(row=r, column=1, value=store)
            ws.cell(row=r, column=2, value=str(x.get(V_SKU, "")))
            ws.cell(row=r, column=3, value=str(x.get(V_TITLE, ""))[:45])
            ws.cell(row=r, column=4, value=1).number_format = INT
            ws.cell(row=r, column=5, value=int(nz(x.get(V_UNITS)))).number_format = INT
            money_cells(r, [(6, nz(x.get(site.V_REVENUE))), (7, nz(x.get(site.V_SHIP_INCOME))),
                            (8, nz(x.get("com_pure"))), (9, nz(x.get("tax_ret"))),
                            (10, -nz(x.get(site.V_SHIP_COST))), (11, -nz(x.get(site.V_REFUND))),
                            (12, nz(x.get("coupon")))])
            if cost_ref:
                detail_cost_rows.append(r)
                ws.cell(row=r, column=14, value="=E%d*M%d" % (r, r)).number_format = site.MNY
            else:
                ws.cell(row=r, column=14, value=0).number_format = site.MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = site.MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            # ---- Q 之后：明细独有的项 ----
            ws.cell(row=r, column=17, value=str(x.get("oid", "")))
            d = x.get("fecha")
            ws.cell(row=r, column=18, value=("" if pd.isna(d) else d.strftime("%Y-%m-%d")))
            ws.cell(row=r, column=19, value=zh)
            ws.cell(row=r, column=20, value=who)
            ws.cell(row=r, column=21, value=inv)
            if pd.notna(rd):
                ws.cell(row=r, column=22, value=rd.strftime("%Y-%m-%d"))
            elif kind == "back":
                # 已经归到可回收，说明判定确实做了，只是平台没给日期 ——
                # 不能写成"尚未质检"，那是另一回事。
                ws.cell(row=r, column=22, value="平台未给日期").font = F(9, False, C_AMBER)
            elif rres:
                ws.cell(row=r, column=22, value="平台未给日期").font = F(9, False, C_AMBER)
            elif kind == "other":
                ws.cell(row=r, column=22, value="尚未质检").font = F(9, False, C_AMBER)
            ws.cell(row=r, column=23, value=rres)
            ws.cell(row=r, column=24, value=rwhy)
            ws.cell(row=r, column=25, value=rdest)
            ws.cell(row=r, column=26, value=rsrc)
            ws.cell(row=r, column=27, value=estado)
            for c in range(1, SKU_NCOL + 1):
                ws.cell(row=r, column=c).border = BOX
                if c != 13 and ws.cell(row=r, column=c).font.color is None:
                    ws.cell(row=r, column=c).font = F(9)
            r += 1
        total_row(r, start, "合计 %d 单" % len(rows))
        r += 2
        return start, r - 3

    detail("表二 · 取消订单明细（商品从未发出，库存未减少）",
           "这些单在发货前就取消了：没有出库、库存没有减少，所以**不应产生任何采购成本**，"
           "采购成本合计一律记 0，净贡献只反映平台可能仍收取或已退还的费用。"
           "货还在手上可以再卖，所以表五**不把它们算进各 SKU 的盈利**，"
           "只在表五末尾单列一行残留费用。",
           "cancel", cost_ref=False)

    detail("表三 · 退货可回收明细（已质检判定可再售）",
           "退货已经过质检、判定可以重新销售，库存实际收回，因此同样不消耗采购成本，"
           "表五也不把它们算进各 SKU 的盈利。"
           "“确定可回收日期”＝ 平台做出这个判定的那一天（销售报表的 Fecha de revisión，"
           "缺失时取 Returns 质检报表的质检日）；“判定来源”列标明这一行取自哪个报表。",
           "back", cost_ref=False)

    span = detail("表四 · 其他异常订单明细（不可再售、仲裁、投诉、换货、未送达、尚未质检）",
                  "这些单的商品要么已判定不可再售、要么去向未定、要么平台还没质检，"
                  "库存价值不能假定收得回来 —— 所以**表五把它们并进对应 SKU**，"
                  "按件计采购成本。单位成本自动取自表五，不用在这里重填。",
                  "other", cost_ref=True)
    if span:
        d0, d1 = span
        for store, row in other_cost_cell.items():
            ws.cell(row=row, column=14,
                    value='=SUMIF($A$%d:$A$%d,"%s",$N$%d:$N$%d)' % (d0, d1, store, d0, d1)
                    ).number_format = site.MNY

    # ──────────────────────────────────────────────────────────────
    # 表五 · SKU 真实盈利。正常订单 ＋ 不可再售的异常订单，按 SKU 合并。
    # ──────────────────────────────────────────────────────────────
    r += 1
    ws.cell(row=r, column=1, value="表五 · SKU 真实盈利（最终口径）").font = F(13, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "口径：正常订单 ＋ 表四的“其他异常”订单，按 SKU 合并。"
        "取消（表二）与退货可再售（表三）**不计入各 SKU** —— 这两类的货还在手上、还能再卖，"
        "把它们摊到 SKU 上会凭空压低该 SKU 的盈利；而表四那批货已经确定收不回来，"
        "货值实实在在损失了，必须算进这个 SKU 的成本里。"
    )).font = F(9, False, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "但取消与可再售在**钱**上并不恒等于零（平台常常退了货款却不退配送费），"
        "所以每店末尾单列两行残留费用，把这部分留在合计里。"
        "因此本表的店铺合计与表一的店铺合计相等，也与损益表对得上 —— 只是钱的归属换了个分法。"
        "（两表可能相差 1 分：表一按“正常/异常”分组后各自四舍五入到分，本表按 SKU 合并后再舍入，"
        "落点不同。差额只在佣金与%s之间互相抵消，两列之和不变。）" % site.TAX_SHORT
    )).font = F(9, False, C_GREY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "★ 黄色 M 列是全页**唯一**的成本录入位：填在这里，表一与表四会自动引用。"
    )).font = F(9, True, C_RED)
    r += 1
    header(r, True)
    r += 1
    t5_start = r

    for S in stores:
        ok, ab = frames[S["store"]]
        other = ab[ab["_kind"] == "other"] if len(ab) else ab
        keys = sorted(set(ok[V_SKU].astype(str)) | set(other[V_SKU].astype(str)))
        rows = []
        for sku in keys:
            o = ok[ok[V_SKU].astype(str) == sku]
            t = other[other[V_SKU].astype(str) == sku] if len(other) else other
            a = sku_agg(pd.concat([o, t]) if len(t) else o)
            title = ""
            for D in (o, t):
                if len(D) and V_TITLE in D.columns:
                    title = str(D[V_TITLE].iloc[0])[:45]
                    break
            a.update({"SKU": sku, "名称": title,
                      "正常件数": float(num(o, V_UNITS).sum()),
                      "损失件数": float(num(t, V_UNITS).sum()) if len(t) else 0.0})
            rows.append(a)
        rows.sort(key=lambda a: (-a["GMV"], a["SKU"]))

        start = r
        for a in rows:
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value=a["SKU"])
            ws.cell(row=r, column=3, value=a["名称"])
            ws.cell(row=r, column=4, value=a["订单数"]).number_format = INT
            ws.cell(row=r, column=5, value=int(a["件数"])).number_format = INT
            money_cells(r, [(6, a["GMV"]), (7, a["运费收入"]), (8, a["佣金"]),
                            (9, a["代扣税"]), (10, a["配送费"]), (11, a["退款"]),
                            (12, a["券"])])
            cc = ws.cell(row=r, column=13, value=0)
            cc.number_format = site.MNY
            cc.fill = fill(C_INPUT)
            cc.font = F(10, True, "0000FF")
            ws.cell(row=r, column=14, value="=E%d*M%d" % (r, r)).number_format = site.MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = site.MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            ws.cell(row=r, column=28, value=int(a["正常件数"])).number_format = INT
            ws.cell(row=r, column=29, value=int(a["损失件数"])).number_format = INT
            ws.cell(row=r, column=30,
                    value="=IF(E%d=0,0,AC%d/E%d)" % (r, r, r)).number_format = PCT
            for c in range(1, SKU_NCOL + 1):
                ws.cell(row=r, column=c).border = BOX
            for c in list(range(1, 13)) + list(range(14, 17)) + [28, 29, 30]:
                ws.cell(row=r, column=c).font = F(10)
            if a["损失件数"] > 0:
                for c in (2, 29, 30):
                    ws.cell(row=r, column=c).fill = fill(C_WARN)
            r += 1

        for kind, label, why in (
            ("cancel", "取消订单残留费用", "货未发出，不摊到任何 SKU；此处只留平台未退还的那部分费用"),
            ("back", "退货可再售残留费用", "货已收回可再卖，不摊到任何 SKU；此处只留平台未退还的那部分费用"),
        ):
            sub = [x for st, x, _ in buckets[kind] if st == S["store"]]
            if not sub:
                continue
            a = sku_agg(pd.DataFrame(sub))
            ws.cell(row=r, column=2, value=label).font = F(10, True, C_AMBER)
            ws.cell(row=r, column=3, value=why).font = F(9, False, C_GREY)
            ws.cell(row=r, column=4, value=a["订单数"]).number_format = INT
            ws.cell(row=r, column=5, value=int(a["件数"])).number_format = INT
            money_cells(r, [(6, a["GMV"]), (7, a["运费收入"]), (8, a["佣金"]),
                            (9, a["代扣税"]), (10, a["配送费"]), (11, a["退款"]),
                            (12, a["券"])])
            ws.cell(row=r, column=14, value=0).number_format = site.MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = site.MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            band(r, C_WARN)
            r += 1

        total_row(r, start, "%s 合计（＝表一同名行）" % S["store"], extra=(28, 29))
        r += 2
    t5_end = r - 2

    # 表一与表四的单位采购成本回指表五。用 SUMIFS 按【店铺＋SKU】两个键匹配 ——
    # 汇总报表里同一个 SKU 可能出现在好几家店，只按 SKU 找会取到别家的成本。
    cost_ref = ('=SUMIFS($M$%d:$M$%d,$A$%d:$A$%d,$A{r},$B$%d:$B$%d,$B{r})'
                % (t5_start, t5_end, t5_start, t5_end, t5_start, t5_end))
    for row in cost_ref_rows + detail_cost_rows:
        c = ws.cell(row=row, column=13, value=cost_ref.format(r=row))
        c.number_format = site.MNY
        c.fill = fill(C_CALC)
        c.font = F(10, False, C_GREY)
    # ⑨ 库存页也从这里取成本，同样按店铺＋SKU 两个键
    ctx["sku_cost_ref"] = (SHEET_NAMES[6], t5_start, t5_end)

    # ---- 表末：佣金/代扣税拆分的取数说明，避免有人以为这两列是平台直接给的 ----
    r += 1
    ws.cell(row=r, column=1, value="佣金与%s的拆分口径" % ("代扣代缴税金" if site.WITHHOLDING else site.TAXCOL_CN)).font = F(11, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "平台只在 Ventas 报表给出合并值 Cargo por venta e impuestos（佣金＋代扣税），"
        "账单明细 Facturación 只给纯佣金 Cargo por venta，代扣税没有独立科目。"
        "本页按订单号把账单的纯佣金挂回每一单，代扣税 ＝ 合并值 − 纯佣金。"
        "两列之和等于损益表的『平台销售佣金 ＋ 平台代扣代缴税金』。" if site.WITHHOLDING else
        "平台只在 Ventas 报表给出合并值 Cargo por venta e impuestos。智利平台不代扣税：正常订单上"
        "这个值就是账单的纯佣金 Cargo por venta；退货订单上还并进了退货处理费。本页按订单号把账单的"
        "纯佣金挂回每一单，I 列 ＝ 合并值 − 纯佣金（主要是退货处理费）。两列之和等于损益表的"
        "『平台销售佣金 ＋ 退货处理费』再加上桥 E 的时间/关联差。")).font = F(9, False, C_GREY)
    r += 1
    for t in split_notes:
        ws.cell(row=r, column=1, value=t).font = F(9, False, C_GREY)
        r += 1

    ws.freeze_panes = "A7"
