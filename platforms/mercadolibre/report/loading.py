# -*- coding: utf-8 -*-
"""MercadoLibre report — loading layer: find and read one store's downloaded files.

Ventas, billing (Facturación / Notas de crédito), Pagos, Cargos Full, storage,
returns, stock, MercadoPago settlement / cash, run_meta.json; the bill-cycle
rules (Chile's bills run from the store's cut-off day to the next month's) and
the credit-note × Ventas cross-check. Nothing here writes a sheet."""

import datetime
import glob
import io
import json
import os
import re

import pandas as pd

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import (ABNORMAL_RE, BILLING_HEADER_ROW,
    CARGOS_HEADER_ROW, CH_NO, FEE_BY_ES, F_AMOUNT, F_CHANNEL, F_CHARGE_DATE, F_CHARGE_NO,
    F_DETALLE, F_PACK_NO, F_REVERSES, F_REVERSES_ALT, F_SALE_DATE, F_SALE_NO, MONTH_FULL_ES,
    PAGOS_HEADER_ROW, RETURNS_EN_COLUMNS, RETURNS_EN_VALUES, RETURNS_HEADER_ROW,
    REVERSAL_TO_PARENT, SPANISH_MONTHS, STORAGE_HEADER_ROW, VENTAS_HEADER_ROW, V_DATE, V_DISCOUNT,
    V_ORDER_ID, V_PACK_FLAG, V_PACK_ID, V_SKU, V_STATUS, file_month_token, fmt, idkey,
    is_reversal, num, parse_spanish_date, storage_file_range)


BILL_OPEN_RE = re.compile(r"podr[ií]a tener cambios hasta el cierre", re.I)


def open_bills(folder, cycles):
    """会计月用到的账单里，哪几张还没关账（文件顶部写着"Este reporte podría tener
    cambios hasta el cierre del período de facturación"）。返回账单名列表。"""
    out = []
    for y, mth, _a, _b in cycles:
        p = _cycle_file("Reporte_Facturacion_MercadoLibre_*.xlsx", folder, y, mth)
        if not p:
            continue
        try:
            top = pd.read_excel(p, sheet_name="REPORT", header=None, nrows=BILLING_HEADER_ROW)
        except Exception:
            continue
        if any(isinstance(v, str) and BILL_OPEN_RE.search(v) for v in top.values.ravel()):
            out.append("%s %d" % (MONTH_FULL_ES[mth], y))
    return out


def cycle_day(S):
    """这家店的账单切期日（1 = 自然月）。"""
    return S.get("cycle_day", site.BILL_CYCLE_DAY)


def cycle_cn(day):
    """'8 号到次月 7 号'。"""
    return "%d 号到次月 %d 号" % (day, day - 1)


def infer_cycle_day(B, folder):
    """从这家店自己的文件里推账单切期日（智利用；墨西哥恒为自然月，不调用）。

    投票：每份 Facturación 文件（名为 M 的账单）里最早的计费日若落在 M-1 月，
    它的"日"就是切期日（广告和仓储费几乎每天都有，第一天基本不会空着）；
    每份仓储费报表文件名的起始日（24-08-26_23-09-26_… → 24）也投一票。
    取票数最多的；一票都没有时用站点默认值。
    """
    votes = []
    fac = B[B["_doc"] == "FAC"] if "_doc" in B.columns else B
    for f, g in fac.groupby("_file"):
        tok = file_month_token(str(f))
        d0 = g["cdate"].min()
        if tok and pd.notna(d0):
            prev = pd.Timestamp(tok[0], tok[1], 1) - pd.DateOffset(months=1)
            if (d0.year, d0.month) == (prev.year, prev.month) and d0.day > 1:
                votes.append(d0.day)
    for p in glob.glob(os.path.join(folder, "*Costos_por_servicio_almacenamiento.xlsx")):
        rng = storage_file_range(os.path.basename(p))
        if rng:
            votes.append(rng[0].day)
    if not votes:
        return site.BILL_CYCLE_DAY
    return max(sorted(set(votes)), key=votes.count)


def bill_cycles(year, month, day=None):
    """会计月（自然月）跨到的账单周期：[(账单年, 账单月, 周期起, 周期止), ...]。

    会计口径两国一样：收入和订单级费用按下单日、账单级费用按计费日
    （Fecha del cargo）归入**自然月**。不同的只是平台的账单周期：

    * 墨西哥：账单 = 自然月，9 月的会计月只碰到"Septiembre"一张账单。
    * 智利：账单从 8 号到次月 7 号。"Septiembre"账单 = 8/8–9/7，
      "Octubre"账单 = 9/8–10/7。所以 9 月这个会计月要用**两张**账单：
      Septiembre 里计费日 9/1–9/7 的行 ＋ Octubre 里计费日 9/8–9/30 的行。

    derive() 本来就按计费日把所有账单文件合并后再筛（不按文件名挑账单），
    所以费用行本身两国同一套代码就对；这里给出的周期用来：判断账单是否下全
    （两张都要在）、把桥 B 的实际付款按计费日拆到会计月、以及 Full 费用报表
    和仓储费报表的核对范围。
    """
    p0 = pd.Timestamp(year, month, 1)
    p1 = (p0 + pd.offsets.MonthEnd(1)).normalize()
    day = site.BILL_CYCLE_DAY if day is None else day
    out = []
    if day == 1:
        return [(year, month, p0, p1)]
    # 名为 M 的账单覆盖 [M-1 月 day 日, M 月 day-1 日]
    for k in (0, 1):
        b = p0 + pd.DateOffset(months=k)
        start = (b - pd.DateOffset(months=1)).replace(day=day)
        end = b.replace(day=day - 1)
        if end >= p0 and start <= p1:
            out.append((b.year, b.month, start, end))
    return out


# ════════════════════════════════════════════════════════════════════════
# 载入层
# ════════════════════════════════════════════════════════════════════════

def _first(patterns, folder):
    for p in patterns:
        hits = sorted(g for g in glob.glob(os.path.join(folder, p))
                      if not os.path.basename(g).startswith("~$"))
        if hits:
            return hits[0]
    return None


def unreadable(errors, path, err):
    """一份**可选**文件读不出来（平台把出错网页存成了 .xlsx、压缩包坏了、
    页名不对……）：当它不存在，并把 (文件名, 错误) 记进 errors —— derive() 传
    的是 S["load_errors"]，validate() 据此在 ⑩ 页记一条 load:<文件名> 校验，
    和账单基准一样。errors 为 None 时只是安静跳过（老调用方）。同一个文件只记
    一次（结算流水被三个函数各读一遍）。

    2026-10-30 EWTTO_SM：五份库存类文件都是网页，订单基准在这里炸掉，整个
    墨西哥的订单基准汇总都没写出来；账单基准同一天照常出表、评级有缺口。
    只有 Ventas 和 Facturación **不**走这里 —— 没有它们报表就不存在，照旧报错。
    """
    if errors is None:
        return
    name = os.path.basename(path)
    if not any(n == name for n, _ in errors):
        errors.append((name, str(err)))


def cell_int(v):
    """写进单元格：浮点取整，NaN/None 留空。

    直接 int(v) 会在 NaN 上抛 ValueError。这不是假想 —— pandas 3 的
    astype(str) 不再把 NaN 变成 "nan"，于是一个本该是空字符串的状态列
    漏成了 NaN，写表时整个报表崩掉。空值就该留空，不该让它伪装成 0。
    """
    if v is None:
        return None
    if isinstance(v, float):
        return None if pd.isna(v) else int(v)
    return v


def load_run_meta(folder):
    """下载时写下的 run_meta.json。老目录没有这个文件，返回 {}。"""
    p = os.path.join(folder, "run_meta.json")
    if not os.path.isfile(p):
        return {}
    try:
        with io.open(p, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def sales_window_start(meta):
    """销售报表这次实际覆盖到多早。信息不全时返回 None。

    平台的"最近 N 个月"是从导出当天往回算的，所以窗口起点 = 导出日 − N 个月。
    """
    n = meta.get("sales_window_months")
    at = meta.get("downloaded_at")
    if not n or not at:
        return None
    try:
        d = pd.Timestamp(at).normalize()
    except Exception:
        return None
    return d - pd.DateOffset(months=int(n))


# 导出顶部那条横幅：「No pudimos mostrar algunos datos. Intenta generando el
# Excel de ventas nuevamente.」保留已有数据，同时把平台的警告传到报表。
VENTAS_INCOMPLETE_RE = re.compile(r"no pudimos mostrar algunos datos", re.I)


def _ventas_header_row(path, limit=16):
    """定位表头行（0-indexed），并捡出顶部的"数据不全"横幅。返回 (行号, 横幅)。

    表头不在固定行：平台会往顶部插公告行，插一条就把表头往下推一行。9/15
    TOOL_TA03 那份正是被这条横幅推到了第 7 行，而这里原来写死第 6 行，于是
    读到的"表头"是分组行「Ventas」——Fecha de venta 整列不存在，日期全成
    NaT，整家店被剔除，而报错只会说"没有日期"，看不出真正原因。

    横幅要单独检出并传给校验、首页和汇总表，避免残缺数据看起来完全正常。
    """
    top = pd.read_excel(path, sheet_name=site.VENTAS_SHEET, header=None, nrows=limit)
    hdr, notice = None, ""
    for i in range(len(top)):
        vals = [str(v).strip() for v in top.iloc[i].tolist()
                if v is not None and str(v) != "nan"]
        if VENTAS_INCOMPLETE_RE.search(" ".join(vals)):
            notice = next((v for v in vals if VENTAS_INCOMPLETE_RE.search(v)), "")
        if hdr is None and V_ORDER_ID in vals and V_DATE in vals:
            hdr = i
    return hdr, notice


def load_ventas(folder):
    path = _first([site.VENTAS_GLOB, "*Ventas*.xlsx"], folder)
    if not path:
        raise IOError("找不到 Ventas 报表：%s" % folder)
    hdr, notice = _ventas_header_row(path)
    if notice:
        print("    [警告] Ventas 导出可能缺数据：%s；文件：%s"
              % (notice, os.path.basename(path)))
    if hdr is None:
        hdr = VENTAS_HEADER_ROW
        print("    [警告] 没在前 16 行里找到表头（认的是「%s」+「%s」），"
              "按第 %d 行读。列名可能不对。"
              % (V_ORDER_ID, V_DATE, VENTAS_HEADER_ROW + 1))
    elif hdr != VENTAS_HEADER_ROW:
        print("    [提示] 表头在第 %d 行（通常是第 %d 行），平台在顶部多插了行，"
              "已按实际行读。" % (hdr + 1, VENTAS_HEADER_ROW + 1))
    df = pd.read_excel(path, sheet_name=site.VENTAS_SHEET, header=hdr)
    df.attrs["incomplete_notice"] = notice
    df["fecha"] = df[V_DATE].map(parse_spanish_date) if V_DATE in df.columns else pd.NaT
    # pandas 2 的 astype(str) 把 NaT 变成字面量 "NaT"，pandas 3 让它保持 NaN。
    # 本文件有 5 处 `ym != "NaT"` 的过滤依赖那个字面量（3613/3669/4492 等），
    # NaN 一律穿过去，最后在 sorted() 里拿 float 和 str 比大小炸掉。这里把
    # "NaT" 填回去，恢复其余代码一直假设的不变量。
    per = df["fecha"].dt.to_period("M")
    df["ym"] = per.astype(str).where(per.notna(), "NaT")
    df["oid"] = idkey(df[V_ORDER_ID]) if V_ORDER_ID in df.columns else ""
    df["pid"] = idkey(df[V_PACK_ID]) if V_PACK_ID in df.columns else ""
    st = df[V_STATUS].astype(str) if V_STATUS in df.columns else pd.Series("", index=df.index)
    df["abnormal"] = st.str.contains(ABNORMAL_RE, case=False, na=False)
    # 卖家承担的优惠券：平台把 Descuentos y bonificaciones 整列导成空白，
    # 只能用 Total 与各明细列的差额倒推。Total 为空的行（换货第二行）不参与。
    tot_raw = pd.to_numeric(df[site.V_TOTAL], errors="coerce") if site.V_TOTAL in df.columns else pd.Series(float("nan"), index=df.index)
    if site.BONIF_COL:
        # 秘鲁：Descuentos y bonificaciones 列有值，正数是平台付给卖家的
        # Mercado Envíos Flex 配送补贴（MercadoPago 里是单独一笔 CASHBACK）。
        # 它是卖家的运费收入，所以并进运费收入列 —— 这样桥 A、损益表、⑦ SKU、
        # ⑧ 趋势全部沿用同一条恒等式，不必为秘鲁多开一列；原值留在 bonif /
        # ship_income_buyer 两列里供说明用。负数（若平台哪天在这一列记卖家
        # 优惠券）不并，照旧由下面的 Total 倒推落进优惠券。
        b = num(df, V_DISCOUNT)
        df["bonif"] = b.where(b > 0, 0.0)
        df["ship_income_buyer"] = num(df, site.V_SHIP_INCOME)
        df[site.V_SHIP_INCOME] = df["ship_income_buyer"] + df["bonif"]
    recomputed = (num(df, site.V_REVENUE) + num(df, site.V_SHIP_INCOME) + num(df, site.V_COMMISSION)
                  + num(df, site.V_SHIP_COST) + num(df, site.V_REFUND))
    coupon = -(tot_raw.fillna(0.0) - recomputed)
    coupon[tot_raw.isna()] = 0.0
    df["coupon"] = coupon
    return path, df


def load_billing(folder):
    """Facturación + Notas de Crédito 全部期间合并。两者是不同的凭证，不会重复。"""
    frames = []
    for pat, doc in (("Reporte_Facturacion_MercadoLibre_*.xlsx", "FAC"),
                     ("Reporte_Notas_Credito_MercadoLibre_*.xlsx", "NC")):
        for p in sorted(glob.glob(os.path.join(folder, pat))):
            if os.path.basename(p).startswith("~$"):
                continue
            t = pd.read_excel(p, sheet_name="REPORT", header=BILLING_HEADER_ROW)
            t["_doc"] = doc
            t["_file"] = os.path.basename(p)
            frames.append(t)
    if not frames:
        raise IOError("找不到 Facturación 报表：%s" % folder)
    df = pd.concat(frames, ignore_index=True)
    df = df[df[F_AMOUNT].notna()].copy()
    df["amount"] = pd.to_numeric(df[F_AMOUNT], errors="coerce").fillna(0.0)
    df["cdate"] = pd.to_datetime(df.get(F_CHARGE_DATE), errors="coerce")
    df["sdate"] = pd.to_datetime(df.get(F_SALE_DATE), errors="coerce")
    df["charge_no"] = idkey(df[F_CHARGE_NO]) if F_CHARGE_NO in df.columns else ""
    rcol = next((c for c in (F_REVERSES,) + F_REVERSES_ALT if c in df.columns), None)
    df["reverses"] = idkey(df[rcol]) if rcol else ""
    df["k_sale"] = idkey(df[F_SALE_NO]) if F_SALE_NO in df.columns else ""
    df["k_pack"] = idkey(df[F_PACK_NO]) if F_PACK_NO in df.columns else ""
    df["channel"] = df[F_CHANNEL].astype(str) if F_CHANNEL in df.columns else ""

    # 冲销行归属：优先用 Cargo que bonifica → Número del cargo 的 ID 关联，
    # 关联不上再退回按名称匹配（REVERSAL_TO_PARENT）。
    by_no = {}
    for cn, det, ch in zip(df["charge_no"], df[F_DETALLE].astype(str), df["channel"]):
        if cn and cn.lower() not in ("nan", "none", ""):
            by_no[cn] = (det, ch)
    parents, pchannels, known = [], [], []
    for det, rev in zip(df[F_DETALLE].astype(str), df["reverses"]):
        if is_reversal(det):
            hit = by_no.get(rev)
            if hit:
                parents.append(hit[0]); pchannels.append(hit[1])
                known.append(hit[0] in FEE_BY_ES)
                continue
            p = REVERSAL_TO_PARENT.get(det)
            parents.append(p); pchannels.append(None); known.append(bool(p) and p in FEE_BY_ES)
        else:
            parents.append(det); pchannels.append(None); known.append(det in FEE_BY_ES)
    df["parent_detalle"] = parents
    df["parent_channel"] = pchannels
    df["known"] = known
    df["is_rev"] = df[F_DETALLE].astype(str).map(is_reversal)
    return df


def load_pagos(folder, year, month, errors=None):
    """当期的账单支付与贷记单。返回 (path, DataFrame) 或 (None, 空表)。"""
    best = None
    for p in sorted(glob.glob(os.path.join(folder, "Reporte_Pagos_Facturas_*.xlsx"))):
        tok = file_month_token(os.path.basename(p))
        if tok == (year, month):
            best = p
            break
    if not best:
        return None, pd.DataFrame()
    try:
        df = pd.read_excel(best, sheet_name="Pagos y notas de crédito",
                           header=PAGOS_HEADER_ROW).dropna(how="all")
    except Exception as e:                      # noqa: BLE001 — 见 unreadable()
        unreadable(errors, best, e)
        return None, pd.DataFrame()
    return best, df


def load_cargos_full(folder, year, month, errors=None):
    """Full 费用明细（七张分表）。仅用于交叉核对，绝不与账单相加。"""
    best = None
    for p in sorted(glob.glob(os.path.join(folder, "Reporte_Cargos_Full_*.xlsx"))):
        if file_month_token(os.path.basename(p)) == (year, month):
            best = p
            break
    if not best:
        return None, {}, 0.0
    sheets, total = {}, 0.0
    try:
        xl = pd.ExcelFile(best)
        for sn in xl.sheet_names:
            d = pd.read_excel(best, sheet_name=sn, header=CARGOS_HEADER_ROW)
            col = next((c for c in d.columns if "Monto del cargo" in str(c)), None)
            if col is None:
                continue
            v = pd.to_numeric(d[col], errors="coerce")
            sheets[sn] = (float(v.sum()), int(v.notna().sum()))
            total += float(v.sum())
    except Exception as e:                      # noqa: BLE001 — 见 unreadable()
        unreadable(errors, best, e)
        return None, {}, 0.0
    return best, sheets, total


def load_storage(folder, p0, p1, errors=None):
    """独立仓储费报表（按实际占用日）。挑与报表月重叠最多的那份。"""
    best, best_overlap = None, -1
    for p in sorted(glob.glob(os.path.join(folder, "*Costos_por_servicio_almacenamiento.xlsx"))):
        rng = storage_file_range(os.path.basename(p))
        if not rng:
            continue
        ov = (min(rng[1], p1) - max(rng[0], p0)).days
        if ov > best_overlap:
            best, best_overlap = p, ov
    if not best or best_overlap < 0:
        return None, 0.0, None
    try:
        s = pd.read_excel(best, sheet_name="Resumen", header=STORAGE_HEADER_ROW)
    except Exception as e:                      # noqa: BLE001 — 见 unreadable()
        unreadable(errors, best, e)
        return None, 0.0, None
    col = "Total" if "Total" in s.columns else s.columns[-1]
    vals = pd.to_numeric(s[col], errors="coerce").dropna()
    accrual = float(vals.iloc[-1]) if len(vals) else 0.0
    return best, accrual, storage_file_range(os.path.basename(best))


# ── 智利：账单周期跨月，三份"按账单出"的报表要按计费日拆回自然月 ──────────
# 墨西哥的 Pagos / Cargos Full / 仓储费报表各自正好是一个自然月，上面三个
# load_* 按文件名挑那一份即可。智利的账单是 8 号到次月 7 号（见 bill_cycles），
# 一个会计月横跨两张账单，所以下面三个函数读**两期**文件，再按每一行自己的
# 日期切出会计月那一段。只在 is_cl() 时调用，墨西哥路径不经过这里。

def payable_of(frame):
    """一组账单行里"要单独付钱"的部分：(No 通道费用, 冲销掉 No 通道费用的行)。

    冲销优先按 Cargo que bonifica/anula 关联到的原费用的通道判断，一行都关联
    不上时退回按科目字典的"账单级"判断 —— 与 derive() 里桥 B 的口径相同。
    """
    no = float(frame[frame["channel"] == CH_NO]["amount"].sum())
    rev = frame[frame["is_rev"] & (frame["parent_channel"] == CH_NO)]
    if not len(rev):
        rev = frame[frame["is_rev"] & frame["parent_detalle"].map(
            lambda d: FEE_BY_ES.get(str(d), (None, None, None))[2] is False).astype(bool)]
    return no, float(rev["amount"].sum())


def _cycle_file(pattern, folder, y, mth):
    for p in sorted(glob.glob(os.path.join(folder, pattern))):
        if not os.path.basename(p).startswith("~$") \
                and file_month_token(os.path.basename(p)) == (y, mth):
            return p
    return None


def load_pagos_cycles(folder, cycles, errors=None):
    """每个账单周期的 Reporte_Pagos_Facturas：汇总页 + "Detalle de Pagos del mes"。

    明细页逐行写着"这笔付款的哪一部分付了哪一张费用（Número del cargo /
    Fecha del cargo）"，所以能把一笔付款精确地拆到自然月：DAB_CL 2026-09 实测
    Septiembre 账单的付款 583,596 里 404,998 付的是 9/1–9/7 的费用，Octubre
    账单的付款 2,956,626 里 2,334,297 付的是 9/8–9/30 的费用。
    """
    out = []
    for y, mth, c0, c1 in cycles:
        path = _cycle_file("Reporte_Pagos_Facturas_*.xlsx", folder, y, mth)
        P, D = pd.DataFrame(), pd.DataFrame()
        if path:
            try:
                P = pd.read_excel(path, sheet_name="Pagos y notas de crédito",
                                  header=PAGOS_HEADER_ROW).dropna(how="all")
            except Exception as e:              # noqa: BLE001 — 见 unreadable()
                # 汇总页都读不出来 = 整份文件坏了，当没下载；只缺明细页
                # （下面）是另一回事，_derive_pagos_cycles 自己会说明。
                unreadable(errors, path, e)
                P, path = pd.DataFrame(), None
        if path:
            try:
                D = pd.read_excel(path, sheet_name="Detalle de Pagos del mes",
                                  header=PAGOS_HEADER_ROW).dropna(how="all")
            except Exception:
                D = pd.DataFrame()
        out.append({"bill": (y, mth), "start": c0, "end": c1, "path": path, "P": P, "D": D})
    return out


def _date_col(df, names=("Fecha de cargo", "Fecha del cargo")):
    return next((c for c in df.columns if str(c).strip() in names), None)


def load_cargos_full_cycles(folder, cycles, p0, p1, errors=None):
    """Full 费用明细，两期文件里计费日落在会计月内的行。

    返回 (路径们, {分表: (金额, 行数)}, 合计, 缺的账单名)。
    """
    paths, sheets, total, missing = [], {}, 0.0, []
    for y, mth, c0, c1 in cycles:
        path = _cycle_file("Reporte_Cargos_Full_*.xlsx", folder, y, mth)
        if not path:
            missing.append("%s %d" % (MONTH_FULL_ES[mth], y))
            continue
        try:
            frames = [(sn, pd.read_excel(path, sheet_name=sn, header=CARGOS_HEADER_ROW))
                      for sn in pd.ExcelFile(path).sheet_names]
        except Exception as e:                  # noqa: BLE001 — 见 unreadable()
            unreadable(errors, path, e)
            missing.append("%s %d" % (MONTH_FULL_ES[mth], y))
            continue
        paths.append(path)
        for sn, d in frames:
            col = next((c for c in d.columns if "Monto del cargo" in str(c)), None)
            dc = _date_col(d)
            if col is None or dc is None:
                continue
            dt = pd.to_datetime(d[dc], errors="coerce")
            v = pd.to_numeric(d[col], errors="coerce")
            keep = v.notna() & (dt >= p0) & (dt <= p1)
            s, n = sheets.get(sn, (0.0, 0))
            sheets[sn] = (s + float(v[keep].sum()), n + int(keep.sum()))
            total += float(v[keep].sum())
    return paths, sheets, total, missing


# 仓储费报表（Detalle 页）按"占用日"逐日给出每个 SKU 的费用；账单的仓储费
# 计费日比占用日晚一天（DAB_CL：报表 08/09 合计 23,410.25 ↔ 账单 9/9 的
# 23,409；报表 09/09 22,882.46 ↔ 账单 9/10 22,880），差的几个比索是逐行取整。
STORAGE_BILL_LAG_DAYS = 1


def load_storage_daily(folder, p0, p1, errors=None):
    """仓储费报表里占用日落在会计月内的逐日费用之和。

    返回 (路径们, 合计, (覆盖起, 覆盖止) 或 None, 覆盖天数)。智利的仓储报表
    也按 8 号切期，下载的是哪几期就只覆盖哪几天 —— 覆盖不全时如实说出来，
    不拿半个月的数去和整月比。
    """
    paths, total, days = [], 0.0, []
    for p in sorted(glob.glob(os.path.join(folder, "*Costos_por_servicio_almacenamiento.xlsx"))):
        rng = storage_file_range(os.path.basename(p))
        if not rng or rng[1] < p0 or rng[0] > p1:
            continue
        try:
            t = pd.read_excel(p, sheet_name="Detalle", header=STORAGE_HEADER_ROW)
        except Exception as e:                  # noqa: BLE001 — 见 unreadable()
            unreadable(errors, p, e)
            continue
        hit = False
        for c in t.columns:
            dt = pd.to_datetime(str(c), format="%d/%m/%Y", errors="coerce")
            if pd.isna(dt) or dt < p0 or dt > p1:
                continue
            # 每个 SKU 两行：费用行（数字）和件数行（"95 u"），后者转数字为 NaN
            total += float(pd.to_numeric(t[c], errors="coerce").sum())
            days.append(dt)
            hit = True
        if hit:
            paths.append(p)
    if not days:
        return paths, 0.0, None, 0
    days = sorted(set(days))
    return paths, total, (days[0], days[-1]), len(days)


def load_returns(folder, errors=None):
    path = _first(["Returns*.xlsx"], folder)
    if not path:
        return None, pd.DataFrame()
    try:
        r = pd.read_excel(path, sheet_name="Triages", header=RETURNS_HEADER_ROW)
    except Exception as e:                      # noqa: BLE001 — 见 unreadable()
        unreadable(errors, path, e)
        return None, pd.DataFrame()
    r = r.dropna(how="all")
    if "Order" in r.columns and "Número de orden" not in r.columns:
        r = r.rename(columns=RETURNS_EN_COLUMNS)
        for col, table in RETURNS_EN_VALUES.items():
            if col in r.columns:
                r[col] = r[col].map(lambda v, t=table: t.get(v, v) if isinstance(v, str) else v)
    if "Fecha de revisión" in r.columns:
        r["fr"] = pd.to_datetime(r["Fecha de revisión"], format="%d-%m-%Y", errors="coerce")
        if len(r) and not r["fr"].notna().any():
            eg = [str(v) for v in r["Fecha de revisión"].dropna().head(3)]
            print("    [警告] Returns 的「Fecha de revisión」没有一个值能按 "
                  "%%d-%%m-%%Y 解析（样例：%s）。按质检日的统计会失真。"
                  % ("、".join(eg) or "整列为空"))
    else:
        # 取列名要在加 fr 之前 —— 否则报错里会列出我们自己加的那一列，
        # 看的人会以为文件里真有它。
        cols = [str(c) for c in r.columns if not str(c).startswith("Unnamed")]
        r["fr"] = pd.NaT
        print("    [警告] Returns 的 Triages 页没有「Fecha de revisión」列，"
              "表头可能不在第 %d 行。实际列名：%s"
              % (RETURNS_HEADER_ROW + 1, "、".join(cols[:8]) or "（一个都没有）"))
    return path, r


def date_span(s):
    """把一列日期渲染成「起 ~ 止」；没有一个能用的日期就返回 None。

    退货的质检日 `fr` 由固定格式 %d-%m-%Y 加 errors="coerce" 解析，列缺了或
    格式变了就整列 NaT。而 NaT 不支持 strftime —— 直接 ValueError。

    出事的两处守卫写的都是 len(returns_all) 非空。**"有行"和"有能用的日期"
    是两回事**，这个混淆在这个文件里已经犯过三次（build_context 的 cutoff、
    店铺可用性、这里）。9/24 EWTTO_SM 就栽在这儿：十二条校验全过、结果都打
    出来了，倒在写工作簿的最后一步，磁盘上留下的还是上一次的旧报表。
    """
    if s is None or not len(s):
        return None
    lo, hi = s.min(), s.max()
    if pd.isna(lo) or pd.isna(hi):
        return None
    return "%s ~ %s" % (lo.strftime("%Y-%m-%d"), hi.strftime("%Y-%m-%d"))


STOCK_SHEET = "Resumen"
# Resumen 页的表头是**两到三层**：第一层是分组（"Unidades en Full"），
# 下面才是真正的列名（"Aptas para vender"）。而且有两列都叫
# "Pendientes de ingreso"（一列是在途，一列是"建议动作"分类），
# 只能靠分组名区分。所以按 (分组, 列名) 匹配，不按列位置 —— 平台一改列序
# 位置法就会静默错位，而且错得看不出来。
STOCK_COLS = {
    "código ml": "ml",
    "sku": "sku",
    "producto": "title",
    "estado de la publicación": "estado",
    "ofrece full": "full",
    "stock promedio últimos 30 días (u.)": "avg_stock",
    "unidades con antigüedad": "aged",
    "tiempo hasta agotar stock": "runout",
    "en transferencia": "transfer",
    "devueltas por el comprador": "returned",
    "aptas para vender": "sellable",
    "no aptas para vender": "unsellable",
    "extraviadas": "lost",
    "en revisión": "in_review",
    "ventas canceladas": "cancelled_u",
    "unidades que ocupan espacio en full": "occupying",
    "buena calidad": "good",
    "para impulsar ventas": "boost",
    "para poner en venta": "to_list",
    "para evitar descarte": "discard",
    "ventas": "sales30_amt",
    "unidades vendidas": "sales30_u",
    "urgencia de envío": "urgency",
}
STOCK_NUM = ["avg_stock", "aged", "transfer", "returned", "sellable", "unsellable",
             "lost", "in_review", "cancelled_u", "occupying", "in_transit",
             "act_pending", "good", "boost", "to_list", "discard",
             "sales30_amt", "sales30_u"]


def _hnorm(v):
    """表头单元格 → 匹配用的短名：只取第一行、去空白、转小写。

    平台把说明文字塞在同一个单元格里（"Unidades con antigüedad\\n¿Por qué...?"），
    整串拿去匹配会被这些说明拖着变。
    """
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return str(v).split("\n")[0].strip().lower()


def load_stock(folder, errors=None):
    """Full 仓库存报表（stock_general_full）的 Resumen 页。

    这一页本身就是完整的 SKU 级主表，其余四页只是它的切片，所以只读这一页。
    注意：它是**下载当时的快照**，不是月末余额 —— 与损益表不同期，不能相加。

    返回 (文件路径, 快照文字, DataFrame)。没有报表时返回空 DataFrame 而不报错：
    这个报表要单独的权限，不少店根本拉不到。
    """
    path = _first(["stock_general_full*.xlsx"], folder)
    if not path:
        return None, "", pd.DataFrame()
    try:
        raw = pd.read_excel(path, sheet_name=STOCK_SHEET, header=None)
    except Exception as e:                      # noqa: BLE001 — 见 unreadable()
        unreadable(errors, path, e)
        return None, "", pd.DataFrame()

    snap = ""
    hdr = None
    for i in range(len(raw)):
        row = [_hnorm(v) for v in raw.iloc[i]]
        if not snap:
            for v in raw.iloc[i]:
                if isinstance(v, str) and v.strip().startswith("Actualizado el"):
                    snap = v.strip()
        if "código ml" in row:
            hdr = i
            break
    if hdr is None:
        return path, snap, pd.DataFrame()

    ncol = raw.shape[1]
    group = [""] * ncol            # 第一层分组，向右填充
    sub = [""] * ncol              # 最靠下的那一层才是真列名
    cur = ""
    for j in range(ncol):
        g = _hnorm(raw.iloc[hdr, j])
        if g:
            cur = g
        group[j] = cur
        for k in (hdr + 2, hdr + 1, hdr):
            if k < len(raw):
                s = _hnorm(raw.iloc[k, j])
                if s:
                    sub[j] = s
                    break

    field = {}
    for j in range(ncol):
        name = STOCK_COLS.get(sub[j])
        if sub[j] == "pendientes de ingreso":
            # 两列同名，靠分组区分：一列是在途，一列是"建议动作"里的在途
            name = "in_transit" if "camino" in group[j] else "act_pending"
        if name and name not in field:
            field[name] = j

    body = raw.iloc[hdr + 1:]
    if "ml" not in field:
        return path, snap, pd.DataFrame()
    body = body[body.iloc[:, field["ml"]].notna()]     # 末尾的合计行没有 ML 码，正好排除

    out = pd.DataFrame(index=range(len(body)))
    for name, j in field.items():
        out[name] = body.iloc[:, j].values
    for c in STOCK_NUM:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0) if c in out.columns \
            else 0.0
    for c in ("sku", "title", "estado", "runout", "urgency", "full", "ml"):
        # 先 fillna 再 astype。pandas 2 的 astype(str) 把 NaN 变成字面量 "nan"，
        # pandas 3 则**保留 NaN** —— 只靠 replace({"nan": ""}) 在 pandas 3 下
        # 完全不起作用，NaN 会一路漏到写单元格时 int(NaN) 炸掉。
        out[c] = (out[c].fillna("").astype(str).str.strip()
                  if c in out.columns else "")
        out[c] = out[c].replace({"nan": "", "NaT": "", "None": ""})
    out["total_u"] = (out["in_transit"] + out["transfer"] + out["returned"]
                      + out["sellable"] + out["unsellable"] + out["lost"]
                      + out["in_review"] + out["cancelled_u"])
    return path, snap, out


def stock_snapshot_date(snap, folder):
    """把 'Actualizado el 8 de septiembre a las 23:40 hs.' 解析成日期。

    平台**不写年份**，所以用下载目录名里的年份（downloads/<店>/YYYYMMDD_HHMMSS）。
    跨年下载时快照月份会大于目录月份，这时年份减一。
    """
    m = re.search(r"(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)", snap or "")
    if not m:
        return pd.NaT
    mon = SPANISH_MONTHS.get(m.group(2).lower())
    if not mon:
        return pd.NaT
    base = os.path.basename(os.path.normpath(folder))
    try:
        ref = pd.Timestamp(int(base[:4]), int(base[4:6]), int(base[6:8]))
    except Exception:
        ref = pd.Timestamp(datetime.date.today())
    for y in (ref.year, ref.year - 1):
        try:
            d = pd.Timestamp(y, mon, int(m.group(1)))
        except ValueError:
            return pd.NaT
        if d <= ref + pd.Timedelta(days=1):
            return d
    return pd.NaT


MP_COLUMNS = ("TRANSACTION_DATE", "TRANSACTION_TYPE", "SOURCE_ID", "REAL_AMOUNT")


def _read_mp_csv(path, errors, required=MP_COLUMNS):
    """读一份 MercadoPago 流水（settlement_v2 / reserve-release）。读不出来、或
    连必需的列都没有（平台把网页存成了 .csv 时就是这样）→ 记录并返回 None。"""
    try:
        d = pd.read_csv(path, sep=";")
    except Exception as e:                      # noqa: BLE001 — 见 unreadable()
        unreadable(errors, path, e)
        return None
    lack = [c for c in required if c not in d.columns]
    if lack:
        unreadable(errors, path, "缺少列 %s（不是 MercadoPago 流水的格式）" % "、".join(lack))
        return None
    return d


def load_settlement(folder, p0, p1, errors=None):
    """MercadoPago 结算流水（可选）。用于第三方交叉验证；很多店没有报表权限。"""
    path = _first(["settlement_v2*.csv"], folder)
    if not path:
        return None, None
    s = _read_mp_csv(path, errors)
    if s is None:
        return None, None
    try:
        s["dt"] = pd.to_datetime(s["TRANSACTION_DATE"], errors="coerce", utc=True)
    except Exception:
        return path, None
    s = s[(s["dt"].dt.tz_localize(None) >= p0) & (s["dt"].dt.tz_localize(None) <= p1 + pd.Timedelta(days=1))]
    if not len(s):
        return path, None
    g = lambda t: s[s["TRANSACTION_TYPE"] == t]
    st = g("SETTLEMENT")
    out = {
        "n": int(len(st)),
        "trans": float(st["TRANSACTION_AMOUNT"].sum()),
        "fee": float(st["FEE_AMOUNT"].sum()),
        "tax": float(st["TAXES_AMOUNT"].sum()),
        "real_settlement": float(st["REAL_AMOUNT"].sum()),
        "refund_real": float(s[s["TRANSACTION_TYPE"].isin(["REFUND", "DISPUTE"])]["REAL_AMOUNT"].sum()),
        "real_all": float(s["REAL_AMOUNT"].sum()),
    }
    pos = st[st["TRANSACTION_AMOUNT"] > 0]
    if len(pos):
        rate = (-pos["TAXES_AMOUNT"] / pos["TRANSACTION_AMOUNT"])
        out["rate_mean"] = float(rate.mean())
        out["rate_std"] = float(rate.std()) if len(pos) > 1 else 0.0
        out["rate_mode_n"] = int((rate.round(4) == round(rate.median(), 4)).sum())
        out["rate_n"] = int(len(pos))
    ref = s[s["TRANSACTION_TYPE"].isin(["REFUND", "DISPUTE"])]
    out["refund_detail"] = " ＋ ".join(
        "%s %s" % (t, fmt(v)) for t, v in
        zip(ref["TRANSACTION_TYPE"], ref["REAL_AMOUNT"])) or "无"
    return path, out


def mp_by_order(A, B, folder, errors=None):
    """MercadoPago 逐单交叉验证：每个订单的销售报表 Total 对它自己那笔收款的净额。

    以前是"本月按交易日的流水合计"对"本月按下单日的销售合计"，两边不是同一批
    钱（9 月的订单常在 10 月到账，8 月订单的退款落在 9 月），11 家店差 3 万到
    79 万，永远不过。逐单比就对得上：2026-10-14 在 10 家店的真实数据上，3,977
    个订单里 3,976 个能对上收款号，96.5% 分毫不差。

    关联链：销售报表订单（# de venta / Orden de compra）→ 账单行（Número de
    venta / Número de paquete，键规则见下面的 credit_note_audit）→ 账单的 Pago
    （收款号）→ 结算流水的 SOURCE_ID。收款的净额取该收款号**所有日期**的
    REAL_AMOUNT 之和（结算、退款、争议都算），所以跨月到账也能对上。

    返回 None 表示没有结算流水；否则 dict：
      linked / agree / unlinked / no_cash 计数，sum_v / sum_mp（可比订单的
      两边合计），differ（[(订单号, 销售报表 Total, MercadoPago 净额)]，按差额
      绝对值从大到小），span（流水覆盖的日期）。
    """
    p = _first(["settlement_v2*.csv"], folder)
    if not p:
        return None
    s = _read_mp_csv(p, errors)
    if s is None:
        return None
    try:
        dt = pd.to_datetime(s["TRANSACTION_DATE"], errors="coerce", utc=True)
    except Exception:
        return None
    cash = {}
    for a, b in zip(s["SOURCE_ID"].map(_srcid), pd.to_numeric(
            s["REAL_AMOUNT"], errors="coerce").fillna(0.0)):
        if a:
            cash[a] = cash.get(a, 0.0) + float(b)
    out = {"linked": 0, "agree": 0, "unlinked": 0, "no_cash": 0,
           "sum_v": 0.0, "sum_mp": 0.0, "differ": [],
           "span": "%s ~ %s" % (str(dt.min())[:10], str(dt.max())[:10])}
    if not len(A) or "Pago" not in B.columns:
        out["unlinked"] = int(len(A))
        return out
    ids = {}
    for i, oid, pid in zip(A.index, A["oid"], A["pid"]):
        for k in (str(oid), str(pid)):
            if k not in ("", "nan"):
                ids[k] = i
    link = {}
    for ks, kp, pago in zip(B["k_sale"], B["k_pack"], B["Pago"]):
        pg = _srcid(pago)
        if not pg:
            continue
        for k in (str(ks).strip(), str(kp).strip()):
            if k in ids:
                link.setdefault(ids[k], set()).add(pg)
    out["unlinked"] = int(len(A) - len(link))
    tot = num(A, site.V_TOTAL)
    # 一次买了几件（Paquete de varios productos = Sí）：销售报表把钱全记在包裹那
    # 一行（# de venta 是 2000014… 号段），每件商品自己那一行金额全是 0；而平台
    # 是按件收款的。2026-10-14 实测 12 家店"销售报表 0、平台有到账"的 67 单，
    # 就是这些子行：BOCINA_TA02 包裹行 336.00 = 两件的收款 206.00 + 130.00。
    # 包裹行已经和同一批收款对上了，子行再比一次只会凭空多出差额，所以跳过。
    child = set()
    if V_PACK_FLAG in A.columns:
        flag = A[V_PACK_FLAG].astype(str).str.strip().str.lower().isin(("sí", "si"))
        child = set(A.index[flag & (tot.abs() <= site.CN_TOL)
                            & (num(A, site.V_REVENUE).abs() <= site.CN_TOL)])
    out["pack_children"] = len([i for i in link if i in child])
    for i, pagos in link.items():
        if i in child:
            continue
        got = [cash[x] for x in pagos if x in cash]
        if not got:
            out["no_cash"] += 1
            continue
        vt, mp = float(tot.loc[i]), float(sum(got))
        out["linked"] += 1
        out["sum_v"] += vt
        out["sum_mp"] += mp
        if abs(mp - vt) <= site.CN_TOL:
            out["agree"] += 1
        else:
            out["differ"].append((str(A.at[i, "oid"]), vt, mp))
    out["differ"].sort(key=lambda x: -abs(x[2] - x[1]))
    return out


# ── 贷记明细 × 销售报表 逐单交叉比对 ────────────────────────────────
# 关联键有个坑：账单的 Número de venta（2000017/18xxx）与销售报表的
# # de venta（2000014xxx）**不是同一个号段**，可靠的桥是
# Número de paquete ↔ 销售报表的 # de venta / Orden de compra。
# 所以不能"k_sale 非空就取 k_sale"，必须挑能对上销售报表的那个键 ——
# 实测这么改，12 家店的匹配率从 132/464 升到 411/464（BOCINA_TA02 从 0 升到 11）。
PARTIAL_RE = re.compile(r"parcial", re.I)


def _srcid(v):
    """MercadoPago 收款号规整成字符串。

    同一列在不同店铺导出成三种 dtype：float64、字符串、甚至非数字的
    "000dj6x3ef"。直接 "%d" % float(v) 会在字符串上抛异常，把整份流水
    连带丢掉（实测 12 家店里 10 家因此拿不到现金凭证）。逐值判断，
    认不出来的原样保留 —— 匹配不上只是少一条凭证，不该拖垮其余的。
    """
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, float):
        return "%d" % v
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def load_mp_cash(folder, errors=None):
    """MercadoPago 逐笔到账净额 {收款号: 净额}，用来给比对结论找现金凭证。

    两个来源都认：reserve-release（明细最全，含 mediation/refund 分解）优先，
    没有就退回 settlement_v2。两者都要单独的报表权限，拿不到就返回空 dict，
    比对照常做，只是结论降级为"无现金凭证可核"。
    """
    out = {}
    p = _first(["reserve-release-*.csv"], folder)
    d = _read_mp_csv(p, errors, required=("SOURCE_ID", "NET_CREDIT_AMOUNT",
                                           "NET_DEBIT_AMOUNT")) if p else None
    if d is not None:
        try:
            net = (pd.to_numeric(d["NET_CREDIT_AMOUNT"], errors="coerce").fillna(0.0)
                   - pd.to_numeric(d["NET_DEBIT_AMOUNT"], errors="coerce").fillna(0.0))
            for a, b in zip(d["SOURCE_ID"].map(_srcid), net):
                if a:
                    out[a] = out.get(a, 0.0) + float(b)
        except Exception:
            out = {}
    if out:
        return out
    p = _first(["settlement_v2*.csv"], folder)
    d = _read_mp_csv(p, errors, required=("SOURCE_ID", "REAL_AMOUNT")) if p else None
    if d is not None:
        try:
            amt = pd.to_numeric(d["REAL_AMOUNT"], errors="coerce").fillna(0.0)
            for a, b in zip(d["SOURCE_ID"].map(_srcid), amt):
                if a:
                    out[a] = out.get(a, 0.0) + float(b)
        except Exception:
            return {}
    return out


def credit_note_audit(S, folder):
    """按订单号把贷记明细和销售报表逐单对照，找三类现有校验看不见的异常。

    现有 7 条校验的盲区：桥A 只在销售报表内部勾稽（残差恒为 0），桥B 只比账单
    总额，MercadoPago 那条只比店铺汇总差额（三种原因混在一起分不开）。下面三类
    都是逐单的，落在盲区里：

      orphan  —— 账单为某单计了费，销售报表里根本查无此单，且计提后没有被冲平。
                 绝大多数"查无此单"是下单后立刻作废（计费又全额冲销，净额为 0），
                 无害；净额不为零的才是真的被收了钱。
      unrev   —— 取消单计了费却没有任何冲销。注意发货前取消的单平台压根不下账单，
                 所以"没有贷记单"本身不是问题，必须回查账单确认确实计过费。
      partial —— 佣金被**部分**冲销（Estado del cargo = Anulado parcialmente）。
                 平台按退款比例 p 同时退还佣金与代扣税，但销售报表的
                 「退款与取消」列在这条路径上会把原佣金**重复扣一次**：
                     应退       = p × (GMV − 佣金 − 代扣税)
                     销售报表写的 = 应退 + 原佣金
                 即销售报表少算了一个原佣金的利润。已用 MercadoPago 现金流水在
                 6 家店 10 单上逐单证实，每一单差额都精确等于该单原佣金。
    """
    A, V, B = S["ventas"], S["ventas_all"], S["billing_all"]
    res = {"orphan": [], "unrev": [], "partial": [], "mismatch": [],
           "orphan_total": 0.0, "unrev_total": 0.0, "partial_total": 0.0,
           "matched": 0, "order_rows": 0, "has_cash": False}
    if not len(B) or "Estado del cargo" not in B.columns:
        return res

    ids = set(V["oid"]) | set(V["pid"])
    ids.discard(""); ids.discard("nan")
    cash = load_mp_cash(folder, S.get("load_errors"))
    res["has_cash"] = bool(cash)

    def okey(ks, kp):
        """账单行 → 销售报表里的订单键；对不上返回 (None, 账单自己的键)。"""
        ks, kp = str(ks).strip(), str(kp).strip()
        if ks in ids:
            return ks, ks
        if kp in ids:
            return kp, kp
        raw = kp if kp not in ("", "nan") else (ks if ks not in ("", "nan") else "")
        return None, raw

    # ---- orphan：账单有、销售报表无，且没被冲平 ----
    agg = {}
    for ks, kp, amt, doc in zip(B["k_sale"], B["k_pack"], B["amount"], B["_doc"]):
        hit, raw = okey(ks, kp)
        if not raw:
            continue                       # 账单级费用（广告等），本来就没有订单号
        res["order_rows"] += 1
        if hit:
            res["matched"] += 1
        else:
            a = agg.setdefault(raw, [0.0, 0])
            a[0] += float(amt); a[1] += 1
    for k, (amt, cnt) in sorted(agg.items(), key=lambda x: -abs(x[1][0])):
        if abs(amt) > site.CN_TOL:
            res["orphan"].append((k, amt, cnt))
            res["orphan_total"] += amt

    # ---- unrev：取消单计了费却没冲销 ----
    nc_keys = set()
    for ks, kp, doc in zip(B["k_sale"], B["k_pack"], B["_doc"]):
        if doc != "NC":
            continue
        hit, _ = okey(ks, kp)
        if hit:
            nc_keys.add(hit)
    if len(A) and V_STATUS in A.columns:
        can = A[A[V_STATUS].astype(str).str.contains("cancel", case=False, na=False)]
        for _, x in can.iterrows():
            keys = [z for z in (str(x["oid"]), str(x["pid"])) if z not in ("", "nan", " ")]
            if set(keys) & nc_keys:
                continue
            b = B[B["k_sale"].astype(str).isin(keys) | B["k_pack"].astype(str).isin(keys)]
            net = float(b["amount"].sum()) if len(b) else 0.0
            if len(b) and abs(net) > site.CN_TOL:     # 真的计过费又没退
                res["unrev"].append((keys[0], net, str(x.get(V_STATUS, "")),
                                     x.get("fecha"), str(x.get(V_SKU, ""))))
                res["unrev_total"] += net

    # ---- partial：佣金被部分冲销，销售报表重复扣了一次原佣金 ----
    par = B[(B["_doc"] != "NC") & (B["Detalle"].astype(str) == "Cargo por venta")
            & B["Estado del cargo"].astype(str).str.contains(PARTIAL_RE, na=False)]
    for _, x in par.iterrows():
        com = float(x["amount"])
        if com <= 0:
            continue
        back = -float(B[B["reverses"].astype(str) == str(x["charge_no"])]["amount"].sum())
        p = back / com
        hit, _ = okey(x["k_sale"], x["k_pack"])
        if not hit:
            continue
        v = A[(A["oid"] == hit) | (A["pid"] == hit)]
        if not len(v):
            continue                        # 订单不在本会计期，留给它自己的月份处理
        gmv = float(num(v, site.V_REVENUE).sum())
        tax = float(-num(v, site.V_COMMISSION).sum()) - com
        vref = float(-num(v, site.V_REFUND).sum())
        vtot = float(num(v, site.V_TOTAL).sum())
        want = p * (gmv - com - tax)        # 闭式：退多少货款就退多少比例的费用
        expect = want + com                 # 销售报表在这条路径上实际会写的数
        # 现金凭证：MercadoPago 实际净额与销售报表 Total 的差应恰好等于原佣金
        pago = _srcid(x.get("Pago"))
        real = cash.get(pago)
        proof = ("无 MercadoPago 流水" if not cash else
                 "流水缺此笔" if real is None else
                 "现金凭证确认" if abs((real - vtot) - com) <= site.CN_TOL else
                 "现金凭证不符（差 %s）" % fmt(real - vtot - com))
        # 认定顺序：现金凭证 > 公式吻合。流水是独立凭证，直接证明销售报表的
        # Total 比实际净额少了一个原佣金；公式只是"长得像这个模式"。两者冲突时
        # 以流水为准 —— 实测有两单公式差几块钱，流水却分毫不差。
        by_cash = real is not None and abs((real - vtot) - com) <= site.CN_TOL
        row = (hit, p, com, gmv, tax, vref, want, vtot, real, proof,
               str(x.get("Fecha de venta", ""))[:10], str(v[V_SKU].iloc[0]),
               "现金凭证" if by_cash else "按规律推定")
        if by_cash or abs(vref - expect) <= site.CN_TOL:
            res["partial"].append(row)
            res["partial_total"] += com
        else:
            res["mismatch"].append(row + (vref - expect,))
    return res
