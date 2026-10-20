# -*- coding: utf-8 -*-
"""The experiment's two workbooks, in the report family's look (core/report_kit,
core/report_family): navy titles, blue section bands, grey formula rows, Chinese
labels with the platform's own term next to them.

  <store>_<YYYY-MM>_账单口径实验.xlsx   ① 说明 ② 账单口径损益 ③ 与现行订单口径对比
                                       ④ 费用明细 ⑤ 销售与冲销明细 ⑥ 口径稳定性（多次下载）
  <store>_<YYYY-MM>_口径稳定性.xlsx     ① 说明 ② 稳定性总表 ③ 订单变化明细 ④ 账单行变化明细

Everything is written from the plain data of report.compute() / stability.analyse();
nothing here reads the MercadoLibre site globals."""

import os

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from core import report_family as RF
from core import report_kit as K
from core.report_kit import BOX, F, fill, wrap
from platforms.mercadolibre.report.catalog import CHANNEL_MEANING, CHANNEL_ZH, FEE_CATALOG

from experiments.ml_invoice_basis.common import snapshot_label
from experiments.ml_invoice_basis.compare import CAT, order_key, rec_summary
from experiments.ml_invoice_basis.invoice import LINK_ORDER, bill_name

SHEETS = ["① 说明", "② 账单口径损益", "③ 与现行订单口径对比", "④ 费用明细",
          "⑤ 销售与冲销明细", "⑥ 口径稳定性（多次下载）"]
STAB_SHEETS = ["① 说明", "② 稳定性总表", "③ 订单变化明细", "④ 账单行变化明细"]
UNAVAILABLE = "无法得到"
BUCKET_ZH = {"commission": "佣金", "shipping": "配送", "returns": "退货", "ads": "广告",
             "storage": "仓储", "fulfillment": "Full 履约", "penalty": "罚金",
             "subscription": "订阅", "misc": "其他（待确认）"}


# ════════════════════════════════════════════════════════════════════════
# small writers
# ════════════════════════════════════════════════════════════════════════

def _d(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return str(v)


def _span(view):
    return "%s–%s" % (view["start"].strftime("%Y-%m-%d"), view["end"].strftime("%m-%d"))


def _put(ws, r, c, v, fmt=None, bold=False, color="000000", fillc=None, wrap_text=False,
         border=True, italic=False, size=10):
    if v is not None and not isinstance(v, (str, int, float)):
        v = _d(v)
    if isinstance(v, float):
        v = None if pd.isna(v) else round(v, 6)     # no 3.6e-12 float noise on the page
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = F(size, bold, color, italic)
    if fmt:
        cell.number_format = fmt
    if fillc:
        cell.fill = fill(fillc)
    if wrap_text:
        cell.alignment = wrap("top")
    if border:
        cell.border = BOX
    return cell


def _head(ws, r, cols, height=30):
    K.header_cells(ws, r, cols)
    ws.row_dimensions[r].height = height
    return r + 1


def _band(ws, r, text, ncols):
    return RF.section(ws, r, text, ncols)


def _note(ws, r, text, ncols, color=K.C_GREY, bold=False):
    c = ws.cell(row=r, column=1, value=text)
    c.font = F(10, bold, color)
    c.alignment = wrap("top")
    if ncols > 1:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    ws.row_dimensions[r].height = K.est_height(text, per_line=max(40, ncols * 14))
    return r + 1


def _fmt_of(kind, money):
    return {"m": money, "n": K.INT, "p": K.PCT, "r": "0.0000"}.get(kind)


def _table(ws, r, cols, rows, money):
    """cols: [(header, kind)] kind t/m/n/p/r/d; rows: lists of values."""
    r = _head(ws, r, [c[0] for c in cols])
    for row in rows:
        for j, ((_h, kind), v) in enumerate(zip(cols, row), 1):
            if kind == "d":
                v = _d(v)
            _put(ws, r, j, v, _fmt_of(kind, money), wrap_text=(kind == "w"))
        r += 1
    return r


def _value(ws, r, c, v, money, kind="m", bold=False, fillc=None):
    if v is None:
        return _put(ws, r, c, UNAVAILABLE, color=K.C_RED, fillc=fillc)
    return _put(ws, r, c, v, _fmt_of(kind, money), bold=bold, fillc=fillc)


def _diff_formula(a, b, r):
    return '=IF(AND(ISNUMBER(%s%d),ISNUMBER(%s%d)),%s%d-%s%d,"")' % (b, r, a, r, b, r, a, r)


# ════════════════════════════════════════════════════════════════════════
# ① 说明
# ════════════════════════════════════════════════════════════════════════

def _periods_text(D):
    P, C = D["invs"]["period"]["view"], D["invs"]["calendar"]["view"]
    y, m = int(D["month"][:4]), int(D["month"][5:7])
    if D["cycle_day"] == 1:
        p = ("%s 账单：计费日 %s 至 %s。%s的账单按自然月出，所以账单期就是自然月，两种读法数字相同。"
             % (bill_name((y, m)), _d(P["start"]), _d(P["end"]), D["country"]))
    else:
        p = ("%s 账单：计费日 %s 至 %s（本店账单切期日 = 每月 %d 号，按本店自己的账单/仓储文件推出；"
             "%s的账单不是自然月，名为“%s”的账单覆盖上月 %d 号到本月 %d 号）。"
             % (bill_name((y, m)), _d(P["start"]), _d(P["end"]), D["cycle_day"], D["country"],
                bill_name((y, m)), D["cycle_day"], D["cycle_day"] - 1))
    c = ("计费日 %s 至 %s 的账单行，取自 %s。" % (_d(C["start"]), _d(C["end"]),
                                       " + ".join(bill_name(b) for b in C["bills"])))
    return p, c


def readme_sheet(wb, D, A=None):
    ws = wb.create_sheet(SHEETS[0])
    y, mth = int(D["month"][:4]), int(D["month"][5:7])
    K.title_rows(ws, "美客多 账单口径实验 · %s · %d年%d月" % (D["store"], y, mth),
                 "实验报表，不是月度财务报表，不参与每月的运行。同一店同一月的正式报表仍是 %s_%s.xlsx。"
                 % (D["store"], D["month"]))
    issues = D["invs"]["period"]["view"]["issues"] + D["invs"]["calendar"]["view"]["issues"]
    issues = [i for i in issues if "Notas de crédito 文件" not in i]
    if issues:
        ws["A3"] = "⚠ " + "；".join(sorted(set(issues)))
        ws["A3"].font = F(10, True, K.C_RED)
    RF.set_widths(ws, [2, 26, 120])
    P, C = D["invs"]["period"], D["invs"]["calendar"]
    S, m = D["S"], D["S"]["m"]
    pt, ct = _periods_text(D)
    money = lambda v: "{:,.{d}f}".format(v, d=D["dec"])
    docs = []
    for d in D["docs"]:
        n = int((S["billing_all"]["_file"] == d["name"]).sum())
        docs.append(("", d["name"], "%s · %s · %d 行" % (
            "账单 Facturación" if d["doc"] == "FAC" else "贷记单 Notas de crédito",
            "还没关账（文件顶部写着 podría tener cambios）" if d["open"] else "已关账", n)))
    status = (P["view"]["issues"] + C["view"]["issues"]) or ["两种读法用到的账单都已下载、都已关账。"]
    ship_note = "（墨西哥只有少数订单买家另付运费）" if D["site"] == "mx" else ""
    cannot = [("", lab, "%s —— %s" % (
        "参考值 %s（取自现行订单口径用的文件，不计入账单口径）" % money(v) if v is not None else "两边都没有",
        why)) for lab, _es, v, why in D["unavailable"]]
    methods = P["n"]["rev_methods"]
    rows = [
        ("S", "一、这份报表是什么", ""),
        ("", "目的", "实验：如果把损益的依据从销售报表（Ventas）换成当月的账单（Facturación + Notas de "
                    "crédito），能做到多好、和现行报表差多少、哪个口径更稳定。现行月度报表一个数都不改。"),
        ("", "现行订单口径", "收入、退款、运费收入、优惠券取 Ventas，按下单日归月；佣金和退货处理费按订单号到"
                         "所有已下载的账单里去找（不论哪张账单计费）；配送费取 Ventas；广告/仓储等月度费用按"
                         "计费日归自然月。Ventas 是平台的实时导出，订单状态、退款每天都在变。"),
        ("", "账单口径（本实验）", "收入、冲销、费用全部取账单里的行：每笔销售在账单上有一行“Cargo por venta”，"
                              "带着售价、件数、单价、下单日和商品；取消/退款在账单上是“Anulación del cargo"
                              " por venta”（冲销），费用按科目逐行列出。账单关账后就是定稿的凭证。"),
        ("S", "二、期间（重要）", ""),
        ("", "账单期", pt),
        ("", "自然月切", ct + "这一列给出“按自然月看”的读法，方便和现行报表对比。"),
        ("", "现行订单口径", "下单日 %s 至 %s（Ventas 导出于 %s）。"
         % (_d(D["S"]["p0"]), _d(D["S"]["p1"]), _d(S["meta"].get("downloaded_at", "")[:10]) or "?")),
        ("", "账单状态", "；".join(status)),
        ("S", "三、用到的文件", ""),
        ("", "下载目录", D["folder"]),
    ] + docs + [
        ("", "Ventas（现行口径用）", "%s%s" % (os.path.basename(S["files"].get("ventas") or ""),
                                         "；⚠ 平台提示导出可能缺数据：%s" % S["ventas_warning"]
                                         if S.get("ventas_warning") else "")),
        ("S", "四、账单能给什么", ""),
        ("", "开票销售", "每笔被计佣金的销售：售价（Total de la venta）、件数、单价、下单日、商品、包裹号。"),
        ("", "取消与退款", "冲销行（Anulación del cargo por venta，在账单或贷记单里）。冲掉的销售额按比例估算，见五。"),
        ("", "平台费用", "全部科目：佣金、配送费、退货处理费、广告、Full 仓储/揽收/罚金、页面维护费……以及它们的冲销。"
                      "科目名称和中文与现行报表一致（FEE_CATALOG）；字典里没有的科目单独列出、标 ⚠，绝不并进别的行。"),
        ("", "买家支付的运费", "账单的“Envío a cargo del cliente”列。实测与 Ventas 的“Ingresos por envío”逐单相等"
                          "（DAB_CL 592/592、BOCINA_TA02 13/13）%s。" % ship_note),
        ("S", "五、账单不能给什么（发票口径无法得到）", ""),
    ] + cannot + [
        ("S", "六、方法", ""),
        ("", "期间", "账单期 = 名为本月的账单文件（Facturación + 同名的 Notas de crédito）的全部行；"
                   "自然月切 = 所有已下载账单里计费日（Fecha del cargo）在本月的行。"),
        ("", "销售", "每个 Número de venta 算一笔（以 Cargo por venta 行为准；联盟佣金行只加佣金，不重复计销售）。"
                   "多件包裹在账单上每件一笔，和 Ventas 对照时按 Número de paquete 并回包裹。"),
        ("", "冲销销售额", "冲销比例 = 冲销的佣金 ÷ 这笔销售的原佣金；估算冲销销售额 = 比例 × 售价。"
                        "依据：平台按退款比例同时退还佣金（现行报表的贷记单核对已用 MercadoPago 现金逐单证实）。"
                        "原佣金的找法依次是：%s。本期 %d 笔冲销：%s。"
         % ("；".join("（%d）%s" % (i + 1, x) for i, x in enumerate(LINK_ORDER)), P["n"]["n_rev"],
            "，".join("%s %d 笔" % (x.split("（")[0].split("：")[0], methods[x])
                     for x in LINK_ORDER if methods.get(x)) or "无")),
        ("", "运费收入", "买家支付的运费取 Cargo por envíos 行上的“Envío a cargo del cliente”；配送费被冲销时，"
                      "按同样比例算作退回买家。"),
        ("", "费用", "账单净额 = 计费 − 冲销（含贷记单）。全部账单行金额之和 = 费用合计，② 页有核对行（残差应为 0）。"),
        ("", "经营贡献毛利", "账单口径 = 净开票销售额 + 买家运费（净）− 平台费用合计；没有扣代扣税（现行报表也不扣，"
                          "它是可抵税款），没有扣卖家优惠券（账单里没有），没有商品成本。"),
        ("S", "七、本期要点", ""),
        ("", "商品销售额", "现行 %s；账单期 %s；自然月切 %s。"
         % (money(m["ing"]), money(P["n"]["gross_sales"]), money(C["n"]["gross_sales"]))),
        ("", "经营贡献毛利", "现行 %s；账单期 %s；自然月切 %s。"
         % (money(S["contrib"]), money(P["n"]["contribution"]), money(C["n"]["contribution"]))),
        ("", "平台费用合计", "现行 %s；账单期 %s；自然月切 %s。"
         % (money(_cur_fees(D)), money(-P["n"]["fee_total"]), money(-C["n"]["fee_total"]))),
        ("", "差异怎么来的", "逐行原因见 ③；逐单名单见 ⑤。"),
    ]
    if A and A.get("pairs"):
        for p in A["pairs"]:
            rows.append(("", "稳定性 %s → %s" % (snapshot_label(p["a"]["folder"])[:10],
                                              snapshot_label(p["b"]["folder"])[:10]),
                         "订单口径：%s 账单口径：%s %s" % (" ".join(p["why"]["order"][:2]),
                                                   p["why"]["period"], p["why"]["calendar"])))
    rows += [
        ("S", "八、各页", ""),
        ("", SHEETS[1], "账单口径的损益表：账单期与自然月切两列；账单得不到的数单列备查，不计入合计。"),
        ("", SHEETS[2], "与现行订单口径逐行对比：差额、原因类别、原因说明；下面是订单 ↔ 开票销售的逐单对照摘要。"),
        ("", SHEETS[3], "费用按科目：计费、冲销、净额；未登记科目单列；结算通道。"),
        ("", SHEETS[4], "冲销明细（关联方法）、只在一边的订单/销售、金额不一致的订单、开票销售明细。"),
        ("", SHEETS[5], "同一店同一月多次下载时，两种口径各变了多少、为什么。"),
    ]
    RF.readme_rows(ws, 5, rows)
    return ws


def _cur_fees(D):
    m = D["S"]["m"]
    return -(m["com_gross"] + m["env_cost"] + m["dev_net"] + m["unc_order"]
             + sum(m["bill_fee_net"].values()) + m["unc_bill"])


# ════════════════════════════════════════════════════════════════════════
# ② 账单口径损益
# ════════════════════════════════════════════════════════════════════════

def pl_sheet(wb, D):
    ws = wb.create_sheet(SHEETS[1])
    money = D["money"]
    P, C = D["invs"]["period"], D["invs"]["calendar"]
    pt, ct = _periods_text(D)
    K.title_rows(ws, "账单口径损益表（实验）· %s · %s" % (D["store"], D["month"]),
                 "账单期 = " + pt, "自然月切 = " + ct)
    RF.set_widths(ws, [34, 44, 18, 18, 80])
    r = _head(ws, 5, ["项目", "原文（平台科目）", "账单期\n%s" % _span(P["view"]),
                      "自然月切\n%s" % _span(C["view"]), "说明"])
    at = {}

    def row(k, label, es, kind, vp, vc, note, style=None):
        nonlocal r
        fillc = {"sum": K.C_SUB, "f": K.C_CALC, "tot": K.C_TOTAL}.get(style)
        bold = style is not None
        _put(ws, r, 1, label, bold=bold, fillc=fillc)
        _put(ws, r, 2, es, color=K.C_GREY, fillc=fillc)
        for c, v in ((3, vp), (4, vc)):
            if isinstance(v, str) and v.startswith("="):
                _put(ws, r, c, v, _fmt_of(kind, money), bold=bold, fillc=fillc)
            else:
                _value(ws, r, c, v, money, kind, bold, fillc)
        _put(ws, r, 5, note, color=K.C_GREY, wrap_text=True, fillc=fillc)
        if note and len(note) > 70:
            ws.row_dimensions[r].height = K.est_height(note, 62)
        at[k] = r
        r += 1

    def f(*keys, col):
        return "=" + "+".join("%s%d" % (col, at[k]) for k in keys)

    r = _band(ws, r, "收入（账单的 Cargo por venta / Cargo por envíos 行）", 5)
    n = lambda inv, k: inv["n"][k]
    row("n_sales", "开票销售笔数", "Número de venta", "n", n(P, "n_sales"), n(C, "n_sales"),
        "一件商品一笔；按计费日落在期间内（下单月份分布见下）")
    row("units", "开票件数", "Cantidad vendida", "n", n(P, "units"), n(C, "units"), "")
    row("gross", "开票销售额", "Total de la venta（Cargo por venta 行）", "m", n(P, "gross_sales"),
        n(C, "gross_sales"), "买家实付商品金额（含税售价）；每笔销售算一次")
    row("rev", "减：冲销销售额（估算）", "Anulación del cargo por venta", "m", -n(P, "rev_sales"),
        -n(C, "rev_sales"),
        "冲销佣金 ÷ 原佣金 × 售价。账单期 %d 笔（全额 %d 笔），自然月切 %d 笔；逐笔见 ⑤"
        % (n(P, "n_rev"), n(P, "rev_full"), n(C, "n_rev")))
    row("net_sales", "净开票销售额", "", "m", f("gross", "rev", col="C"), f("gross", "rev", col="D"),
        "＝ 开票销售额 − 冲销销售额", "f")
    row("ship", "加：买家支付的运费", "Envío a cargo del cliente（Cargo por envíos 行）", "m",
        n(P, "ship_income"), n(C, "ship_income"),
        "实测与 Ventas 的 Ingresos por envío 逐单相等；账单期 %d 笔" % n(P, "ship_n"))
    row("ship_rev", "减：运费退回买家（随配送费冲销）", "Anulación del cargo por envíos", "m",
        -n(P, "ship_income_rev"), -n(C, "ship_income_rev"), "配送费被冲销时，按同样比例算作退回买家的运费")
    row("income", "净收入合计", "", "m", f("net_sales", "ship", "ship_rev", col="C"),
        f("net_sales", "ship", "ship_rev", col="D"), "", "sum")
    r = _band(ws, r, "平台费用（账单净额 = 计费 − 冲销，含贷记单；负数 = 费用）", 5)
    first = None
    fp = dict((x["es"], x) for x in P["fees"])
    fc = dict((x["es"], x) for x in C["fees"])
    for es, zh, bucket, order_level in FEE_CATALOG:
        a, b = fp[es], fc[es]
        row("fee:" + es, "减：" + zh, es, "m", -a["net"], -b["net"],
            "%s；账单期计费 %d 笔 %s，冲销 %d 笔 %s" % (
                "订单级" if order_level else "账单级", a["n_gross"],
                "{:,.{d}f}".format(a["gross"], d=D["dec"]), a["n_rev"],
                "{:,.{d}f}".format(a["rev"], d=D["dec"])))
        first = first or at["fee:" + es]
    unk = {}
    for v, inv in (("p", P), ("c", C)):
        for u in inv["unknown"]:
            unk.setdefault(u["es"], {"p": 0.0, "c": 0.0, "n": 0})[v] += u["amount"]
    if unk:
        for es, u in sorted(unk.items()):
            row("unk:" + es, "减：⚠ 未登记科目", es, "m", -u["p"], -u["c"],
                "FEE_CATALOG 里没有这个 Detalle：金额照常计入，科目要人工确认")
            for c in range(1, 6):
                ws.cell(row=at["unk:" + es], column=c).fill = fill(K.C_BAD)
    else:
        row("unk", "减：未登记科目 ⚠", "(未在 FEE_CATALOG 登记)", "m", 0.0, 0.0, "本期没有")
    last = r - 1
    row("fees", "平台费用合计", "", "m", "=SUM(C%d:C%d)" % (first, last),
        "=SUM(D%d:D%d)" % (first, last), "", "sum")
    row("contrib", "经营贡献毛利（账单口径，未扣商品成本）", "", "m", f("income", "fees", col="C"),
        f("income", "fees", col="D"), "＝ 净收入合计 + 平台费用合计；未扣代扣税、卖家优惠券、商品成本（见下）", "tot")
    r += 1
    r = _band(ws, r, "核对", 5)
    row("chk", "账单行金额合计（取负）", "Σ Valor del cargo", "m", -n(P, "fee_total"), -n(C, "fee_total"),
        "期间内所有账单行（Facturación + Notas de crédito）的金额之和")
    row("resid", "残差（应为 0）", "", "m", "=C%d-C%d" % (at["fees"], at["chk"]),
        "=D%d-D%d" % (at["fees"], at["chk"]), "上面逐科目加总 − 全部账单行；不为 0 说明有行没归到任何科目", "tot")
    r += 1
    r = _band(ws, r, "发票口径无法得到（现行订单口径的值，仅供参考，不计入上面任何合计）", 5)
    r = _head(ws, r, ["项目", "原文", "现行订单口径（参考）", "", "为什么账单里没有"])
    for lab, es, v, why in D["unavailable"]:
        _put(ws, r, 1, lab, color=K.C_GREY)
        _put(ws, r, 2, es, color=K.C_GREY)
        if v is None:
            _put(ws, r, 3, "—", color=K.C_GREY)
        else:
            _put(ws, r, 3, v, money, color=K.C_GREY, italic=True)
        _put(ws, r, 4, "")
        _put(ws, r, 5, why, color=K.C_GREY, wrap_text=True)
        ws.row_dimensions[r].height = K.est_height(why, 62)
        r += 1
    r += 1
    r = _band(ws, r, "结算通道（Descontado de la operación）", 5)
    r = _head(ws, r, ["通道", "含义", "账单期", "自然月切", ""])
    for ch in ("Si", "No", "No aplica"):
        _put(ws, r, 1, "%s（%s）" % (CHANNEL_ZH[ch], ch))
        _put(ws, r, 2, CHANNEL_MEANING[ch], color=K.C_GREY)
        _put(ws, r, 3, -P["n"]["channel"][ch], money)
        _put(ws, r, 4, -C["n"]["channel"][ch], money)
        _put(ws, r, 5, "")
        r += 1
    r += 1
    r = _band(ws, r, "开票销售的下单月份（看出有多少是别的月份的订单）", 5)
    r = _head(ws, r, ["下单月份", "", "账单期 笔数", "自然月切 笔数", ""])
    months = sorted(set(P["n"]["sale_months"]) | set(C["n"]["sale_months"]))
    for mo in months:
        _put(ws, r, 1, mo)
        _put(ws, r, 2, "")
        _put(ws, r, 3, P["n"]["sale_months"].get(mo, 0), K.INT)
        _put(ws, r, 4, C["n"]["sale_months"].get(mo, 0), K.INT)
        _put(ws, r, 5, "")
        r += 1
    ws.freeze_panes = "C6"
    return ws


# ════════════════════════════════════════════════════════════════════════
# ③ 与现行订单口径对比
# ════════════════════════════════════════════════════════════════════════

def compare_sheet(wb, D):
    ws = wb.create_sheet(SHEETS[2])
    money = D["money"]
    P, C = D["invs"]["period"], D["invs"]["calendar"]
    K.title_rows(ws, "与现行订单口径逐行对比 · %s · %s" % (D["store"], D["month"]),
                 "现行 = 月度报表 ② 合并损益表的数（同一个下载目录，用它自己的 derive() 算出，未做任何改动）。"
                 "差额 = 账单口径 − 现行。",
                 "负数 = 费用/减项。“无法得到”= 账单里没有这一项。")
    RF.set_widths(ws, [32, 34, 17, 17, 17, 16, 16, 26, 62, 62])
    r = _head(ws, 5, ["项目", "原文（Ventas ↔ 账单）", "现行订单口径\n下单日 %s" % D["month"],
                      "账单口径·账单期\n%s" % _span(P["view"]), "账单口径·自然月切\n%s" % _span(C["view"]),
                      "差额\n账单期 − 现行", "差额\n自然月切 − 现行", "原因类别", "说明（账单期）",
                      "说明（自然月切）"], height=44)
    at, block = {}, []
    for ln in D["lines"]:
        k, kind = ln["key"], ln["kind"]
        style = {"sub": K.C_SUB, "tot": K.C_TOTAL}.get(kind)
        bold = style is not None
        _put(ws, r, 1, ln["label"], bold=bold, fillc=style)
        _put(ws, r, 2, ln["es"], color=K.C_GREY, fillc=style)
        if kind == "sub":
            rng = [x for x in block]
            for col in "CDE":
                _put(ws, r, " CDE".index(col) + 2, "=SUM(%s)" % ",".join("%s%d" % (col, x) for x in rng)
                     if rng else 0, money, bold=True, fillc=style)
            block = []
        elif kind == "tot":
            subs = [at["_net_sales"], at["_order_fees"], at["_bill_fees"]]
            for col in "CDE":
                _put(ws, r, " CDE".index(col) + 2, "=SUM(%s)" % ",".join("%s%d" % (col, x) for x in subs),
                     money, bold=True, fillc=style)
        else:
            _value(ws, r, 3, ln["cur"], money, "n" if kind == "n" else "m")
            _value(ws, r, 4, ln["inv"].get("period"), money, "n" if kind == "n" else "m")
            _value(ws, r, 5, ln["inv"].get("calendar"), money, "n" if kind == "n" else "m")
            if kind == "m":
                block.append(r)
        fm = K.INT if kind == "n" else money
        _put(ws, r, 6, _diff_formula("C", "D", r), fm, bold=bold, fillc=style)
        _put(ws, r, 7, _diff_formula("C", "E", r), fm, bold=bold, fillc=style)
        cats = ln["cats"]
        if isinstance(cats, dict):
            cats = list(dict.fromkeys(cats.get("period", []) + cats.get("calendar", [])))
        _put(ws, r, 8, "、".join(CAT[c] for c in cats), wrap_text=True, fillc=style)
        npd, ncl = ln["notes"].get("period", ""), ln["notes"].get("calendar", "")
        _put(ws, r, 9, npd, color=K.C_GREY, wrap_text=True, fillc=style)
        _put(ws, r, 10, "同左" if ncl == npd and ncl else ncl, color=K.C_GREY, wrap_text=True,
             fillc=style)
        longest = max(len(npd or ""), len(ncl or ""))
        if longest > 60:
            ws.row_dimensions[r].height = K.est_height("x" * longest, 52)
        at[k] = r
        r += 1
    r += 1
    r = _band(ws, r, "核对", 10)
    _put(ws, r, 1, "经营贡献毛利（独立计算）")
    _put(ws, r, 2, "现行：report_build.build_context；账单：本实验", color=K.C_GREY)
    _put(ws, r, 3, D["S"]["contrib"], money)
    _put(ws, r, 4, P["n"]["contribution"], money)
    _put(ws, r, 5, C["n"]["contribution"], money)
    chk = r
    r += 1
    _put(ws, r, 1, "残差（应为 0）", bold=True, fillc=K.C_TOTAL)
    _put(ws, r, 2, "上表合计 − 独立计算", color=K.C_GREY, fillc=K.C_TOTAL)
    for c in "CDE":
        _put(ws, r, " CDE".index(c) + 2, "=%s%d-%s%d" % (c, at["_contrib"], c, chk), money,
             bold=True, fillc=K.C_TOTAL)
    r += 2

    r = _band(ws, r, "逐单对照：本月订单（Ventas，下单日）↔ 开票销售（账单）", 10)
    r = _note(ws, r, "Ventas 的多件包裹（“Paquete de N productos”母行 + N 行子行）在这里算一单；账单每件一笔，"
                     "按 Número de paquete 并回包裹。名单见 ⑤。", 10)
    r = _head(ws, r, ["项目", "", "", "账单期", "自然月切", "", "", "", "说明", ""])
    rs = dict((v, rec_summary(D["recs"][v])) for v in ("period", "calendar"))
    R = D["recs"]
    items = [
        ("本月订单（Ventas）", R["period"]["n_month_orders"], R["calendar"]["n_month_orders"], "n",
         "多件包裹算一单"),
        ("开票销售笔数", P["n"]["n_sales"], C["n"]["n_sales"], "n", "账单上一件商品一笔"),
        ("两边都有的订单", rs["period"]["matched_n"], rs["calendar"]["matched_n"], "n", ""),
        ("只在账单：上月订单本期计费", rs["period"]["extra_prev"], rs["calendar"]["extra_prev"], "n", ""),
        ("只在账单：下月订单本期计费", rs["period"]["extra_next"], rs["calendar"]["extra_next"], "n",
         "多为时区：账单的下单日按平台时间，Ventas 按本地时间"),
        ("只在账单：Ventas 里查无此单", rs["period"]["extra_absent"], rs["calendar"]["extra_absent"], "n",
         "多为下单后随即取消、账单上计费又冲销的订单；Ventas 不列这类订单"),
        ("只在 Ventas：在别的账单计费", rs["period"]["missing_elsewhere"],
         rs["calendar"]["missing_elsewhere"], "n", "智利：账单期以外（切期日之后）的订单在下一张账单"),
        ("只在 Ventas：取消，平台未计费", rs["period"]["missing_cancel"], rs["calendar"]["missing_cancel"],
         "n", ""),
        ("只在 Ventas：其他（未计费 / 金额为 0）", rs["period"]["missing_other"],
         rs["calendar"]["missing_other"], "n", "已下载的账单里都没有这单的佣金"),
        ("两边金额不一致的订单", rs["period"]["differ_n"], rs["calendar"]["differ_n"], "n", "见 ⑤"),
    ]
    for lab, a, b, kind, note in items:
        _put(ws, r, 1, lab)
        for c in (2, 3, 6, 7, 8, 10):
            _put(ws, r, c, "")
        _put(ws, r, 4, a, _fmt_of(kind, money))
        _put(ws, r, 5, b, _fmt_of(kind, money))
        _put(ws, r, 9, note, color=K.C_GREY, wrap_text=True)
        r += 1
    r += 1
    r = _head(ws, r, ["GMV 差额拆解", "", "", "账单期", "自然月切", "", "", "", "说明", ""])
    g0 = r
    parts = [
        ("开票销售额", P["n"]["gross_sales"], C["n"]["gross_sales"], "账单"),
        ("减：本月订单 GMV（Ventas，按订单汇总）", -R["period"]["month_gmv"], -R["calendar"]["month_gmv"],
         "与现行 ② 的商品销售额同一批 Ventas 行"),
        ("＝ 差额", "=D%d+D%d" % (g0, g0 + 1), "=E%d+E%d" % (g0, g0 + 1), ""),
        ("其中：只在账单的销售", rs["period"]["extra_total"], rs["calendar"]["extra_total"], "+"),
        ("其中：只在 Ventas 的订单", -rs["period"]["missing_gmv"], -rs["calendar"]["missing_gmv"], "−"),
        ("其中：两边都有、金额不同", rs["period"]["matched_diff"], rs["calendar"]["matched_diff"], ""),
        ("残差（应为 0）", "=D%d-SUM(D%d:D%d)" % (g0 + 2, g0 + 3, g0 + 5),
         "=E%d-SUM(E%d:E%d)" % (g0 + 2, g0 + 3, g0 + 5), "拆解完整时为 0"),
    ]
    for lab, a, b, note in parts:
        last = lab.startswith("残差") or lab.startswith("＝")
        fc = K.C_TOTAL if lab.startswith("残差") else (K.C_CALC if lab.startswith("＝") else None)
        _put(ws, r, 1, lab, bold=last, fillc=fc)
        for c in (2, 3, 6, 7, 8, 10):
            _put(ws, r, c, "", fillc=fc)
        _put(ws, r, 4, a, money, bold=last, fillc=fc)
        _put(ws, r, 5, b, money, bold=last, fillc=fc)
        _put(ws, r, 9, note, color=K.C_GREY, fillc=fc)
        r += 1
    ws.freeze_panes = "C6"
    return ws


# ════════════════════════════════════════════════════════════════════════
# ④ 费用明细
# ════════════════════════════════════════════════════════════════════════

def fees_sheet(wb, D):
    ws = wb.create_sheet(SHEETS[3])
    money = D["money"]
    P, C = D["invs"]["period"], D["invs"]["calendar"]
    S = D["S"]
    K.title_rows(ws, "平台费用明细（账单口径）· %s · %s" % (D["store"], D["month"]),
                 "金额为正 = 平台向卖家收费（账单原样）；冲销为负。账单期 %s，自然月切 %s。"
                 % (_span(P["view"]), _span(C["view"])),
                 "现行（参考）：订单级 = 本月订单关联到的账单行净额（现行 ② 的佣金行只取毛额、配送费取 Ventas，"
                 "见 ③）；账单级 = 现行 ② 的同名行。")
    RF.set_widths(ws, [30, 46, 12, 10, 11, 16, 11, 16, 16, 16, 16])
    r = _head(ws, 5, ["科目", "Detalle 原文", "类别", "级别", "账单期\n计费笔数", "账单期\n计费金额",
                      "账单期\n冲销笔数", "账单期\n冲销金额", "账单期\n净额", "自然月切\n净额",
                      "现行订单口径\n（参考）"], height=40)
    L = S["billing_linked"]
    fc = dict((x["es"], x) for x in C["fees"])
    first = r
    for f in P["fees"]:
        es = f["es"]
        if f["order_level"]:
            cur = float(L[L["parent_detalle"].astype(str) == es]["amount"].sum()) if len(L) else 0.0
        else:
            cur = S["m"]["bill_fee_net"].get(es, 0.0)
        vals = [f["zh"], es, BUCKET_ZH.get(f["bucket"], f["bucket"]),
                "订单级" if f["order_level"] else "账单级", f["n_gross"], f["gross"], f["n_rev"],
                f["rev"], "=F%d+H%d" % (r, r), fc[es]["net"], cur]
        kinds = "tttnmnmmmmm"
        quiet = not (f["n_gross"] or f["n_rev"] or abs(fc[es]["net"]) > 0.004 or abs(cur) > 0.004)
        for j, (v, kd) in enumerate(zip(vals, kinds), 1):
            _put(ws, r, j, v, _fmt_of(kd, money), color=K.C_GREY if quiet else "000000",
                 fillc=K.C_CALC if j == 9 else None)
        r += 1
    unk = {}
    for v, inv in (("p", P), ("c", C)):
        for u in inv["unknown"]:
            x = unk.setdefault(u["es"], {"p": 0.0, "c": 0.0, "n": 0, "lvl": u["order_level"],
                                         "ch": u["channels"]})
            x[v] += u["amount"]
            if v == "p":
                x["n"] += u["n"]
    for es, u in sorted(unk.items()):
        vals = ["⚠ 未登记科目", es, "未登记", "订单级" if u["lvl"] else "账单级", u["n"], u["p"], "", "",
                "=F%d" % r, u["c"], ""]
        for j, (v, kd) in enumerate(zip(vals, "tttnmnmmmmm"), 1):
            _put(ws, r, j, v, _fmt_of(kd, money), fillc=K.C_BAD)
        r += 1
    last = r - 1
    _put(ws, r, 1, "合计", bold=True, fillc=K.C_SUB)
    for j in range(2, 12):
        col = get_column_letter(j)
        v = "=SUM(%s%d:%s%d)" % (col, first, col, last) if j >= 5 else ""
        _put(ws, r, j, v, K.INT if j in (5, 7) else money, bold=True, fillc=K.C_SUB)
    tot = r
    r += 1
    _put(ws, r, 1, "核对：账单行金额合计", bold=False)
    for j in range(2, 12):
        _put(ws, r, j, "")
    _put(ws, r, 9, P["n"]["fee_total"], money)
    _put(ws, r, 10, C["n"]["fee_total"], money)
    chk = r
    r += 1
    _put(ws, r, 1, "残差（应为 0）", bold=True, fillc=K.C_TOTAL)
    for j in range(2, 12):
        _put(ws, r, j, "", fillc=K.C_TOTAL)
    _put(ws, r, 9, "=I%d-I%d" % (tot, chk), money, bold=True, fillc=K.C_TOTAL)
    _put(ws, r, 10, "=J%d-J%d" % (tot, chk), money, bold=True, fillc=K.C_TOTAL)
    r += 2
    if not unk:
        r = _note(ws, r, "本期没有未登记科目：账单里每一个 Detalle 都在 FEE_CATALOG 里。", 11, K.C_GREEN)
    else:
        r = _note(ws, r, "⚠ 有 %d 个 Detalle 不在 FEE_CATALOG 里（红底行）：金额已计入，科目归属请人工确认，"
                         "确认后在 platforms/mercadolibre/report/catalog.py 登记。" % len(unk), 11, K.C_RED, True)
    ws.freeze_panes = "C6"
    return ws


# ════════════════════════════════════════════════════════════════════════
# ⑤ 销售与冲销明细
# ════════════════════════════════════════════════════════════════════════

def _views_to_list(D):
    """MX/PE: the two readings are the same rows -> list them once."""
    if D["cycle_day"] == 1:
        return [("period", "账单期（= 自然月）")]
    return [("period", "账单期 %s" % _span(D["invs"]["period"]["view"])),
            ("calendar", "自然月切 %s" % _span(D["invs"]["calendar"]["view"]))]


def detail_sheet(wb, D):
    ws = wb.create_sheet(SHEETS[4])
    money = D["money"]
    K.title_rows(ws, "销售与冲销明细 · %s · %s" % (D["store"], D["month"]),
                 "每一笔都能回到账单原始行（销售号 = Número de venta，计费号 = Número del cargo）。",
                 "墨西哥/秘鲁的账单按自然月，账单期和自然月切是同一批行，只列一次。"
                 if D["cycle_day"] == 1 else "智利：账单期和自然月切分别列出。")
    RF.set_widths(ws, [20, 20, 14, 14, 14, 13, 13, 11, 14, 15, 11, 44, 14, 30, 30, 15, 40])
    orders, parent_of = D["orders"], D["parent_of"]
    r = 5
    r = _band(ws, r, "A. 冲销的关联方法（找到这笔销售的原佣金，才能算冲掉了多少）", 17)
    r = _head(ws, r, ["方法", "", "", "账单期", "自然月切"] + [""] * 12)
    for mth in LINK_ORDER:
        _put(ws, r, 1, mth, wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        _put(ws, r, 4, D["invs"]["period"]["n"]["rev_methods"].get(mth, 0), K.INT)
        _put(ws, r, 5, D["invs"]["calendar"]["n"]["rev_methods"].get(mth, 0), K.INT)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    rev_cols = [("冲销计费日", "d"), ("凭证", "t"), ("销售号", "t"), ("包裹号", "t"), ("下单日", "d"),
                ("冲销佣金", "m"), ("原佣金", "m"), ("冲销比例", "p"), ("售价", "m"),
                ("估算冲销销售额", "m"), ("估算件数", "r"), ("关联方法", "t"), ("原佣金计费日", "d"),
                ("Ventas 状态（参考）", "t"), ("商品", "t"), ("Ventas 退款净额（参考）", "m"),
                ("文件", "t")]
    for v, name in _views_to_list(D):
        R = D["invs"][v]["reversals"]
        r = _band(ws, r, "B. 冲销明细 · %s（%d 笔）" % (name, len(R)), 17)
        rows = []
        for _, x in R.sort_values("cdate").iterrows():
            o = orders.get(order_key(orders, parent_of, x["kpack"], x["ksale"])) or {}
            rows.append([x["cdate"], x["doc"], x["ksale"], x["kpack"], x["sdate"], x["rev_commission"],
                         x["orig_commission"], x["ratio"], x["sale_total"], x["est_sales"],
                         x["est_units"], x["method"].split("（")[0], x["parent_date"],
                         o.get("status", "Ventas 查无"), x["title"], o.get("refund"), x["file"]])
        r = _table(ws, r, rev_cols, rows, money) + 1
    rec_cols = [("销售号", "t"), ("Ventas 订单号", "t"), ("下单日（账单）", "d"), ("计费日", "d"),
                ("Ventas 下单日", "d"), ("件数", "n"), ("售价", "m"), ("原因", "t"), ("商品", "t")]
    miss_cols = [("Ventas 订单号", "t"), ("下单日", "d"), ("状态", "t"), ("件数", "n"), ("GMV", "m"),
                 ("退款净额", "m"), ("原因", "t"), ("SKU", "t"), ("商品", "t")]
    diff_cols = [("Ventas 订单号", "t"), ("下单日", "d"), ("状态", "t"), ("账单售价合计", "m"),
                 ("Ventas GMV", "m"), ("差额", "m"), ("账单件数", "n"), ("Ventas 件数", "n")]
    for v, name in _views_to_list(D):
        rec = D["recs"][v]
        E, M, T = rec["extra"], rec["missing"], rec["differ"]
        r = _band(ws, r, "C. 本期开票、但不是本月订单 · %s（%d 笔）" % (name, len(E)), 17)
        r = _table(ws, r, rec_cols, [[x["ksale"], x["okey"], x["sdate"], x["cdate"], x["odate"],
                                      x["units"], x["total"], x["why"], x["title"]]
                                     for _, x in E.sort_values(["code", "sdate"]).iterrows()],
                   money) + 1
        r = _band(ws, r, "D. 本月订单、但不在本期开票 · %s（%d 单）" % (name, len(M)), 17)
        r = _table(ws, r, miss_cols, [[x["okey"], x["odate"], x["status"], x["units"], x["gmv"],
                                       x["refund"], x["why"], x["sku"], x["title"]]
                                      for _, x in M.sort_values(["code", "odate"]).iterrows()],
                   money) + 1
        r = _band(ws, r, "E. 两边都有、金额不一致的订单 · %s（%d 单）" % (name, len(T)), 17)
        r = _table(ws, r, diff_cols, [[x["okey"], x["odate"], x["status"], x["inv_total"], x["gmv"],
                                       x["diff"], x["inv_units"], x["units"]]
                                      for _, x in T.iterrows()], money) + 1
    sale_cols = [("销售号", "t"), ("包裹号", "t"), ("Ventas 订单号", "t"), ("下单日", "d"),
                 ("计费日", "d"), ("件数", "n"), ("单价", "m"), ("售价", "m"), ("佣金", "m"),
                 ("Estado del cargo", "t"), ("通道", "t"), ("凭证", "t"), ("商品", "t")]
    for v, name in _views_to_list(D):
        sales = D["recs"][v]["sales"]
        r = _band(ws, r, "F. 开票销售明细 · %s（%d 笔）" % (name, len(sales)), 17)
        r = _table(ws, r, sale_cols, [[x["ksale"], x["kpack"], x["okey"], x["sdate"], x["cdate"],
                                       x["units"], x["price"], x["total"], x["commission"],
                                       x["state"], x["channel"], x["doc"], x["title"]]
                                      for _, x in sales.sort_values(["cdate", "ksale"]).iterrows()],
                   money) + 1
    ws.freeze_panes = "A5"
    return ws


# ════════════════════════════════════════════════════════════════════════
# 稳定性（⑥ of the experiment report, ② of the stability workbook)
# ════════════════════════════════════════════════════════════════════════

def stability_table(ws, r, A, money):
    """One block per basis: metrics × snapshots, change last − first, %, why."""
    snaps, kns = A["snaps"], A["kns"]
    n = len(snaps)
    labels = [snapshot_label(D["folder"]) for D in snaps]
    ncols = n + 4
    RF.set_widths(ws, [44] + [17] * n + [16, 10, 80])
    for b, title in (("order", "A. 现行订单口径（Ventas，按下单日）"),
                     ("period", "B. 账单口径 · 账单期"),
                     ("calendar", "C. 账单口径 · 自然月切")):
        r = _band(ws, r, title, ncols)
        r = _head(ws, r, ["指标"] + labels + ["变化\n（末次 − 首次）", "变化 %", "说明"], height=36)
        c0, c1 = get_column_letter(2), get_column_letter(n + 1)
        for k, lab, kind in A["rows"]:
            vals = [kn[b].get(k) for kn in kns]
            if all(v is None for v in vals):
                _put(ws, r, 1, lab, color=K.C_GREY)
                for j in range(n):
                    _put(ws, r, 2 + j, UNAVAILABLE if b != "order" else "—", color=K.C_GREY)
                for j in range(n + 2, ncols + 1):
                    _put(ws, r, j, "")
                r += 1
                continue
            moved = any(v is not None and vals[0] is not None and abs(v - vals[0]) > 0.004
                        for v in vals)
            fc = K.C_WARN if moved else None
            _put(ws, r, 1, lab, fillc=fc)
            for j, v in enumerate(vals):
                _value(ws, r, 2 + j, v, money, kind, fillc=fc)
            fmtv = K.INT if kind == "n" else money
            _put(ws, r, n + 2, _diff_formula(c0, c1, r), fmtv, fillc=fc)
            _put(ws, r, n + 3, '=IF(AND(ISNUMBER(%s%d),ISNUMBER(%s%d)),IF(%s%d=0,"",(%s%d-%s%d)/ABS(%s%d)),"")'
                 % (c0, r, c1, r, c0, r, c1, r, c0, r, c0, r), K.PCT, fillc=fc)
            _put(ws, r, n + 4, "有变化" if moved else "", color=K.C_AMBER, fillc=fc)
            r += 1
        for p in A["pairs"]:
            head = "%s → %s：" % (snapshot_label(p["a"]["folder"]), snapshot_label(p["b"]["folder"]))
            texts = p["why"]["order"] if b == "order" else [p["why"][b]]
            r = _note(ws, r, head + " ".join(texts), ncols, K.C_NAVY)
        r += 1
    return r


def stability_sheet(wb, name, A, money, store, month, reason=None):
    ws = wb.create_sheet(name)
    K.title_rows(ws, "口径稳定性（同一月多次下载）· %s · %s" % (store, month),
                 "每次下载用同一套代码分别算两种口径；黄底 = 这一行在几次下载之间变了。",
                 "订单口径变 = Ventas（实时导出）里的订单、状态、退款变了，或下月账单里新计了本月订单的费用；"
                 "账单口径变 = 账单行本身变了（关账的账单不应该变）。")
    if not A:
        ws.column_dimensions["A"].width = 120
        ws["A5"] = reason or "只有一次下载，无法比较。"
        ws["A5"].font = F(11, False, K.C_GREY)
        ws["A6"] = ("要比较请把同一月的几次下载放在同一店铺目录下（文件夹名以 _m<YYYYMM> 结尾），"
                    "或运行：python -m experiments.ml_invoice_basis.stability --store <店> --month <YYYY-MM> "
                    "--folders <目录1> <目录2> …")
        ws["A6"].font = F(10, False, K.C_GREY)
        return ws
    r = 5
    for D in A["snaps"]:
        r = _note(ws, r, "%s：%s　文件：%s" % (snapshot_label(D["folder"]), D["folder"], "、".join(
            "%s%s" % (d["name"], "（未关账）" if d["open"] else "") for d in D["docs"])),
            len(A["snaps"]) + 4)
    r += 1
    stability_table(ws, r, A, money)
    ws.freeze_panes = "B5"
    return ws


# ════════════════════════════════════════════════════════════════════════
# workbooks
# ════════════════════════════════════════════════════════════════════════

def write_report_workbook(D, A, out, reason=None):
    wb = Workbook()
    wb.remove(wb.active)
    readme_sheet(wb, D, A)
    pl_sheet(wb, D)
    compare_sheet(wb, D)
    fees_sheet(wb, D)
    detail_sheet(wb, D)
    stability_sheet(wb, SHEETS[5], A, D["money"], D["store"], D["month"], reason)
    return _save(wb, out)


def write_stability_workbook(A, out):
    D0 = A["snaps"][-1]
    money = D0["money"]
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(STAB_SHEETS[0])
    K.title_rows(ws, "美客多 口径稳定性实验 · %s · %s" % (A["store"], A["month"]),
                 "实验报表，不是月度财务报表。问题：Ventas（订单表）一直在变，那换成账单口径会不会稳定？")
    RF.set_widths(ws, [2, 26, 120])
    rows = [("S", "一、比较了哪几次下载", "")]
    for D in A["snaps"]:
        rows.append(("", snapshot_label(D["folder"]), "%s；账单：%s；Ventas：%s%s" % (
            D["folder"], "、".join("%s%s" % (d["name"], "（未关账）" if d["open"] else "")
                                   for d in D["docs"]),
            os.path.basename(D["S"]["files"].get("ventas") or ""),
            "（⚠ 平台提示导出可能缺数据）" if D["S"].get("ventas_warning") else "")))
    rows += [("S", "二、怎么比", ""),
             ("", "两种口径", "每次下载都用同一套代码算：现行订单口径 = 月度报表自己的 derive()（一行未改）；"
                         "账单口径 = 本实验（账单期、自然月切两种读法）。"),
             ("", "逐单 / 逐行", "订单口径按订单号比对两次 Ventas（新增、消失、状态/金额变化），见 %s；"
                            "账单口径按（凭证, Número del cargo）比对两次账单行（新增、消失、金额/Estado/通道/发票号变化），"
                            "见 %s。" % (STAB_SHEETS[2], STAB_SHEETS[3])),
             ("S", "三、结论", "")]
    for p in A["pairs"]:
        tag = "%s → %s" % (snapshot_label(p["a"]["folder"])[:10], snapshot_label(p["b"]["folder"])[:10])
        rows.append(("", "订单口径 " + tag, " ".join(p["why"]["order"])))
        rows.append(("", "账单期 " + tag, p["why"]["period"]))
        rows.append(("", "自然月切 " + tag, p["why"]["calendar"]))
    RF.readme_rows(ws, 4, rows)

    stability_sheet(wb, STAB_SHEETS[1], A, money, A["store"], A["month"])

    ws = wb.create_sheet(STAB_SHEETS[2])
    K.title_rows(ws, "订单变化明细（Ventas，本月下单的订单）", "每对相邻的两次下载之间：新增、消失、状态或金额变化的订单。")
    RF.set_widths(ws, [20, 12, 26, 34, 34, 14, 14, 14, 14, 14, 14, 9, 9, 14, 14])
    r = 4
    cols = [("订单号", "t"), ("下单日", "d"), ("变化", "t"), ("状态（前）", "t"), ("状态（后）", "t"),
            ("GMV 前", "m"), ("GMV 后", "m"), ("退款 前", "m"), ("退款 后", "m"), ("Total 前", "m"),
            ("Total 后", "m"), ("件数 前", "n"), ("件数 后", "n"), ("佣金及税金 前", "m"),
            ("佣金及税金 后", "m")]
    for p in A["pairs"]:
        r = _band(ws, r, "%s → %s（%d 单）" % (snapshot_label(p["a"]["folder"]),
                                              snapshot_label(p["b"]["folder"]), len(p["ventas"])), 15)
        rows = []
        for x in p["ventas"]:
            b, a = x["before"] or {}, x["after"] or {}
            rows.append([x["okey"], x["odate"], x["kind"], b.get("status"), a.get("status"),
                         b.get("gmv"), a.get("gmv"), b.get("refund"), a.get("refund"), b.get("total"),
                         a.get("total"), b.get("units"), a.get("units"), b.get("com_col"),
                         a.get("com_col")])
        if not rows:
            r = _note(ws, r, "两次 Ventas 里本月的订单完全相同（订单、状态、金额都没变）。", 15, K.C_GREEN)
        else:
            r = _table(ws, r, cols, rows, money)
        r += 1
    ws.freeze_panes = "A4"

    ws = wb.create_sheet(STAB_SHEETS[3])
    K.title_rows(ws, "账单行变化明细", "按（凭证, Número del cargo）比对相邻两次下载的账单行。关账的账单应当一行不变。")
    RF.set_widths(ws, [12, 8, 16, 12, 40, 16, 14, 14, 26, 26, 12, 12, 16, 16, 40])
    r = 4
    cols = [("读法", "t"), ("凭证", "t"), ("计费号", "t"), ("计费日", "d"), ("Detalle", "t"),
            ("变化", "t"), ("金额 前", "m"), ("金额 后", "m"), ("Estado 前", "t"), ("Estado 后", "t"),
            ("通道 前", "t"), ("通道 后", "t"), ("发票号 前", "t"), ("发票号 后", "t"), ("文件", "t")]
    for p in A["pairs"]:
        tag = "%s → %s" % (snapshot_label(p["a"]["folder"]), snapshot_label(p["b"]["folder"]))
        for v, name in (("period", "账单期"), ("calendar", "自然月切")):
            diff, na, nb = p["invoice"][v]
            r = _band(ws, r, "%s · %s：%d → %d 行，%d 行不同" % (tag, name, na, nb, len(diff)), 15)
            if not diff:
                r = _note(ws, r, p["why"][v], 15, K.C_GREEN) + 1
                continue
            rows = []
            for x in diff:
                b, a = x["before"] or {}, x["after"] or {}
                z = a or b
                rows.append([name, x["doc"], x["charge"], z.get("cdate"), z.get("det"), x["kind"],
                             b.get("amount"), a.get("amount"), b.get("state"), a.get("state"),
                             b.get("channel"), a.get("channel"), b.get("invoice"), a.get("invoice"),
                             z.get("file")])
            r = _table(ws, r, cols, rows, money) + 1
        lr = p["linked"]
        r = _note(ws, r, "另：现行订单口径关联到本月订单的账单行（跨全部已下载账单）新增 %d 行、消失 %d 行。"
                         % (lr["new_n"], lr["gone_n"]), 15) + 1
    ws.freeze_panes = "A4"
    return _save(wb, out)


def _save(wb, out):
    d = os.path.dirname(os.path.abspath(out))
    if d and not os.path.isdir(d):
        os.makedirs(d)
    wb.save(out)
    return out


def finish(out, no_recalc=False):
    """Recalculate with LibreOffice when available and scan for formula errors.
    Returns the exit code: 0, or 1 when a formula error is in the file."""
    if no_recalc:
        return 0
    ok, msg = K.recalc(out)
    print("[重算] %s" % msg)
    if ok:
        nf, ne, nn, bad = K.scan_formula_errors(out)
        print("[公式] 共 %d 条，错误 %d 条，未求值 %d 条" % (nf, ne, nn))
        for b in bad:
            print("       %s" % b)
        if ne:
            print("[警告] 文件里有公式错误。")
            return 1
    return 0
