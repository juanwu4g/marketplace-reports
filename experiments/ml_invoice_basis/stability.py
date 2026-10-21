# -*- coding: utf-8 -*-
"""How much each basis moves between several downloads of the same store-month.

    python -m experiments.ml_invoice_basis.stability --store <name> --month YYYY-MM
        [--folders DIR DIR ...] [--base DIR] [--out-dir DIR] [--no-recalc]

Default folders: every `*_m<YYYYMM>` folder under downloads/<store>/, oldest
first. For each snapshot both bases are computed (order basis = the monthly
report's own derive(); invoice basis = this experiment), then consecutive
snapshots are diffed order by order (Ventas) and charge by charge (invoice rows),
so every movement of a number can be traced to the rows that moved it.

Writes `<store>_<YYYY-MM>_口径稳定性.xlsx` (never the monthly report's name).
"""

import argparse
import os
import sys

import pandas as pd

from core import console  # noqa: F401  中文输出编码保护
from platforms.mercadolibre.report.catalog import FEE_BY_ES, fmt

from experiments.ml_invoice_basis import report
from experiments.ml_invoice_basis.common import (STABILITY_SUFFIX, ExperimentError, month_folders,
    out_path, parse_month, snapshot_label, store_root)
from experiments.ml_invoice_basis.compare import month_orders

BASES = (("order", "现行订单口径（Ventas，按下单日；费用按订单号跨账单关联）"),
         ("period", "账单口径 · 账单期（当月账单文件）"),
         ("calendar", "账单口径 · 自然月切（计费日在本月）"))

ORDER_FIELDS = (("status", "状态"), ("gmv", "GMV"), ("refund", "退款"), ("total", "Total"),
                ("units", "件数"), ("com_col", "佣金及税金列"), ("ship_cost", "配送费"),
                ("ship_income", "运费收入"))
ROW_FIELDS = (("amount", "金额"), ("state", "Estado del cargo"), ("channel", "通道"),
              ("invoice", "发票号"))


def metric_rows(snaps):
    """(key, label, kind) of the stability table; bill-level fees only when any
    snapshot has them on any basis (the full list is in the experiment report)."""
    rows = [("count", "订单 / 开票销售笔数", "n"), ("units", "件数", "n"),
            ("gmv", "商品销售额 (GMV)", "m"), ("refund", "取消与退款（订单口径）/ 冲销销售额（账单口径）", "m"),
            ("ship_in", "运费收入（买家支付）", "m"), ("coupon", "卖家承担优惠券", "m"),
            ("commission", "平台销售佣金（订单口径毛额；账单口径净额）", "m"),
            ("shipping", "平台配送费（订单口径 Ventas；账单口径净额）", "m"),
            ("returns", "退货处理费（净）", "m")]
    fees = []
    for es in FEE_BY_ES:
        k = "fee:" + es
        if any(abs(kn[b].get(k) or 0.0) > 0.004 for kn in snaps for b in ("order", "period",
                                                                          "calendar")):
            fees.append((k, "减：" + FEE_BY_ES[es][0], "m"))
    rows += fees
    rows += [("unc", "未登记科目", "m"), ("fees", "平台费用合计", "m"),
             ("tax", "平台代扣税（墨西哥，Ventas − 账单）", "m"),
             ("contrib", "经营贡献毛利（未扣商品成本；账单口径冲销按估算）", "m")]
    rows += [(k, lab, "m") for k, lab in ADDED]
    return rows


# The accountant's lines (actual refunds, coupons, accrual, net of VAT): on the
# invoice basis from MercadoPago / later invoices; the order basis has only the
# net-of-VAT pair.
ADDED = (("refund_actual", "取消与退款（实际，MercadoPago；没有流水的沿用估算）"),
         ("refund_err", "冲销估算的误差（估算 − 实际，仅有实际数的笔）"),
         ("coupon_mp", "卖家承担优惠券（MercadoPago 实际）"),
         ("accrual", "应计费用（下期账单已出现的本期费用）"),
         ("contrib_adj", "经营贡献毛利（调整后，含税）"),
         ("net_sales_exvat", "净收入（不含税）"),
         ("contrib_exvat", "经营贡献毛利（不含税；账单口径为调整后）"))


SHORT = {"refund_actual": "实际退款", "refund_err": "估算误差", "coupon_mp": "优惠券", "accrual": "应计费用",
         "contrib_adj": "调整后毛利（含税）", "net_sales_exvat": "不含税净收入",
         "contrib_exvat": "不含税毛利"}


def _f(v, dec):
    return "无法得到" if v is None else fmt(round(v, dec) + 0.0, dec)     # no "-0.00"


def mp_sentence(k1, k2, view, dec, src1, src2):
    """Why the accountant's lines of one invoice reading moved between two snapshots."""
    a, b = k1[view], k2[view]
    moved = []
    for k, lab in ADDED:
        x, y = a.get(k), b.get(k)
        if x is None and y is None:
            continue
        if x is None or y is None:
            moved.append("%s：%s → %s" % (SHORT[k], _f(x, dec), _f(y, dec)))
        elif abs(y - x) > 0.004:
            moved.append("%s %s" % (SHORT[k], _f(y - x, dec)))
    head = "实际数与不含税数："
    if src1 != src2:
        head += "MercadoPago 来源不同（前一次 %s，后一次 %s），实际退款/优惠券的取数方法随之不同；" % (
            SRC_ZH.get(src1, "没有文件"), SRC_ZH.get(src2, "没有文件"))
    if not moved:
        return head + "都没有变。"
    return head + "变了的有 " + "；".join(moved) + "。"


SRC_ZH = {"settlement": "结算流水 settlement_v2", "collection": "Cobros（没有 settlement_v2）"}


# ════════════════════════════════════════════════════════════════════════
# Diffs between two snapshots
# ════════════════════════════════════════════════════════════════════════

def ventas_diff(D1, D2):
    """The month's orders (order date) in two Ventas downloads, order by order."""
    a, b = month_orders(D1["orders"], D1["month"]), month_orders(D2["orders"], D2["month"])
    out = []
    for k in list(a) + [k for k in b if k not in a]:
        x, y = a.get(k), b.get(k)
        if x is None or y is None:
            r = x or y
            out.append({"okey": k, "odate": r["date"], "kind": "新增" if x is None else "消失",
                        "before": x, "after": y})
            continue
        ch = [f for f, _l in ORDER_FIELDS
              if (x[f] != y[f] if f == "status" else abs(x[f] - y[f]) > 0.004)]
        if ch:
            out.append({"okey": k, "odate": x["date"], "kind": "变化：" + "、".join(
                lab for f, lab in ORDER_FIELDS if f in ch), "before": x, "after": y,
                "fields": ch})
    return out


def _row_table(rows):
    t = {}
    for i, r in rows.iterrows():
        t[(r["_doc"], r["kcharge"] or "row:%s" % i)] = {
            "amount": float(r["amount"]), "state": str(r.get("Estado del cargo") or "")
            if pd.notna(r.get("Estado del cargo")) else "",
            "channel": str(r["channel"]) if pd.notna(r["channel"]) else "",
            "invoice": str(r.get("N° de factura fiscal") or "")
            if pd.notna(r.get("N° de factura fiscal")) else "",
            "cdate": r["cdate"], "det": r["det"], "file": r["_file"], "ksale": r["ksale"]}
    return t


def invoice_diff(D1, D2, view):
    """The rows of one invoice reading in two downloads, charge by charge."""
    a = _row_table(D1["invs"][view]["view"]["rows"])
    b = _row_table(D2["invs"][view]["view"]["rows"])
    out = []
    for k in list(a) + [k for k in b if k not in a]:
        x, y = a.get(k), b.get(k)
        if x is None or y is None:
            out.append({"doc": k[0], "charge": k[1], "kind": "新增行" if x is None else "消失行",
                        "before": x, "after": y})
            continue
        ch = [f for f, _l in ROW_FIELDS
              if (abs(x[f] - y[f]) > 0.004 if f == "amount" else x[f] != y[f])]
        if ch:
            out.append({"doc": k[0], "charge": k[1], "kind": "变化：" + "、".join(
                lab for f, lab in ROW_FIELDS if f in ch), "before": x, "after": y})
    return out, len(a), len(b)


def linked_rows_diff(D1, D2):
    """Order basis fees come from every invoice row linked to the month's orders
    (also rows of a later, still open invoice): what was added between snapshots."""
    def tab(D):
        L = D["S"]["billing_linked"]
        return dict(((d, str(c)), float(a)) for d, c, a in zip(L["_doc"], L["charge_no"],
                                                               L["amount"]))
    a, b = tab(D1), tab(D2)
    new = [k for k in b if k not in a]
    gone = [k for k in a if k not in b]
    return {"new_n": len(new), "new_amount": sum(b[k] for k in new),
            "gone_n": len(gone), "gone_amount": sum(a[k] for k in gone)}


def doc_status(D):
    return ["%s%s" % (d["name"], "（未关账）" if d["open"] else "") for d in D["docs"]]


def explain(D1, D2, vd, ids, lrd):
    """Plain-Chinese sentences: why each basis moved (or did not) between two snapshots."""
    m1, m2 = D1["S"]["m"], D2["S"]["m"]
    dec = D1["dec"]
    kinds = {}
    for x in vd:
        kinds.setdefault(x["kind"].split("：")[0], []).append(x)
    st = [x for x in vd if "fields" in x and "status" in x["fields"]]
    ref = [x for x in vd if "fields" in x and "refund" in x["fields"]]
    order = []
    order.append("本月订单 %d → %d 单（Ventas 行 %d → %d）：新增 %d 单、消失 %d 单、内容变化 %d 单。"
                 % (len(month_orders(D1["orders"], D1["month"])),
                    len(month_orders(D2["orders"], D2["month"])), m1["orders_unique"],
                    m2["orders_unique"], len(kinds.get("新增", [])), len(kinds.get("消失", [])),
                    len(kinds.get("变化", []))))
    for kind, D in (("消失", D2), ("新增", D1)):
        xs = kinds.get(kind, [])
        if not xs:
            continue
        ds = [x["odate"] for x in xs if pd.notna(x["odate"])]
        g = sum((x["before"] or x["after"])["gmv"] for x in xs)
        other = sum(1 for x in xs if x["okey"] in D["orders"])
        order.append("%s的 %d 单下单日 %s 至 %s，GMV 合计 %s；其中 %d 单在%s的 Ventas 里还在、只是不再落在本月，"
                     "其余 %d 单整单不在那次的导出里%s。"
                     % (kind, len(xs), min(ds).strftime("%m/%d") if ds else "?",
                        max(ds).strftime("%m/%d") if ds else "?", fmt(g, dec), other,
                        "后一次" if kind == "消失" else "前一次", len(xs) - other,
                        "（那次导出平台提示“No pudimos mostrar algunos datos”——导出本身缺数据，"
                        "不是订单真的没了）" if D["S"].get("ventas_warning") else ""))
    if st:
        tr = {}
        for x in st:
            t = "%s → %s" % (x["before"]["status"] or "（空）", x["after"]["status"] or "（空）")
            tr[t] = tr.get(t, 0) + 1
        order.append("状态变化 %d 单：%s。" % (len(st), "；".join(
            "%s（%d 单）" % (t, n) for t, n in sorted(tr.items(), key=lambda z: -z[1])[:8])))
    if ref:
        d = sum(x["after"]["refund"] - x["before"]["refund"] for x in ref)
        order.append("退款金额变化 %d 单，合计 %s（买家退货/取消在第一次下载之后才处理完，"
                     "Ventas 把退款记回原订单的下单月）。" % (len(ref), fmt(d, dec)))
    other = [x for x in vd if "fields" in x and not {"status", "refund"} & set(x["fields"])]
    if other:
        tot = sum(x["after"]["total"] - x["before"]["total"] for x in other)
        shp = sum(x["after"]["ship_cost"] - x["before"]["ship_cost"] for x in other)
        order.append("另有 %d 单状态、退款都没变，只是金额列变了（%s）：Total 合计变化 %s，配送费合计变化 %s。"
                     % (len(other), "；".join(sorted(set(x["kind"].split("：")[1] for x in other))),
                        fmt(tot, dec), fmt(shp, dec)))
    if abs(m2["ing"] - m1["ing"]) > 0.004:
        order.append("GMV 变化 %s。" % fmt(m2["ing"] - m1["ing"], dec))
    if lrd["new_n"]:
        order.append("本月订单关联到的账单行新增 %d 行（%s）：现行口径的佣金、退货处理费按订单号跨账单取数，"
                     "后一次下载时还没关账的下月账单里新计的费用/冲销也会算回本月。"
                     % (lrd["new_n"], fmt(lrd["new_amount"], dec)))
    if lrd["gone_n"]:
        order.append("本月订单关联到的账单行少了 %d 行（%s）：账单行本身没变，是对应的订单不在后一次的 Ventas 里了，"
                     "现行口径就关联不上它们的佣金/配送费。" % (lrd["gone_n"], fmt(lrd["gone_amount"], dec)))
    if D1["S"].get("ventas_warning") or D2["S"].get("ventas_warning"):
        order.append("注意：平台自己提示 Ventas 导出可能缺数据（%s）。" % "；".join(
            "%s：%s" % (snapshot_label(D["folder"]), D["S"]["ventas_warning"])
            for D in (D1, D2) if D["S"].get("ventas_warning")))
    inv = {}
    for v, name in (("period", "账单期"), ("calendar", "自然月切")):
        diff, na, nb = ids[v]
        v1, v2 = D1["invs"][v]["view"], D2["invs"][v]["view"]
        files = sorted(set(v1["files"]) | set(v2["files"]))
        status = "；".join(v2["issues"]) or "用到的账单都已关账"
        if not diff:
            inv[v] = ("%s：两次下载的账单行逐行相同（%d 行；金额、Estado、通道、发票号都没变）。所用文件：%s。%s。"
                      % (name, na, "、".join(files), status))
        else:
            k = {}
            for x in diff:
                k[x["kind"]] = k.get(x["kind"], 0) + 1
            amt = sum(((x["after"] or {}).get("amount", 0.0) - (x["before"] or {}).get("amount", 0.0))
                      for x in diff)
            inv[v] = ("%s：账单行 %d → %d，有 %d 行不同（%s），金额合计变化 %s。%s。"
                      % (name, na, nb, len(diff), "、".join("%s %d" % z for z in k.items()),
                         fmt(amt, dec), status))
    return {"order": order, "period": inv["period"], "calendar": inv["calendar"]}


def analyse(store, month, folders, quiet=True, cache=None):
    """Every snapshot computed (or taken from cache {folder: compute() result})
    and every consecutive pair diffed."""
    snaps = []
    for f in folders:
        if cache and f in cache:
            snaps.append(cache[f])
            continue
        print("[读取] %s ← %s" % (snapshot_label(f), f))
        snaps.append(report.compute(store, f, month, quiet=quiet))
    kns = [report.key_numbers(D) for D in snaps]
    pairs = []
    for D1, D2 in zip(snaps, snaps[1:]):
        vd = ventas_diff(D1, D2)
        ids = dict((v, invoice_diff(D1, D2, v)) for v in ("period", "calendar"))
        lrd = linked_rows_diff(D1, D2)
        why = explain(D1, D2, vd, ids, lrd)
        k1, k2 = report.key_numbers(D1), report.key_numbers(D2)
        for v in ("period", "calendar"):
            why[v] += " " + mp_sentence(k1, k2, v, D1["dec"], k1["mp_source"], k2["mp_source"])
        pairs.append({"a": D1, "b": D2, "ventas": vd, "invoice": ids, "linked": lrd, "why": why})
    return {"store": store, "month": month, "snaps": snaps, "kns": kns, "pairs": pairs,
            "rows": metric_rows(kns)}


def default_folders(store, month, base=None):
    return month_folders(store_root(store, base), month)


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="python -m experiments.ml_invoice_basis.stability",
        description="实验：同一店同一月多次下载，现行订单口径和账单口径各变了多少、为什么变")
    ap.add_argument("--store", required=True, help="店铺名（downloads/ 下的文件夹名）")
    ap.add_argument("--month", required=True, help="会计月，如 2026-09")
    ap.add_argument("--folders", nargs="+", default=None,
                    help="要比较的下载目录（按给的顺序当作先后）；不给就取 downloads/<店>/ 下"
                         "所有 *_m<YYYYMM> 目录，按时间从早到晚")
    ap.add_argument("--base", default=None, help="店铺目录的上一级，默认是仓库的 downloads/")
    ap.add_argument("--out-dir", default=None,
                    help="输出目录，默认 downloads/data/reports/experiments/")
    ap.add_argument("--no-recalc", action="store_true", help="不用 LibreOffice 重算公式")
    a = ap.parse_args(argv)
    try:
        parse_month(a.month)
        folders = [os.path.abspath(os.path.expanduser(f)) for f in a.folders] if a.folders \
            else default_folders(a.store, a.month, a.base)
        for f in folders:
            if not os.path.isdir(f):
                raise ExperimentError("下载目录不存在：%s" % f)
        if len(folders) < 2:
            raise ExperimentError("要至少两次下载才能比较，只找到 %d 个：%s。用 --folders 指定。"
                                  % (len(folders), "、".join(folders) or
                                     "（%s 下没有 *_m%s 目录）" % (
                                         store_root(a.store, a.base), a.month.replace("-", ""))))
        A = analyse(a.store, a.month, folders)
        from experiments.ml_invoice_basis import workbook
        out = out_path(a.out_dir, a.store, a.month, STABILITY_SUFFIX)
        workbook.write_stability_workbook(A, out)
        print("[输出] %s" % out)
        print_summary(A)
        return workbook.finish(out, a.no_recalc)
    except ExperimentError as e:
        print("[错误] %s" % e)
        return 2


def print_summary(A):
    for p in A["pairs"]:
        print("\n%s → %s" % (snapshot_label(p["a"]["folder"]), snapshot_label(p["b"]["folder"])))
        for s in p["why"]["order"]:
            print("  订单口径：%s" % s)
        print("  %s" % p["why"]["period"])
        print("  %s" % p["why"]["calendar"])


if __name__ == "__main__":
    sys.exit(main())
