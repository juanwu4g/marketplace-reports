# -*- coding: utf-8 -*-
"""Experiment: one store's month on the invoice basis, next to the current order basis.

    python -m experiments.ml_invoice_basis --store <name> --month YYYY-MM
        [--folder DIR] [--base DIR] [--out-dir DIR] [--no-recalc] [--no-stability]
        [--threshold-pct 1.0] [--threshold-abs 1000]

--folder   the download folder to use; default: the newest folder of the month
           under downloads/<store>/ that holds the month's Facturación
--base     the folder that holds the store folders (default: the repo's downloads/)
--out-dir  default: downloads/data/reports/experiments/
--threshold-pct / --threshold-abs
           materiality of ② 差异摘要: a difference is flagged when it is at least
           this % of the current basis' net revenue, or at least this amount in the
           store's currency (defaults 1 % and 1,000; both editable in the workbook)

Writes `<store>_<YYYY-MM>_账单口径实验.xlsx`. When the folder has sibling
downloads of the same month (`*_m<YYYYMM>`), ⑨ compares them; otherwise ⑨ says
how to run the stability command. Exit 0 on success, 1 when the written file has
formula errors, 2 on a problem the person must fix (message in Chinese).
"""

import argparse
import os
import sys

from core import console  # noqa: F401  中文输出编码保护
from platforms.mercadolibre.report.catalog import fmt

from experiments.ml_invoice_basis import accounting, report, stability, workbook
from experiments.ml_invoice_basis import mercadopago as MP
from experiments.ml_invoice_basis.common import (REPORT_SUFFIX, ExperimentError, month_folders,
    out_path, parse_month, pick_folder, store_root)


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="python -m experiments.ml_invoice_basis",
        description="实验：把美客多的损益改成账单口径（当月 Facturación + Notas de crédito），并与现行订单口径对比")
    ap.add_argument("--store", required=True, help="店铺名（downloads/ 下的文件夹名）")
    ap.add_argument("--month", required=True, help="会计月，如 2026-09")
    ap.add_argument("--folder", default=None, help="下载目录；不给就自动挑")
    ap.add_argument("--base", default=None, help="店铺目录的上一级，默认是仓库的 downloads/")
    ap.add_argument("--out-dir", default=None,
                    help="输出目录，默认 downloads/data/reports/experiments/")
    ap.add_argument("--no-recalc", action="store_true", help="不用 LibreOffice 重算公式")
    ap.add_argument("--no-stability", action="store_true", help="⑨ 不比较同月的其他下载")
    ap.add_argument("--threshold-pct", type=float, default=accounting.DEFAULT_PCT,
                    help="② 重要性标准一：差额占现行净收入的百分比，默认 %g" % accounting.DEFAULT_PCT)
    ap.add_argument("--threshold-abs", type=float, default=accounting.DEFAULT_ABS,
                    help="② 重要性标准二：差额金额（店铺币种），默认 %g" % accounting.DEFAULT_ABS)
    a = ap.parse_args(argv)
    try:
        y, m = parse_month(a.month)
        if a.folder:
            folder = os.path.abspath(os.path.expanduser(a.folder))
            if not os.path.isdir(folder):
                raise ExperimentError("下载目录不存在：%s" % folder)
        else:
            folder = pick_folder(a.store, a.month, a.base)
            if not folder:
                raise ExperimentError("在 %s 下没找到含 %s 账单（Reporte_Facturacion_MercadoLibre_…）的下载目录。"
                                      "用 --folder 指定，或先下载这个月的账单。"
                                      % (store_root(a.store, a.base), a.month))
        print("[读取] %s ← %s" % (a.store, folder))
        D = report.compute(a.store, folder, a.month)
        for line in D["log"].splitlines():
            if "警告" in line:
                print(line)
        A, reason = None, None
        if a.no_stability:
            reason = "按 --no-stability 没有比较其他下载。"
        else:
            sibs = month_folders(os.path.dirname(folder), a.month)
            if len(sibs) >= 2 and folder in sibs:
                A = stability.analyse(a.store, a.month, sibs, cache={folder: D})
            else:
                reason = ("这个下载目录旁边没有同一月（*_m%s）的其他下载，无法比较。"
                          % a.month.replace("-", ""))
        out = out_path(a.out_dir, a.store, a.month, REPORT_SUFFIX)
        workbook.write_report_workbook(D, A, out, reason, a.threshold_pct, a.threshold_abs)
        print("[输出] %s" % out)
        print_summary(D, a.threshold_pct, a.threshold_abs)
        return workbook.finish(out, a.no_recalc)
    except ExperimentError as e:
        print("[错误] %s" % e)
        return 2


def print_summary(D, pct=accounting.DEFAULT_PCT, absv=accounting.DEFAULT_ABS):
    P, C, m = D["invs"]["period"], D["invs"]["calendar"], D["S"]["m"]
    dec = D["dec"]
    print("\n[期间] 账单期 %s – %s；自然月切 %s – %s"
          % (P["view"]["start"].date(), P["view"]["end"].date(), C["view"]["start"].date(),
             C["view"]["end"].date()))
    for t in P["view"]["issues"] + C["view"]["issues"]:
        print("  ⚠ %s" % t)
    print("%-18s %16s %16s %16s" % ("", "现行订单口径", "账单期", "自然月切"))
    print("%-18s %16s %16s %16s" % ("订单 / 开票销售笔数", fmt(m["orders_unique"], 0),
                                     fmt(P["n"]["n_sales"], 0), fmt(C["n"]["n_sales"], 0)))
    for lab, a, b, c in (("商品销售额", m["ing"], P["n"]["gross_sales"], C["n"]["gross_sales"]),
                         ("取消与退款", -m["refund"], -P["n"]["rev_sales"], -C["n"]["rev_sales"]),
                         ("平台费用合计", workbook._cur_fees(D), -P["n"]["fee_total"],
                          -C["n"]["fee_total"]),
                         ("经营贡献毛利", D["S"]["contrib"], P["n"]["contribution"],
                          C["n"]["contribution"])):
        print("%-18s %16s %16s %16s" % (lab, fmt(a, dec), fmt(b, dec), fmt(c, dec)))
    AP, AC = D["acc"]["period"], D["acc"]["calendar"]
    for lab, b, c in (("调整后（含税）", AP["contrib_adj"], AC["contrib_adj"]),
                      ("调整后（不含税）", AP["vat"]["contrib_net"], AC["vat"]["contrib_net"])):
        print("%-18s %16s %16s %16s" % (lab, "", fmt(b, dec), fmt(c, dec)))
    t = AC["rev"]
    print("\n[MercadoPago] %s" % MP.source_text(D["mp"]))
    if AC["mp"]["available"]:
        print("  冲销（自然月切）：估算 %s，实际 %s（%d/%d 笔有实际数），误差 %s；优惠券 %s；应计 %s"
              % (fmt(t["est"], dec), fmt(t["value"], dec), t["n_known"], t["n"], fmt(t["err"], dec),
                 fmt(AC["coupon"], dec), fmt(AC["accrual"]["total"], dec)))
    rows = accounting.summary_rows(D)
    base = accounting.base_net_revenue(D)
    fl = accounting.flagged(rows, base, pct, absv)
    print("[重要性] 净收入 %s；标准 %g%% 或 %s：%d 项超过" % (fmt(base, dec), pct, fmt(absv, dec), len(fl)))
    for x in fl:
        print("  ⚠ %-30s %16s  %s" % (x["label"][:30], fmt(x["value"], dec), x["cat"]))


if __name__ == "__main__":
    sys.exit(main())
