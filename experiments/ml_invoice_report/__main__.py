# -*- coding: utf-8 -*-
"""python -m experiments.ml_invoice_report --store <name> --month 2026-09 --folder <download folder>
       [--site mx|cl|pe] [--out-dir DIR] [--sku-costs FILE] [--no-recalc]

Reads the folder read-only, writes <store>_<YYYY-MM>_账单基准报表.xlsx. The site
comes from the folder's run_meta.json ("site"); older Mexican folders have none —
pass --site mx. Exit code: 0 ok, 1 error-level checks failed, 2 cannot run.
"""
import argparse
import os
import sys

from core import report_family as RF
from experiments.ml_invoice_report import build


def main(argv=None):
    ap = argparse.ArgumentParser(prog="python -m experiments.ml_invoice_report",
                                 description="MercadoLibre：以账单为基准的财务报表（实验）")
    ap.add_argument("--store", required=True)
    ap.add_argument("--month", required=True, help="YYYY-MM（按计费日的会计月）")
    ap.add_argument("--folder", required=True, help="该店该月的下载文件夹")
    ap.add_argument("--site", choices=("mx", "cl", "pe"), help="国家；不填则读 run_meta.json 的 site")
    ap.add_argument("--out-dir", help="输出目录（默认 downloads/data/reports/experiments/）")
    ap.add_argument("--sku-costs", help="SKU 成本表（.xlsx/.csv，可选；格式见 docs/REPORT_FAMILY.md §7）")
    ap.add_argument("--no-recalc", action="store_true", help="不用 LibreOffice 重算")
    a = ap.parse_args(argv)
    if not os.path.isdir(a.folder):
        print("找不到文件夹：%s" % a.folder)
        return 2
    site = build.site_of(a.folder, a.site)
    if not site:
        print("这个文件夹的 run_meta.json 没有写国家（site）。请加 --site mx / cl / pe。")
        return 2
    res = build.run(a.store, a.folder, a.month, site, a.out_dir, a.sku_costs, recalc=not a.no_recalc)
    for c in res["checks"]:
        tag = "跳过" if c["skipped"] else ("通过" if c["ok"] else ("错误" if c["severity"] == "error" else "注意"))
        print("[%s] %s：%s" % (tag, c["name"], c["detail"]))
    if res["recalc"] is not None:
        print("重算：%s" % res["recalc"][1])
    if res["scan"] is not None:
        n_f, n_err, n_none, bad = res["scan"]
        print("公式 %d 个，错误 %d 个，无缓存值 %d 个%s" % (n_f, n_err, n_none, ("：" + "；".join(bad)) if bad else ""))
    print("输出：%s" % res["out"])
    return 1 if RF.failed_errors(res["checks"]) else 0


if __name__ == "__main__":
    sys.exit(main())
