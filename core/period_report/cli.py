# -*- coding: utf-8 -*-
"""期间报表（季度 / 全年 / 自定义）：只用各月的定稿月度结果，不重新读平台数据。

    python run_period.py --quarter 2026-Q3                 # 第 3 季度，全部平台和国家
    python run_period.py --year 2026 --region MX           # 2026 全年，美客多墨西哥
    python run_period.py --from 2026-07 --to 2026-11 --region FA-CL
    python run_period.py --quarter 2026-Q3 --region MX --basis order    # 美客多订单基准
    python run_period.py --quarter 2026-Q3 --stores BOCINA_TA02 EWTTO_SM
    python run_period.py --quarter 2026-Q3 --preview       # 只看各店各月有没有完整版本，不写文件

每个市场（平台 + 国家，币种不同）一个文件，放在 layout.period_dir 里：
    新布局  <报表根目录>\\quarterly\\2026-Q3\\ML-MX\\MX_ML_2026-Q3_账单基准.xlsx
    旧布局  <报表目录>\\期间报表\\quarterly\\2026-Q3\\ML-MX\\…
--out-dir 换掉报表根目录（旧布局是那个平铺的报表目录），和窗口里「报表输出到」一样。

退出码
    0  都生成了，每个店铺每个月都有完整版本（--preview：都有完整版本）
    1  生成了，但有店铺月份没有完整版本（没计入合计，见报表①和⑩）
    2  一份都没生成，或者有市场出错（文件被 Excel 打开着、参数不对 …）
"""
import argparse
import os
import sys

from core import console  # noqa: F401  中文输出编码保护
from core import markets
from core import results as R
from core.period_report import output as G
from core.period_report import model as MD

OK, INCOMPLETE, FAILED = 0, 1, 2


def parse_args(argv=None):
    ap = argparse.ArgumentParser(prog="run_period.py", description="生成期间报表（季度 / 全年 / 自定义）")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--quarter", metavar="YYYY-Qn", help="季度，如 2026-Q3")
    g.add_argument("--year", metavar="YYYY", help="全年，如 2026")
    g.add_argument("--from", dest="first", metavar="YYYY-MM", help="自定义期间的第一个月（和 --to 一起用）")
    ap.add_argument("--to", dest="last", metavar="YYYY-MM", help="自定义期间的最后一个月")
    ap.add_argument("--region", type=lambda v: str(v).upper().replace(" ", ""), default="ALL",
                    help="哪些平台和国家：MX / CL / PE（美客多）、FA-CL、AZ-MX …，ALL = 全部（默认），"
                         "几个用逗号隔开，如 MX,FA-CL")
    ap.add_argument("--stores", nargs="+", default=None, metavar="NAME",
                    help="只要这几家店（每家店进它所在市场的那份报表）")
    ap.add_argument("--basis", choices=["bill", "order"], default=None,
                    help="美客多用哪个基准：bill 账单基准（默认，会计入账）/ order 订单基准；别的平台只有一种")
    ap.add_argument("--out-dir", default=None,
                    help="报表根目录。不给就用窗口里「报表输出到」选的，再不行用默认目录")
    ap.add_argument("--preview", action="store_true", help="只列出各店各月有没有完整版本，不写文件")
    args = ap.parse_args(argv)
    if args.first and not args.last:
        ap.error("--from 要和 --to 一起用")
    if args.last and not args.first:
        ap.error("--to 要和 --from 一起用")
    try:
        args.period = (MD.quarter(args.quarter) if args.quarter else MD.year(args.year) if args.year
                       else MD.custom(args.first, args.last))
    except ValueError as e:
        ap.error(str(e))
    try:
        args.markets = G.markets_of(args.region)
    except markets.UnknownMarket as e:
        ap.error("--region %s：%s" % (args.region, e))
    if not args.markets:
        ap.error("--region 是空的")
    return args


def stray_stores(stores, chosen, config_path=None):
    """--stores that belong to none of the chosen markets (by results tree or config)."""
    from core import store_config
    out = []
    for s in stores or []:
        home = None
        for m in chosen:
            pf, cc = markets.split(m)
            if s in R.stores(pf, cc, config_path):
                home = m
                break
        if home is None:
            try:
                home = markets.market_id(*store_config.market_parts(s))
            except Exception:                            # noqa: BLE001
                home = None
        if home not in chosen:
            out.append(s)
    return out


def grid_lines(D):
    """'  BOCINA_TA02   2026-07 ✓ 完整   2026-08 ✗ 无完整版本' — one line per store."""
    width = max([len(s) for s in D.stores] + [4])
    out = []
    for s in D.stores:
        parts = []
        for m in D.months:
            f = D.final(s, m)
            parts.append("%s %s" % (m, ("✓ " + R.GRADE_CN[f["grade"]["level"]]) if f else "✗ 无完整版本"))
        out.append("  %s  %s" % (s.ljust(width), "　".join(parts)))
    return out


def main(argv=None, out=None):
    out = out or sys.stdout

    def say(text=""):
        out.write(text + "\n")
        out.flush()

    args = parse_args(argv)
    from core.run_monthly import resolve_out_dir         # the window's 「报表输出到」, one rule
    from core.store_config import StoreConfigError
    try:
        out_dir, src = resolve_out_dir(args.out_dir)
    except StoreConfigError as e:
        say("[错误] %s" % e)
        return FAILED
    p = args.period
    say("期间报表：%s（%s）" % (MD.period_cn(p), MD.KIND_CN[p.kind]))
    say("范围　　：%s%s" % ("、".join(markets.label(m) for m in args.markets),
                         "；只要 %s" % "、".join(args.stores) if args.stores else ""))
    if not args.preview:
        say("报表根目录：%s（%s）" % (os.path.abspath(out_dir), src))
    stray = stray_stores(args.stores, args.markets)
    if stray:
        say("[错误] 这些店铺不属于所选的平台和国家（或者 config.json 里没有、也没有任何月度结果）：%s"
            % "、".join(stray))
        return FAILED
    codes, written = [], 0
    for m in args.markets:
        pf, _ = markets.split(m)
        basis = MD.basis_of(pf, args.basis)
        say("")
        say("━━━━━━━━ %s%s ━━━━━━━━" % (markets.label(m), "（%s）" % MD.basis_cn(basis)
                                          if MD.basis_cn(basis) else ""))
        stores = [s for s in args.stores or []
                  if s not in stray_stores([s], [m])] if args.stores else None
        if args.stores and not stores:
            say("  （所列店铺都不在这个市场，跳过）")
            continue
        if args.preview:
            try:
                D = MD.gather(m, p, basis, stores)
            except Exception as e:                       # noqa: BLE001
                say("  [错误] 读月度结果时出错：%s: %s" % (e.__class__.__name__, e))
                codes.append(FAILED)
                continue
            if not any(c.final or c.newest for c in D.cells.values()):
                say("  期间内没有任何月度结果。")
                continue
            for line in grid_lines(D):
                say(line)
            for c in D.missing():
                say("  ✗ %s %s：%s" % (c.store, c.month, c.reason()))
            codes.append(INCOMPLETE if D.missing() else OK)
            written += 1
            continue
        res = G.generate(p, m, basis, stores, out_dir)
        if res["status"] == G.EMPTY:
            say("  " + res["message"])
            continue
        if res["data"] is not None and res["status"] != G.ERROR:
            for line in grid_lines(res["data"]):
                say(line)
        for line in G.summary_lines(res):
            say(line)
        if res["status"] == G.ERROR:
            codes.append(FAILED)
        else:
            written += 1
            codes.append(INCOMPLETE if res["status"] == G.INCOMPLETE else OK)
    say("")
    if not written and FAILED not in codes:
        say("结果：一份都没生成 —— 所选范围在这个期间里没有任何月度结果。先出月度报表（「立即运行」）。")
        return FAILED
    code = max(codes) if codes else FAILED
    say({OK: "结果：%s，每个店铺每个月都有完整版本。" % ("都看过了" if args.preview else "都生成了"),
         INCOMPLETE: "结果：%s，但有店铺月份没有完整版本（没计入合计，见上面 ✗ 和报表的①、⑩）。"
                     % ("看过了" if args.preview else "已生成"),
         FAILED: "结果：有市场没生成（见上面的错误）。"}[code])
    return code


if __name__ == "__main__":
    sys.exit(main())
