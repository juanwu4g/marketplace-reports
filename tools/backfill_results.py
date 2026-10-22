# -*- coding: utf-8 -*-
"""
补录月度结果文件：给本功能上线前就出过报表的月份，从已有的下载目录重出一次报表，
让每家店每月都有结果文件（core/results）—— 期间报表只读结果文件。不开浏览器、不下载。

    python -m tools.backfill_results --month 2026-09 --region MX              # 只列出要做什么（默认）
    python -m tools.backfill_results --from 2026-07 --to 2026-09 --region ALL
    python -m tools.backfill_results --month 2026-09 --region FA-CL --stores "DAB_CL-Fala (Coco)"
    python -m tools.backfill_results --month 2026-09 --region MX --apply      # 真的去做

规则
----
* 默认只列出（dry-run），什么都不写、什么都不删；加 --apply 才出报表。
* 已经有结果文件的店-月份默认跳过（不重复造版本）；--force 照样重出（多一个新版本）。
* 每家店用的下载目录和平时出报表一样挑（美客多 financial_report.latest_run，其他平台
  各自 run_reports 的 latest_folder）；没有可用目录的店列出原因、跳过。
* 出报表用的就是平时的 run_reports（子进程，一个单元一个月一次）：
  - 新布局（config.json 有 paths.data_root）：报表进 reports/monthly/<月>/<平台-国家>/，
    定稿、汇总表、版本说明、清理都照常（core/finalize）；
  - 旧布局：为了不覆盖当时已经发出去的报表，报表写到 <报表目录>/_backfill/<月>/
    （--out-dir 可改），只出单店报表；结果文件照常写进 downloads/data/results/。
* 结果文件的数字是**今天的代码**算出来的，不一定和当时发出的工作簿一模一样（规则改过
  的话）；版本说明和结果文件里记着代码版本。

退出码：0 全部成功（或 dry-run）；1 有单元的 run_reports 退出码是 1（有校验要看）；
2 有单元没出来；4 配置错误。
"""
import argparse
import os
import subprocess
import sys

from core import console  # noqa: F401  中文输出编码保护
from core import layout
from core import markets
from core import paths
from core import results
from core import store_config
from core.store_config import StoreConfigError


def months_of(args):
    if args.month:
        return sorted(set(args.month))
    if args.from_month and args.to_month:
        return results.months_between(args.from_month, args.to_month)
    raise StoreConfigError("要给 --month（可多个），或 --from 和 --to")


def units_of(regions):
    """--region 的值 -> 运行单元（美客多是 "MX"，其他平台是 "FA-CL"）。ALL = config 里有店的全部单元。"""
    out = []
    for r in regions:
        r = str(r).upper()
        if r == "ALL":
            out += [markets.unit_of_market(m) for m in store_config.markets_configured()]
        else:
            out.append(markets.unit_of_market(markets.unit_market(r)))
    return list(dict.fromkeys(out))


def stores_of(unit, wanted=None):
    if wanted:
        return [s for s in [store_config.require_allowed(x) for x in wanted] if store_config.unit_of(s) == unit]
    return list(store_config.stores_for_region(unit))


def find_folder(unit, store, month):
    """(download folder or None, note) — the same choice the report step makes."""
    platform = markets.unit_platform(unit)
    if platform == markets.DEFAULT_PLATFORM:
        # the report step falls back to any folder when none holds the month's bill;
        # a backfill must not make a month out of another month's data
        from platforms.mercadolibre import financial_report
        folder, note = financial_report.latest_run(store, month=month)
        want = (int(month[:4]), int(month[5:7]))
        if folder and want not in financial_report.billing_months(folder):
            return None, "没有含 %s 账单的下载目录（%s）" % (month, note)
        return folder, note
    import importlib
    rr = importlib.import_module(markets.platform(platform).package + ".run_reports")
    folder = rr.latest_folder(store, month)
    return folder, (os.path.basename(folder) if folder else rr.RUNNER.no_folder_reason % month)


def out_dir_for(month, override=None):
    """v2: the reports root (run_reports gets the month folder below); legacy: a
    separate folder so the reports sent at the time stay as they are."""
    if layout.mode() == layout.V2:
        return layout.reports_root(override)
    return os.path.abspath(override) if override else os.path.join(layout.reports_root(), "_backfill", month)


def report_dir(unit, month, out_dir):
    if layout.mode() == layout.V2:
        return layout.monthly_dir(month, markets.unit_market(unit), override=out_dir)
    return out_dir


def command(unit, month, stores, out_dir, py=None):
    """The run_reports command of one unit and month (as core/run_monthly builds it).
    Legacy: --no-combined (a summary of only the backfilled stores would mislead);
    new layout: the summary is rebuilt from every store's final (core/finalize)."""
    py = py or sys.executable
    platform = markets.unit_platform(unit)
    pkg = markets.platform(platform).package
    cmd = [py, "-u", "-m", pkg + ".run_reports", "--month", month]
    if platform == markets.DEFAULT_PLATFORM:
        cmd += ["--region", markets.split(markets.unit_market(unit))[1]]
    else:
        cmd += ["--market", markets.unit_market(unit)]
    cmd += ["--stores"] + list(stores)
    if layout.mode() != layout.V2:
        cmd.append("--no-combined")
    return cmd + ["--out-dir", report_dir(unit, month, out_dir)]


def plan(units, months, wanted=None, force=False, say=print):
    """[(unit, month, [stores to build], [(store, why skipped)])], printed as a table."""
    out = []
    for unit in units:
        names = stores_of(unit, wanted)
        for month in months:
            todo, skip = [], []
            for s in names:
                have = results.versions(s, month)
                fin = [b for b in sorted(set(v["basis"] for v in have)) if results.final(s, month, b)]
                folder, note = find_folder(unit, s, month)
                state = ("已有 %d 个版本%s" % (len(have), "，最终：%s" % "、".join(fin) if fin else "")) if have \
                    else "没有结果文件"
                if not folder:
                    skip.append((s, "没有可用的下载目录：%s" % note))
                elif have and not force:
                    skip.append((s, state + "（--force 才重出）"))
                else:
                    todo.append(s)
                say("  %-6s %s  %-26s %-14s %s" % (unit, month, s, state, folder or "—"))
            out.append((unit, month, todo, skip))
    return out


def main(argv=None, run=subprocess.call, say=print):
    ap = argparse.ArgumentParser(description="补录月度结果文件（从已有下载目录重出报表，不开浏览器）")
    ap.add_argument("--month", action="append", default=None, help="月份 2026-09，可重复")
    ap.add_argument("--from", dest="from_month", default=None)
    ap.add_argument("--to", dest="to_month", default=None)
    ap.add_argument("--region", nargs="+", required=True, help="MX / CL / PE / FA-CL / AZ-MX … 或 ALL")
    ap.add_argument("--stores", nargs="+", default=None, help="只做这几家店（默认单元里的全部店）")
    ap.add_argument("--out-dir", default=None, help="报表根目录（新布局）/ 输出目录（旧布局）")
    ap.add_argument("--force", action="store_true", help="已有结果文件的店-月份也重出")
    ap.add_argument("--apply", action="store_true", help="真的出报表（不加只列出）")
    a = ap.parse_args(argv)
    months = months_of(a)
    units = units_of(a.region)
    say("补录月度结果文件 —— %s，%s；数据布局：%s" % ("、".join(units), "、".join(months),
                                             "新布局" if layout.mode() == layout.V2 else "旧布局"))
    steps = plan(units, months, a.stores, a.force, say)
    for unit, month, todo, skip in steps:
        for s, why in skip:
            say("  跳过 %s %s %s：%s" % (unit, month, s, why))
    work = [(u, m, t) for u, m, t, _s in steps if t]
    if not work:
        say("\n没有要补录的店-月份。")
        return 0
    for unit, month, todo in work:
        say("  %s %s 要重出：%s" % (unit, month, "、".join(todo)))
        say("    %s" % " ".join(command(unit, month, todo, out_dir_for(month, a.out_dir))))
    if not a.apply:
        say("\n这是预演（dry-run），什么都没有写。确认无误后加 --apply 执行。")
        return 0
    worst = 0
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    for unit, month, todo in work:
        cmd = command(unit, month, todo, out_dir_for(month, a.out_dir))
        say("\n== %s %s：%d 家店 ==" % (unit, month, len(todo)))
        rc = run(cmd, cwd=paths.ROOT, env=env)
        say("   run_reports 退出码 %s" % rc)
        worst = max(worst, 1 if rc == 1 else (0 if rc == 0 else 2))
    return worst


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (StoreConfigError, markets.UnknownMarket, ValueError) as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)
