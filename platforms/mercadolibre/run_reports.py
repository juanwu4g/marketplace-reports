# -*- coding: utf-8 -*-
"""
生成月度财务报表：每店一份 + 全部店铺一份汇总。

    python -m platforms.mercadolibre.run_reports                        # 上个完整月份，白名单里全部店铺
    python -m platforms.mercadolibre.run_reports --month 2026-08
    python -m platforms.mercadolibre.run_reports --stores BOCINA_TA02 UNIT_TA04
    python -m platforms.mercadolibre.run_reports --no-per-store         # 只出汇总
    python -m platforms.mercadolibre.run_reports --dry-run              # 只列出会用哪个下载目录
    python -m platforms.mercadolibre.run_reports --region CL            # 智利店铺（默认 MX 墨西哥）；秘鲁 --region PE

店铺清单来自 config.json 的白名单，和下载流程同一个来源，不在这里另写一份。
每家店用它**最近一个文件齐全**的下载目录 —— 不是无脑取最新，因为一次
`--only ventas` 的运行会留下缺账单的目录。

区域（--region，默认 MX）决定三件事：不给 --stores 时"全部店铺"是哪几家
（store_config.stores_for_region）、全店汇总表的文件名（combined_name），以及
传给 report_build 的 --site（币种、账单周期、税务口径按国家走；下载目录
run_meta.json 记的站点与之不符的店铺会被剔除并在报表里记一笔）。

退出码
------
    0  要的报表全部写出，error 级校验全部通过，源数据没有被平台标为不完整
    1  要的报表全部写出，但有需要看一眼的：校验未过、公式错误、Ventas 数据不全，
       或汇总报表里缺了没有可用下载目录的店铺
    2  至少一份要的报表**没有写出来**（报错、超时、--strict 中止、写入失败、
       某家店没有可用的下载目录所以没有它的单店报表），或者没有任何一家店可用
    3  店铺不在白名单
    4  配置错误

1 和 2 的分界只有一条：文件是不是本次确实写出来了。以前两者都是 1，上层
（run_monthly）没法区分"桥B 照例没过"和"报表根本没出来"，而前者每个月都有。
"""
import argparse
from core import console  # noqa: F401  中文输出编码保护
import datetime
import os
import sys

from core import site_time
from core import store_config
from core.store_config import StoreConfigError, StoreNotAllowed
from platforms.mercadolibre import financial_report

from core.paths import ROOT
OUT_DIR = os.path.join(ROOT, "downloads", "data", "reports", "financial")


def prev_month(today=None):
    """上一个完整月份。当月还没过完，做月报没有意义。"""
    today = today or site_time.site_today()   # 平台那边的今天，见 site_time.py
    first = today.replace(day=1)
    last_prev = first - datetime.timedelta(days=1)
    return "%04d-%02d" % (last_prev.year, last_prev.month)


def combined_name(month, region="MX"):
    """全店汇总表的文件名。

    墨西哥的叫 MX_ML_全店汇总_<月>.xlsx —— 财务和生产上的脚本都认这个名字，
    不能改。别的国家换掉前缀：两国的汇总写在同一个输出目录里，同名就会互相
    覆盖，而且一份"全店汇总"只能是一个国家的（币种都不一样）。
    """
    return "%s_ML_全店汇总_%s.xlsx" % (str(region or "MX").upper(), month)


def store_names(stores=None, region="MX"):
    """这一次出哪几家店：给了 --stores 就是它们（过白名单），否则是这个区域
    的全部店铺。"""
    if stores:
        return [store_config.require_allowed(n) for n in stores]
    return store_config.stores_for_region(region)


def resolve(names, month=None):
    """把店铺名解析成 (名字, 下载目录)，并报告哪些用不了。

    传 month 是为了让每家店优先用**确实含那个月账单**的目录 —— 做历史月份时
    最新目录往往只有最近两期账单。
    """
    usable, unusable = [], []
    for n in names:
        folder, note = financial_report.latest_run(n, month=month)
        if folder:
            usable.append((n, folder))
            print("  %-22s %s" % (n, note))
        else:
            unusable.append((n, note))
            print("  %-22s [跳过] %s" % (n, note))
    return usable, unusable


def summarize_checks(checks):
    """把一份报表里未通过的校验压成一两行。返回 [文字, ...]。

    每条校验的完整说明和排查提示已经在上面那份报表自己的 [校验] 段里打印过
    一次，这里只负责让人在汇总里一眼看到"哪份报表、哪几项、哪几家店"。汇总报
    表一次覆盖 12 家店，逐条重打会把汇总淹掉，所以同一项校验合成一条、后面
    列店铺。error 级和 warn 级分开，因为只有前者影响退出码里的"校验未过"。
    """
    lines = []
    for severity, title in (("error", "未过"), ("warn", "提醒")):
        grouped, order = {}, []
        for c in checks:
            if c.get("ok") or c.get("skipped") or c.get("severity", "error") != severity:
                continue
            key = (c.get("key"), c.get("name"))
            if key not in grouped:
                grouped[key] = []
                order.append(key)
            if c.get("store") and c["store"] not in grouped[key]:
                grouped[key].append(c["store"])
        if order:
            lines.append("%s %d 项：%s" % (title, len(order), "；".join(
                "%s[%s]（%s）" % (name, key, "、".join(grouped[(key, name)]) or "—")
                for key, name in order)))
    return lines


def safe_build(*args, **kwargs):
    """调 financial_report.build，并保证拿回来的一定是一个带 outcome 的结果。

    build 自己承诺不抛异常，但这里是整批报表的循环：万一它真抛了（磁盘满、
    输出目录建不出来），不能让一家店把后面十一家和汇总一起带走，更不能让整个
    脚本以 Python 默认的退出码 1 结束 —— 那正好是"有校验未过"的码。
    """
    try:
        r = financial_report.build(*args, **kwargs)
    except Exception as e:
        out = kwargs.get("out") or (args[2] if len(args) > 2 else "")
        r = {"ok": False, "skipped": None, "returncode": None, "seconds": 0,
             "out": out, "wrote": False, "crashed": True, "checks": [],
             "output": "%s: %s" % (e.__class__.__name__, e)}
        print("    [错误] 生成报表时出错：%s" % r["output"])
    if not r.get("outcome"):
        r["outcome"] = financial_report.classify(r)
    return r


def exit_code(results, usable, unusable, per_store=True):
    """按文件头的约定，从各份报表的结果算出退出码（只管 0 / 1 / 2）。

    `results` 是 build 的结果列表，`usable` / `unusable` 是 resolve() 的两个返回
    值，`per_store` 表示这次是否要出单店报表。
    """
    if not usable:
        return 2
    if any(not financial_report.written(r) for r in results):
        return 2
    # 没有可用下载目录的店铺：要出单店报表时，它那一份就是"要了但没写出来"。
    # 只出汇总时文件是写出来了，但里面少了这家店 —— 不能算干净。
    if unusable:
        return 2 if per_store else 1
    if any(r.get("outcome") != "ok" for r in results):
        return 1
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", default=None,
                    help="会计期间，如 2026-08；默认为上一个完整月份")
    ap.add_argument("--stores", nargs="+", default=None,
                    help="店铺名；默认取 config.json 白名单里的全部")
    ap.add_argument("--region", type=str.upper, choices=["MX", "CL", "PE"],
                    default="MX",
                    help="哪个国家：MX 墨西哥（默认）/ CL 智利 / PE 秘鲁。决定不给 --stores 时"
                         "出哪些店、汇总表叫什么")
    ap.add_argument("--out-dir", default=OUT_DIR, help="报表输出目录")
    ap.add_argument("--no-per-store", action="store_true", help="不出单店报表")
    ap.add_argument("--no-combined", action="store_true", help="不出汇总报表")
    ap.add_argument("--metrics", action="store_true",
                    help="同时导出月度指标 JSON（用于累积指标库）")
    ap.add_argument("--strict", action="store_true",
                    help="任一 error 级校验未通过、或平台声明 Ventas 导出不完整，"
                         "就不写出那份报表（退出码 2）")
    ap.add_argument("--dry-run", action="store_true",
                    help="只列出会用哪个下载目录，不生成报表")
    args = ap.parse_args()

    month = args.month or prev_month()
    names = store_names(args.stores, args.region)
    if not names:
        raise StoreConfigError("config.json 里没有%s的店铺（白名单为空，或没有"
                               "标成 %s 的店），没有店铺可出报表。"
                               % (store_config.region_name(args.region),
                                  args.region))

    print("=" * 62)
    print("月度财务报表 —— %s%s，%d 家店铺"
          % ("" if args.region == "MX"
             else store_config.region_name(args.region) + " ",
             month, len(names)))
    print("=" * 62)
    usable, unusable = resolve(names, month=month)

    if not usable:
        print("\n没有一家店有可用的下载目录，无法生成报表。")
        return 2
    if args.dry_run:
        print("\n--dry-run：不生成报表。")
        return 0

    if not financial_report.has_libreoffice():
        # 公式本身照常写进文件，只是没有缓存值；Excel 打开时会自己算。
        # 说出来，免得看到"跳过重算"以为报表不完整。
        print("\n[提示] 未检测到 LibreOffice，跳过公式重算。")
        print("       公式已写入文件，Excel/WPS 打开时会自动计算；")
        print("       但本次无法自动检查公式是否报错。")

    results = []

    # ---- 单店报表：一家出问题不影响其余 ----
    if not args.no_per_store:
        for name, folder in usable:
            out = os.path.join(args.out_dir, "%s_%s.xlsx" % (name, month))
            print("\n" + "-" * 62)
            print("单店报表：%s" % name)
            print("-" * 62)
            r = safe_build(
                [(name, folder)], month, out, strict=args.strict,
                site=args.region.lower(),   # 国家传给报表：币种、账单周期、税务口径
                metrics_out=(os.path.join(args.out_dir, "%s_%s_指标.json"
                                          % (name, month)) if args.metrics else None))
            r["label"] = name
            results.append(r)

    # ---- 汇总报表：12 家店传进同一次调用 ----
    if not args.no_combined:
        out = os.path.join(args.out_dir, combined_name(month, args.region))
        print("\n" + "-" * 62)
        print("汇总报表：%d 家店铺" % len(usable))
        print("-" * 62)
        r = safe_build(
            usable, month, out, strict=args.strict,
            site=args.region.lower(),       # 一份汇总只属于一个国家，见 report_build --site
            metrics_out=(os.path.join(args.out_dir, "全店_%s_指标.json" % month)
                         if args.metrics else None))
        r["label"] = "全店汇总"
        results.append(r)

    # ---- 汇总 ----
    print("\n" + "=" * 62)
    print("报表生成结果")
    print("=" * 62)
    print("  %-22s %-10s %8s  %s" % ("报表", "结果", "用时", "文件"))
    for r in results:
        # 状态只从 outcome 来（定义在 financial_report.classify）。以前这里按
        # checks_failed / formula_errors 各自判断，--strict 中止的报表两个条件
        # 都沾边，会被显示成"有校验未过"并配一个并不存在的文件名。
        outcome = r["outcome"]
        wrote = financial_report.written(r)
        print("  %-22s %-10s %6ds  %s"
              % (r["label"], financial_report.OUTCOME_LABELS.get(outcome, outcome),
                 r["seconds"], os.path.basename(r["out"]) if wrote else "—"))
        if r.get("skipped"):
            print("    原因：%s" % r["skipped"])
        for line in summarize_checks(r.get("checks", [])):
            print("    %s" % line)
        if not wrote:
            if r.get("returncode") is not None:
                print("    report_build 退出码：%s" % r["returncode"])
            if r.get("diagnostic_log"):
                print("    完整输出已存：%s" % r["diagnostic_log"])
    if unusable:
        print("\n  未参与（没有可用下载目录）：")
        for n, why in unusable:
            print("    %-22s %s" % (n, why))

    print("\n  输出目录：%s" % os.path.abspath(args.out_dir))
    code = exit_code(results, usable, unusable, per_store=not args.no_per_store)
    missing = [r for r in results if not financial_report.written(r)]
    attention = [r for r in results
                 if r["outcome"] in ("checks_failed", "formula_errors")]
    incomplete = [r for r in results if r.get("data_incomplete")
                  and financial_report.written(r)]
    if missing:
        print("  %d/%d 份报表没有写出来：%s"
              % (len(missing), len(results), "、".join(r["label"] for r in missing)))
    if unusable:
        print("  %d 家店铺没有可用的下载目录，%s"
              % (len(unusable), "没有它们的单店报表，汇总里也不含它们"
                 if not args.no_per_store else "汇总里不含它们"))
    if attention:
        print("  %d/%d 份报表已写出，但有校验未过或公式错误，需要关注"
              % (len(attention), len(results)))
    if incomplete:
        print("  %d/%d 份报表已生成，但 Ventas 源数据不完整，需重新下载核对"
              % (len(incomplete), len(results)))
    if code == 0:
        print("  全部 %d 份报表生成成功" % len(results))
    return code


if __name__ == "__main__":
    try:
        sys.exit(main())
    except StoreConfigError as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)
    except StoreNotAllowed as e:
        print("\n店铺不在白名单：%s" % e)
        sys.exit(3)
