# -*- coding: utf-8 -*-
"""
生成月度财务报表：每店一份 + 全部店铺一份汇总。

    python -m platforms.mercadolibre.run_reports                        # 上个完整月份，白名单里全部店铺
    python -m platforms.mercadolibre.run_reports --month 2026-08
    python -m platforms.mercadolibre.run_reports --stores BOCINA_TA02 UNIT_TA04
    python -m platforms.mercadolibre.run_reports --no-per-store         # 只出汇总
    python -m platforms.mercadolibre.run_reports --dry-run              # 只列出会用哪个下载目录
    python -m platforms.mercadolibre.run_reports --region CL            # 智利店铺（默认 MX 墨西哥）；秘鲁 --region PE
    python -m platforms.mercadolibre.run_reports --basis bill           # 只出账单基准（默认两种都出）

两种基准（--basis，默认 both；config.json → reports.ml_basis 可改默认）
-------------------------------------------------------------------------
* 账单基准（bill，会计入账口径）：platforms/mercadolibre/bill_report，文件名加
  `_账单基准` 后缀 —— `<店>_<月>_账单基准.xlsx`、`MX_ML_全店汇总_<月>_账单基准.xlsx`。
* 订单基准（order，经营分析参考）：本文件一直出的那份，文件名不变。
两者独立计算、互不引用；一种出错不影响另一种。账单基准先出、先列。

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
import json
import os
import sys

from core import site_time
from core import store_config
from core.store_config import StoreConfigError, StoreNotAllowed
from platforms.mercadolibre import financial_report

from core.paths import ROOT
from core import finalize
from core import layout
from core import report_kit as K
from core import results as RS
from platforms.mercadolibre import results_files
# 默认输出目录。import 时 config.json 读不了就先记 None，不在这里抛：run_monthly
# 一开头就 import 本文件，抛在这里它连运行记录和告警都写不了（只剩一段
# traceback、退出码 1）。真跑本脚本时 __main__ 会再算一次，读不了就按配置错误退 4。
try:
    OUT_DIR = layout.reports_root()
except Exception:                             # noqa: BLE001
    OUT_DIR = None
BASES = ("bill", "order")                 # 账单基准是会计入账口径，先出、先列
BASIS_CN = {"bill": "账单基准", "order": "订单基准"}


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


def latest_folder(store, month, base=None):
    """这家店做 month 月报表要用的下载目录（和 resolve() 同一条规则：
    financial_report.latest_run），没有返回 None。tools/sku_cost_template 用它。"""
    return financial_report.latest_run(store, base=base, month=month)[0]


def configured_bases(config_path=None):
    """config.json → reports.ml_basis（["bill", "order"] 的子集），没写就两种都出。"""
    import io
    import json
    path = config_path or store_config.CONFIG_PATH
    try:
        with io.open(path, encoding="utf-8") as fh:
            node = (json.load(fh).get("reports") or {}).get("ml_basis")
    except (IOError, OSError, ValueError, AttributeError):
        node = None
    if isinstance(node, str):
        node = [node]
    picked = [b for b in BASES if isinstance(node, list) and b in node]
    return picked or list(BASES)


def bases_from_arg(arg, config_path=None):
    if arg in BASES:
        return [arg]
    if arg == "both":
        return list(BASES)
    return configured_bases(config_path)


def bill_reports(usable, month, region, out_dir, per_store=True, combined=True, stamp=None,
                 store_dir=None, prepared_out=None):
    """账单基准：每店一份 + 全店汇总一份。每家店只读一次数据，单店和汇总共用。
    返回结果列表（outcome / label / out / checks / seconds），和订单基准的同一种。

    stamp：本次出报表的时间戳（core/results.built_stamp）—— 给了就为每份单店报表写
    月度结果文件，版本号 <下载目录名>__<stamp>，两种基准共用。store_dir(folder)：
    单店报表写到哪个目录（新布局是 _versions/<版本>/，默认 out_dir）。
    prepared_out：给一个 dict，就把每家店读好的 (S, checks) 按 (店, 目录) 放进去
    （新布局的汇总表由 core/finalize 按最终版本重出，同一目录的店不必再读一遍）。"""
    from core import sku_costs as SC
    from platforms.mercadolibre.bill_report import build as BB
    site = region.lower()
    results, prepared = [], []
    store_dir = store_dir or (lambda folder: out_dir)

    def save(make):
        if stamp:
            results_files.write_from(make)

    def failed(name, folder, why, checks=(), meta=None):
        save(lambda: results_files.failed_document(name, month, "bill", region, version_id(folder, stamp),
                                                   folder, why, checks, meta))
    try:
        table = SC.load_configured()
    except Exception as e:                    # 成本表读不出来：照常出报表，成本行留空
        print("  [警告] 商品成本表读不出来，本次账单基准报表不填成本：%s" % e)
        table = None
    print("\n" + "=" * 62)
    print("账单基准报表（会计入账口径）—— %d 家店铺" % len(usable))
    print("=" * 62)
    for name, folder in usable:
        print("\n" + "-" * 62)
        print("单店报表：%s" % name)
        print("-" * 62)
        print("  [账单基准] 读取 %s" % folder)
        meta_site = BB.site_of(folder)
        if meta_site and meta_site != site:
            r = {"label": "%s 账单基准" % name, "outcome": "skipped", "checks": [], "out": "", "seconds": 0,
                 "skipped": "下载目录记的站点是 %s，不是 %s" % (meta_site, site)}
            print("    [跳过] %s" % r["skipped"])
            results.append(r)
            if per_store:
                failed(name, folder, r["skipped"])
            continue
        def prep(name=name, folder=folder):
            S, checks = BB.prepare(name, folder, month, site, table)
            return {"S": S, "checks": checks}
        got = BB.guarded("%s 账单基准" % name, prep)
        if got.get("crashed"):
            print("    [错误] 读取或计算出错：%s" % got["output"])
            results.append(got)
            if per_store:
                failed(name, folder, "读取或计算出错：%s" % got["output"])
            continue
        S, checks = got["S"], got["checks"]
        prepared.append((S, checks))
        if prepared_out is not None:
            prepared_out[(name, os.path.normcase(os.path.abspath(folder)))] = (S, checks)
        if per_store:
            # 写结果文件时也在①页盖完整度：保存后、重算前（见 BB.write_store）
            def stamp_banner(path, S=S, checks=checks, folder=folder, formula_errors=False):
                if not stamp:                 # no results files, no banner
                    return None
                try:
                    g = results_files.bill_grade(S, checks, formula_errors)
                except Exception as e:        # noqa: BLE001 — a banner must not cost the report
                    print("    [提示] 完整度标记没写上：%s: %s" % (e.__class__.__name__, e))
                    return None
                return results_files.stamp_banner(path, g, version_id(folder, stamp), folder, "bill",
                                                  log=lambda s: print("    " + s))
            r = BB.guarded("%s 账单基准" % name, BB.write_store, S, checks, store_dir(folder), table,
                           before_recalc=stamp_banner if stamp else None)
            _print_bill(r)
            results.append(r)
            if r.get("crashed"):
                failed(name, folder, "生成报表时出错：%s" % r["output"], checks, S.get("meta"))
            else:
                scan = r.get("scan")
                formula_errors = bool(scan is not None and scan[1])
                if formula_errors and stamp_banner(r["out"], formula_errors=True):
                    # 公式错误压成“有缺口”：重新盖章后再重算一次补回缓存值
                    print("    [重算] %s" % K.recalc(r["out"])[1])
                save(lambda: results_files.bill_document(S, checks, version_id(folder, stamp), r["out"],
                                                         formula_errors=formula_errors))
    if combined and prepared:
        print("\n" + "-" * 62)
        print("汇总报表：%d 家店铺（账单基准）" % len(prepared))
        print("-" * 62)
        r = BB.guarded("全店汇总 账单基准", BB.write_combined, prepared, month, region, out_dir, table,
                       order_combined=combined_name(month, region))
        _print_bill(r)
        results.append(r)
    return results


def version_id(folder, stamp):
    """本次出报表的版本号：<下载目录名>__<stamp>（core/results.version_id）。"""
    return RS.version_id(folder, stamp) if stamp else None


def _print_bill(r):
    if r.get("crashed"):
        print("    [错误] 生成报表时出错：%s" % r["output"])
        return
    for c in r.get("checks") or []:
        if not (c["ok"] or c["skipped"]) and c["severity"] == "error":
            print("    ✗  %-12s %s：%s" % (c["store"], c["name"], c["detail"]))
    if r.get("scan") is not None and r["scan"][1]:
        print("    [警告] 重算后有 %d 个公式报错：%s" % (r["scan"][1], "；".join(r["scan"][3][:5])))
    print("    [输出] %s" % r["out"])


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


def store_workbook_name(store, month, basis="order"):
    """单店报表的文件名：订单基准 <店>_<月>.xlsx，账单基准加 _账单基准。"""
    if basis == "bill":
        from platforms.mercadolibre.bill_report import build as BB
        return BB.workbook_name(store, month)
    return "%s_%s.xlsx" % (store, month)


def order_failed(name, folder, month, region, version, r):
    """订单基准的单店报表没写出来：子进程没机会写结果文件，这里补一个“失败”版本。"""
    if not version or os.path.isfile(RS.path_for(name, month, version, "order")):
        return None
    why = financial_report.OUTCOME_LABELS.get(r.get("outcome"), r.get("outcome"))
    if r.get("skipped"):
        why += "：%s" % r["skipped"]
    return results_files.write_from(lambda: results_files.failed_document(
        name, month, "order", region, version, folder, why, r.get("checks") or ()))


def finalize_step(args, month, bases, usable, unusable, prepared):
    """core/finalize：定出每家店每种基准的最终版本。新布局下还要把定稿复制进月份
    文件夹、按最终版本出全店汇总（两种基准各一份）、写版本说明.txt、清理。
    返回汇总表的生成结果（和上面的单店结果同一种），没有就是 []。"""
    from platforms.mercadolibre.bill_report import build as BB
    region = args.region
    market = results_files.market_of(region)
    try:
        all_stores = store_config.stores_for_region(region)
    except Exception:                         # noqa: BLE001 — the summary then covers the stores at hand
        all_stores = []
    unit = finalize.Unit(
        "ML", market, month, args.out_dir, tuple(bases), store_workbook_name,
        lambda m, b: BB.combined_name(m, region) if b == "bill" else combined_name(m, region))

    def build_combined(basis, pairs, dst):
        print("\n" + "-" * 62)
        print("汇总报表：%d 家店铺（%s，按各店最终版本）" % (len(pairs), BASIS_CN[basis]))
        print("-" * 62)
        folders = [(store, (doc.get("source_folders") or [None])[0]) for store, doc in pairs]
        if basis == "bill":
            r = BB.guarded("全店汇总 账单基准", _bill_combined, folders, month, region, dst, prepared)
            _print_bill(r)
            return r
        r = safe_build(folders, month, dst, strict=args.strict, site=region.lower(),
                       metrics_out=(os.path.join(args.out_dir, "全店_%s_指标.json" % month)
                                    if args.metrics else None))
        r["label"] = "全店汇总" + (" 订单基准" if len(bases) > 1 else "")
        return r

    try:
        out = finalize.run(unit, [n for n, _f in usable], all_stores, build_combined,
                           no_combined=args.no_combined)
        args.publish_failed = out.get("publish_failed") or []     # main 据此把退出码至少设为 1
    except Exception as e:                    # noqa: BLE001 — the workbooks are out; say what went wrong
        import traceback
        print("\n[错误] 整理最终版本时出错：%s: %s" % (e.__class__.__name__, e))
        traceback.print_exc(file=sys.stdout)
        return []
    return out["combined"]


def _bill_combined(folders, month, region, dst, prepared):
    """账单基准的全店汇总，从各店最终版本的下载目录出；本次已读过的目录不再读。"""
    from core import sku_costs as SC
    from platforms.mercadolibre.bill_report import build as BB
    try:
        table = SC.load_configured()
    except Exception:                         # noqa: BLE001 — as bill_reports: no cost rows
        table = None
    prep = []
    for store, folder in folders:
        got = prepared.get((store, os.path.normcase(os.path.abspath(folder))))
        prep.append(got or BB.prepare(store, folder, month, region.lower(), table))
    return BB.write_combined(prep, month, region, out_dir=os.path.dirname(dst), table=table,
                             name=os.path.basename(dst), order_combined=combined_name(month, region))


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
    ap.add_argument("--out-dir", default=None,
                    help="报表输出目录；默认：旧布局是报表目录，新布局是该月该国的月份文件夹")
    ap.add_argument("--no-per-store", action="store_true", help="不出单店报表")
    ap.add_argument("--no-combined", action="store_true", help="不出汇总报表")
    ap.add_argument("--metrics", action="store_true",
                    help="同时导出月度指标 JSON（用于累积指标库）")
    ap.add_argument("--strict", action="store_true",
                    help="任一 error 级校验未通过、或平台声明 Ventas 导出不完整，"
                         "就不写出那份报表（退出码 2）")
    ap.add_argument("--dry-run", action="store_true",
                    help="只列出会用哪个下载目录，不生成报表")
    ap.add_argument("--basis", choices=["both", "bill", "order"], default=None,
                    help="出哪种基准：both 两种（默认；config.json reports.ml_basis 可改）/ "
                         "bill 只出账单基准（会计入账口径）/ order 只出订单基准")
    args = ap.parse_args()
    bases = bases_from_arg(args.basis)

    month = args.month or prev_month()
    if not args.out_dir:
        # 新布局：手动出的报表也进月份文件夹，core/finalize 才会复制定稿、清理
        args.out_dir = layout.monthly_dir(month, results_files.market_of(args.region))
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
    print("报表基准：%s" % "、".join(BASIS_CN[b] for b in bases))
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

    # 一次出报表一个时间戳：两种基准的每份单店报表都是版本 <下载目录名>__<stamp>，
    # 各写一个月度结果文件（core/results）。新布局下单店报表先写进月份文件夹的
    # _versions/<版本>/，汇总表由 core/finalize 按各店的最终版本出（见文件尾）。
    stamp = RS.built_stamp()
    v2 = finalize.is_v2()

    def store_dir(folder):
        return finalize.version_dir(args.out_dir, version_id(folder, stamp)) if v2 else args.out_dir

    results = []
    prepared = {}
    if "bill" in bases:
        results += bill_reports(usable, month, args.region, args.out_dir,
                                per_store=not args.no_per_store, combined=not args.no_combined and not v2,
                                stamp=stamp, store_dir=store_dir, prepared_out=prepared)
        if "order" in bases:
            print("\n" + "=" * 62)
            print("订单基准报表（经营分析参考）—— %d 家店铺" % len(usable))
            print("=" * 62)

    # ---- 单店报表：一家出问题不影响其余 ----
    if "order" in bases and not args.no_per_store:
        for name, folder in usable:
            out = os.path.join(store_dir(folder), store_workbook_name(name, month))
            print("\n" + "-" * 62)
            print("单店报表：%s" % name)
            print("-" * 62)
            r = safe_build(
                [(name, folder)], month, out, strict=args.strict,
                site=args.region.lower(),   # 国家传给报表：币种、账单周期、税务口径
                metrics_out=(os.path.join(args.out_dir, "%s_%s_指标.json"
                                          % (name, month)) if args.metrics else None),
                results_version=version_id(folder, stamp))
            r["label"] = name + (" 订单基准" if len(bases) > 1 else "")
            results.append(r)
            if not financial_report.written(r):
                order_failed(name, folder, month, args.region, version_id(folder, stamp), r)

    # ---- 汇总报表：12 家店传进同一次调用（新布局：core/finalize 按最终版本出） ----
    if "order" in bases and not args.no_combined and not v2:
        out = os.path.join(args.out_dir, combined_name(month, args.region))
        print("\n" + "-" * 62)
        print("汇总报表：%d 家店铺" % len(usable))
        print("-" * 62)
        r = safe_build(
            usable, month, out, strict=args.strict,
            site=args.region.lower(),       # 一份汇总只属于一个国家，见 report_build --site
            metrics_out=(os.path.join(args.out_dir, "全店_%s_指标.json" % month)
                         if args.metrics else None))
        r["label"] = "全店汇总" + (" 订单基准" if len(bases) > 1 else "")
        results.append(r)

    # ---- 最终版本：每家店每种基准；新布局下复制定稿、出汇总表、写版本说明、清理 ----
    results += finalize_step(args, month, bases, usable, unusable, prepared)

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
    unpublished = getattr(args, "publish_failed", None) or []
    if unpublished:
        print("  %d 个文件没能更新进月份文件夹（正被打开？关闭后重出）：%s"
              % (len(unpublished), "、".join(os.path.basename(p) for p in unpublished)))
        code = max(code, 1)
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
    print_report_summary(results, unusable, missing, attention, incomplete)
    return code


def print_report_summary(results, unusable, missing, attention, incomplete):
    """最后一行：机器可读的汇总（runlog.MARKERS 里的 REPORT_SUMMARY_JSON），给
    run_monthly 把"报表这一步出了什么问题"说准 —— 哪家店没有下载目录、哪份报表
    没写出来，而不是笼统的"崩溃 / 超时"。"""
    doc = {"written": [r["label"] for r in results if r not in missing],
           "missing": [r["label"] for r in missing],
           "unusable": [[n, why] for n, why in unusable],
           "attention": [r["label"] for r in attention],
           "incomplete": [r["label"] for r in incomplete]}
    print("[REPORT_SUMMARY_JSON] " + json.dumps(doc, ensure_ascii=False))


def run_as_script(entry=None):
    """命令行入口。没接住的异常一律按硬失败退 2：Python 自己的退出码 1 在
    run_monthly 眼里是"跑完了、有要关注的项"，会把一次崩溃当成正常结束。"""
    global OUT_DIR
    try:
        if OUT_DIR is None:
            OUT_DIR = layout.reports_root()   # import 时没算出来：再算一次，坏了就抛配置错误
        sys.exit((entry or main)())
    except StoreConfigError as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)
    except StoreNotAllowed as e:
        print("\n店铺不在白名单：%s" % e)
        sys.exit(3)
    except Exception:                         # noqa: BLE001
        import traceback
        traceback.print_exc()
        print("\n出报表时程序出错（见上面的错误信息），按硬失败退出。")
        sys.exit(2)


if __name__ == "__main__":
    run_as_script()
