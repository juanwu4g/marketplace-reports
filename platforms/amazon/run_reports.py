# -*- coding: utf-8 -*-
"""
Amazon monthly financial reports: one workbook per store + one combined.

    python -m platforms.amazon.run_reports --month YYYY-MM --market AZ-MX
    python -m platforms.amazon.run_reports --month YYYY-MM --stores "<store>" --no-combined
    python -m platforms.amazon.run_reports --dry-run

The shared contract (core/run_monthly treats every platform alike). Each store
uses its newest download folder **for that month**
(downloads/<store>/<stamp>_m<YYYYMM>/) that has a run_meta.json — change
latest_folder() if the report needs a particular file.

Exit codes
    0  every workbook written, no error-level check failed
    1  every workbook written, but something needs a look
    2  a workbook that was asked for was not written, or no store had data
    3  store not on the allowlist
    4  configuration error
"""
import argparse
import datetime
import glob
import os
import sys
import time
import traceback

from core import console  # noqa: F401  中文输出编码保护
from core import markets
from core import report_kit
from core import store_config
from core.paths import ROOT
from core.store_config import StoreConfigError, StoreNotAllowed
from platforms.amazon import periods

PLATFORM = "AZ"
DEFAULT_MARKET = "AZ-MX"
OUT_DIR = os.path.join(ROOT, "downloads", "data", "reports", "financial")
DOWNLOADS = os.path.join(ROOT, "downloads")
OUTCOME_LABELS = {"ok": "成功", "checks_failed": "有校验未过", "formula_errors": "公式错误",
                  "strict_aborted": "严格模式未写出", "crashed": "生成出错", "skipped": "跳过"}
WRITTEN = ("ok", "checks_failed", "formula_errors")


def latest_folder(store, month, base=None):
    y, m = periods.parse_month(month)
    hits = sorted(glob.glob(os.path.join(base or DOWNLOADS, store, "*_m%04d%02d" % (y, m))))
    for d in reversed(hits):
        if os.path.exists(os.path.join(d, "run_meta.json")):
            return d
    return None


def resolve(names, month, base=None):
    usable, unusable = [], []
    for n in names:
        d = latest_folder(n, month, base)
        if d:
            usable.append((n, d))
            print("  %-26s %s" % (n, d))
        else:
            unusable.append((n, "没有 %s 的下载目录" % month))
            print("  %-26s —— 没有可用的下载目录" % n)
    return usable, unusable


def safe_build(pairs, month, out, strict=False):
    from platforms.amazon import report_build
    t0 = time.time()
    try:
        r = report_build.build(pairs, month, out, strict=strict)
    except Exception as e:
        traceback.print_exc()
        return {"outcome": "crashed", "checks": [], "out": out, "error": str(e),
                "seconds": int(time.time() - t0)}
    if r["outcome"] in WRITTEN:
        ok, msg = report_kit.recalc(out)
        if ok:
            _n, n_err, _none, bad = report_kit.scan_formula_errors(out)
            if n_err:
                r["outcome"] = "formula_errors"
                r["formula_errors"] = bad
        else:
            print("  [提示] %s" % msg)
    r["seconds"] = int(time.time() - t0)
    return r


def exit_code(results, usable, unusable, per_store=True):
    if not usable:
        return 2
    if any(r["outcome"] not in WRITTEN for r in results):
        return 2
    if unusable:
        return 2 if per_store else 1
    if any(r["outcome"] != "ok" for r in results):
        return 1
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Amazon monthly financial reports")
    ap.add_argument("--month", default=None, help="会计期间，如 2026-09；默认上一个完整月份")
    ap.add_argument("--stores", nargs="+", default=None)
    ap.add_argument("--market", type=str.upper, default=DEFAULT_MARKET)
    ap.add_argument("--region", type=str.upper, default=None,
                    help="国家代码（= --market AZ-<代码>）")
    ap.add_argument("--out-dir", default=OUT_DIR)
    ap.add_argument("--no-per-store", action="store_true")
    ap.add_argument("--no-combined", action="store_true")
    ap.add_argument("--strict", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args(argv)
    if a.region:
        a.market = markets.market_id(PLATFORM, a.region)
    a.market = markets.normalize(a.market)
    if markets.split(a.market)[0] != PLATFORM:
        raise StoreConfigError("--market %s 不是 亚马逊 的市场" % a.market)
    country = markets.split(a.market)[1]
    month = a.month or periods.previous_month(datetime.date.today())
    if a.stores:
        names = [store_config.require_allowed(n) for n in a.stores]
    else:
        names = store_config.stores_for_market(a.market)
    store_config.require_platform(names, PLATFORM)
    if not names:
        raise StoreConfigError("config.json 里没有 %s 的店铺" % markets.label(a.market))
    print("=" * 62)
    print("月度财务报表 —— %s %s，%d 家店铺" % (markets.label(a.market), month, len(names)))
    print("=" * 62)
    usable, unusable = resolve(names, month)
    if not usable:
        print("\n没有一家店有可用的下载目录，无法生成报表。")
        return 2
    if a.dry_run:
        return 0
    results = []
    if not a.no_per_store:
        for name, folder in usable:
            out = os.path.join(a.out_dir, report_kit.store_workbook_name(name, month))
            print("\n单店报表：%s" % name)
            r = safe_build([(name, folder)], month, out, strict=a.strict)
            r["label"] = name
            results.append(r)
    if not a.no_combined:
        out = os.path.join(a.out_dir, report_kit.combined_workbook_name(country, PLATFORM, month))
        print("\n汇总报表：%d 家店铺" % len(usable))
        r = safe_build(usable, month, out, strict=a.strict)
        r["label"] = "全店汇总"
        results.append(r)
    print("\n" + "=" * 62)
    print("报表生成结果")
    print("=" * 62)
    for r in results:
        print("  %-26s %-10s %5ds  %s" % (r["label"], OUTCOME_LABELS.get(r["outcome"], r["outcome"]),
                                         r.get("seconds", 0),
                                         os.path.basename(r["out"]) if r["outcome"] in WRITTEN else "—"))
        for c in r.get("checks", []):
            if not c["ok"] and not c["skipped"]:
                print("    [%s] %s：%s" % (c["severity"], c["name"], c["detail"][:150]))
        if r.get("error"):
            print("    错误：%s" % r["error"][:200])
    for n, why in unusable:
        print("  未参与：%-22s %s" % (n, why))
    print("\n  输出目录：%s" % os.path.abspath(a.out_dir))
    return exit_code(results, usable, unusable, per_store=not a.no_per_store)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except StoreNotAllowed as e:
        print("\n店铺不在白名单：%s" % e)
        sys.exit(3)
    except (StoreConfigError, markets.UnknownMarket) as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)
