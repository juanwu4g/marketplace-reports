# -*- coding: utf-8 -*-
"""
The shared monthly-report runner of every platform except MercadoLibre: one
workbook per store + one combined.

A platform's `platforms/<pkg>/run_reports.py` is a thin module: it builds one
`ReportRunner` with what is genuinely its own and hands `main` to it:

    RUNNER = platform_reports.ReportRunner(
        "FA", __package__, default_market="FA-CL", periods=periods,
        folder_markers=lambda month: ("NewReportTransaction_*.xlsx",),
        no_folder_reason="没有 %s 的下载目录（或目录里没有交易明细）")

What a platform declares (the hooks)
    code              two-letter platform code (core/markets.py)
    package           its package; the workbook comes from <package>.report_build
                      (imported only when a workbook is built), whose
                      build(pairs, month, out, strict=) returns {"outcome", "checks", "out", …}
    default_market    --market when none is given
    periods           module with parse_month(month), previous_month(date)
    folder_markers    month -> file names / glob patterns, tried in order: a
                      store's folder is the newest of the month
                      (downloads/<store>/<stamp>_m<YYYYMM>/) holding the first;
                      if none does, the newest holding the second; …
                      Default ("run_meta.json",): the newest finished run.
    no_folder_reason  the 未参与 reason when no folder qualifies ("%s" = month)
    name_zh/name_en   labels; default: the platform's names in core/markets.py

Exit codes (main returns 0–2; a refused store or bad configuration is raised
and mapped by run_as_script, as the `if __name__ == "__main__"` block did)
    0  every workbook written, no error-level check failed
    1  every workbook written, but something needs a look (failed check, formula
       error, a store without data left out of the combined workbook)
    2  a workbook that was asked for was not written, or no store had data
    3  store not on the allowlist
    4  configuration error

The stdout lines `单店报表：<store>`, `汇总报表：<n> 家店铺` and the label
`全店汇总` are read by core/runlog.py — keep them.
"""
import argparse
import datetime
import glob
import importlib
import os
import sys
import time
import traceback

from core import markets
from core import paths
from core import report_kit
from core import store_config
from core.store_config import StoreConfigError, StoreNotAllowed

OUTCOME_LABELS = {"ok": "成功", "checks_failed": "有校验未过", "formula_errors": "公式错误",
                  "strict_aborted": "严格模式未写出", "crashed": "生成出错", "skipped": "跳过"}
WRITTEN = ("ok", "checks_failed", "formula_errors")
DEFAULT_MARKERS = ("run_meta.json",)
DEFAULT_NO_FOLDER = "没有 %s 的下载目录"


def downloads_dir():
    return os.path.join(paths.ROOT, "downloads")


def default_out_dir():
    return os.path.join(paths.ROOT, "downloads", "data", "reports", "financial")


def _holds(folder, marker):
    path = os.path.join(folder, marker)
    if glob.has_magic(marker):
        return bool(glob.glob(path))
    return os.path.exists(path)


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


def run_as_script(main):
    """The `python -m platforms.<pkg>.run_reports` entry: exit with main()'s code;
    a store off the allowlist -> 3, a configuration error -> 4."""
    try:
        sys.exit(main())
    except StoreNotAllowed as e:
        print("\n店铺不在白名单：%s" % e)
        sys.exit(3)
    except (StoreConfigError, markets.UnknownMarket) as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)


class ReportRunner(object):
    """One platform's report runner. See the module docstring for the hooks."""

    def __init__(self, code, package, default_market, periods, folder_markers=None,
                 no_folder_reason=DEFAULT_NO_FOLDER, name_zh=None, name_en=None):
        self.code = code
        self.package = package
        self.default_market = default_market
        self.periods = periods
        self.folder_markers = folder_markers or (lambda month: DEFAULT_MARKERS)
        self.no_folder_reason = no_folder_reason
        self._name_zh = name_zh
        self._name_en = name_en

    def _registered(self, field):
        try:
            return getattr(markets.platform(self.code), field)
        except markets.UnknownMarket:
            return self.code

    @property
    def name_zh(self):
        return self._name_zh or self._registered("name_zh")

    @property
    def name_en(self):
        return self._name_en or self._registered("name_en")

    # ------------------------------------------------------------ download folders

    def latest_folder(self, store, month, base=None):
        y, m = self.periods.parse_month(month)
        hits = sorted(glob.glob(os.path.join(base or downloads_dir(), store,
                                             "*_m%04d%02d" % (y, m))))
        for marker in self.folder_markers(month):
            for d in reversed(hits):
                if _holds(d, marker):
                    return d
        return None

    def resolve(self, names, month, base=None):
        usable, unusable = [], []
        for n in names:
            d = self.latest_folder(n, month, base)
            if d:
                usable.append((n, d))
                print("  %-26s %s" % (n, d))
            else:
                unusable.append((n, self.no_folder_reason % month))
                print("  %-26s —— 没有可用的下载目录" % n)
        return usable, unusable

    # ------------------------------------------------------------ building

    def safe_build(self, pairs, month, out, strict=False):
        report_build = importlib.import_module(self.package + ".report_build")
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

    # ------------------------------------------------------------ command line

    def parser(self):
        ap = argparse.ArgumentParser(description="%s monthly financial reports" % self.name_en)
        ap.add_argument("--month", default=None, help="会计期间，如 2026-09；默认上一个完整月份")
        ap.add_argument("--stores", nargs="+", default=None)
        ap.add_argument("--market", type=str.upper, default=self.default_market)
        ap.add_argument("--region", type=str.upper, default=None,
                        help="国家代码（= --market %s-<代码>）" % self.code)
        ap.add_argument("--out-dir", default=default_out_dir())
        ap.add_argument("--no-per-store", action="store_true")
        ap.add_argument("--no-combined", action="store_true")
        ap.add_argument("--strict", action="store_true")
        ap.add_argument("--dry-run", action="store_true")
        return ap

    def main(self, argv=None):
        a = self.parser().parse_args(argv)
        if a.region:
            a.market = markets.market_id(self.code, a.region)
        a.market = markets.normalize(a.market)
        if markets.split(a.market)[0] != self.code:
            raise StoreConfigError("--market %s 不是 %s 的市场" % (a.market, self.name_zh))
        country = markets.split(a.market)[1]
        month = a.month or self.periods.previous_month(datetime.date.today())
        if a.stores:
            names = [store_config.require_allowed(n) for n in a.stores]
        else:
            names = store_config.stores_for_market(a.market)
        store_config.require_platform(names, self.code)
        if not names:
            raise StoreConfigError("config.json 里没有 %s 的店铺" % markets.label(a.market))
        print("=" * 62)
        print("月度财务报表 —— %s %s，%d 家店铺" % (markets.label(a.market), month, len(names)))
        print("=" * 62)
        usable, unusable = self.resolve(names, month)
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
                r = self.safe_build([(name, folder)], month, out, strict=a.strict)
                r["label"] = name
                results.append(r)
        if not a.no_combined:
            out = os.path.join(a.out_dir,
                               report_kit.combined_workbook_name(country, self.code, month))
            print("\n汇总报表：%d 家店铺" % len(usable))
            r = self.safe_build(usable, month, out, strict=a.strict)
            r["label"] = "全店汇总"
            results.append(r)
        print("\n" + "=" * 62)
        print("报表生成结果")
        print("=" * 62)
        for r in results:
            print("  %-26s %-10s %5ds  %s" % (
                r["label"], OUTCOME_LABELS.get(r["outcome"], r["outcome"]),
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
