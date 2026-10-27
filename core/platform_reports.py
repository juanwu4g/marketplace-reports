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
                      (core/layout.run_folders_for_month) holding the first;
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
import json
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
    from core import layout
    return layout.raw_root()


def default_out_dir(month=None, market=None):
    """Legacy: the flat report folder. New layout: the month folder of the market
    (<reports>/monthly/<month>/<PF-CC>) — a hand-run report must land where
    core/finalize publishes and cleans, not in <reports>/_versions/…"""
    from core import layout
    if month and market and layout.mode() == layout.V2:
        return layout.monthly_dir(month, market)
    return layout.reports_root()


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
        from core import layout
        y, m = self.periods.parse_month(month)
        hits = sorted(p for _n, p in layout.run_folders_for_month(store, "%04d-%02d" % (y, m), base=base))
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

    # ------------------------------------------------------------ results files

    def results_document(self, name, folder, month, market, version, r):
        """The monthly results document of one store build (core/results)."""
        from core import grading
        from core import results as RS
        rb = importlib.import_module(self.package + ".report_build")
        country = markets.split(market)[1]
        cur, dec = getattr(rb, "CUR", ""), getattr(rb, "DECIMALS", 2)
        rec = (r.get("results") or {}).get(name) or {}
        meta = rec.get("meta") or {}
        missing = [x for x in (meta.get("missing") or []) if x and x.get("form")]
        checks = RS.plain(r.get("checks") or [])
        if r["outcome"] in WRITTEN and rec.get("figures") is not None:
            g = grading.grade(self.code, "main", True, checks, missing,
                              formula_errors=r["outcome"] == "formula_errors",
                              unknown_download=grading.download_unknown(rec.get("meta")))
            return RS.document(self.code, market, country, cur, dec, name, month, "main", version,
                               datetime.datetime.now(), [folder], r["out"], g, checks, missing,
                               figures=rec["figures"])
        why = OUTCOME_LABELS.get(r["outcome"], r["outcome"])
        if r["outcome"] in WRITTEN:
            why = "结果数字取不出来：%s" % (rec.get("error") or "报表没有返回结果数字").splitlines()[0]
        elif r.get("error"):
            why += "：%s" % r["error"][:200]
        g = grading.grade(self.code, "main", False, why_not_written=why)
        return RS.document(self.code, market, country, cur, dec, name, month, "main", version,
                           datetime.datetime.now(), [folder], None, g, checks, missing)

    def write_results(self, name, folder, month, market, version, r):
        """Write the store's results file; a problem is printed, never raised."""
        from core import results as RS
        try:
            rec = (r.get("results") or {}).get(name) or {}
            if rec.get("error"):
                print("  [结果数字取不出来] %s" % rec["error"])
            doc = self.results_document(name, folder, month, market, version, r)
            path = RS.write(doc)
            print("  [结果] %s · %s" % (RS.GRADE_CN[doc["grade"]["level"]], path))
            return path
        except Exception as e:                 # noqa: BLE001 — the workbook is out; say what went wrong
            traceback.print_exc()
            print("  [结果文件写入失败] %s: %s" % (e.__class__.__name__, e))
            return None

    def finalize_step(self, a, month, country, stores):
        """core/finalize in the new layout: finals, month folder, 全店汇总 from the
        finals, 版本说明.txt, clean-up. Returns the summary's build results."""
        from core import finalize
        try:
            all_stores = store_config.stores_for_market(a.market)
        except Exception:                      # noqa: BLE001 — the summary then covers the stores at hand
            all_stores = []
        unit = finalize.Unit(self.code, a.market, month, a.out_dir, ("main",),
                             lambda store, m, b: report_kit.store_workbook_name(store, m),
                             lambda m, b: report_kit.combined_workbook_name(country, self.code, m))

        def build_combined(basis, pairs, dst):
            print("\n汇总报表：%d 家店铺（按各店最终版本）" % len(pairs))
            folders = [(store, (doc.get("source_folders") or [None])[0]) for store, doc in pairs]
            r = self.safe_build(folders, month, dst, strict=a.strict)
            r["label"] = "全店汇总"
            return r
        try:
            out = finalize.run(unit, stores, all_stores, build_combined, no_combined=a.no_combined)
            # 定稿没能复制进月份文件夹（文件被打开）：main 把退出码至少设为 1
            a.publish_failed = out.get("publish_failed") or []
            return out["combined"]
        except Exception as e:                 # noqa: BLE001
            traceback.print_exc()
            print("\n[错误] 整理最终版本时出错：%s: %s" % (e.__class__.__name__, e))
            return []

    # ------------------------------------------------------------ command line

    def parser(self):
        ap = argparse.ArgumentParser(description="%s monthly financial reports" % self.name_en)
        ap.add_argument("--month", default=None, help="会计期间，如 2026-09；默认上一个完整月份")
        ap.add_argument("--stores", nargs="+", default=None)
        ap.add_argument("--market", type=str.upper, default=self.default_market)
        ap.add_argument("--region", type=str.upper, default=None,
                        help="国家代码（= --market %s-<代码>）" % self.code)
        ap.add_argument("--out-dir", default=None,
                        help="输出目录；默认：旧布局是报表目录，新布局是该月该市场的月份文件夹")
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
        if not a.out_dir:
            a.out_dir = default_out_dir(month, a.market)
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
        # one stamp per invocation: every store workbook is version <run>__<stamp> and
        # gets a monthly results file (core/results). In the new layout the store
        # workbooks go to <month folder>/_versions/<version>/ and core/finalize makes
        # the month folder (finals, 全店汇总 from them, 版本说明, clean-up).
        from core import finalize
        from core import results as RS
        stamp = RS.built_stamp()
        v2 = finalize.is_v2()
        results = []
        if not a.no_per_store:
            for name, folder in usable:
                version = RS.version_id(folder, stamp)
                where = finalize.version_dir(a.out_dir, version) if v2 else a.out_dir
                out = os.path.join(where, report_kit.store_workbook_name(name, month))
                print("\n单店报表：%s" % name)
                r = self.safe_build([(name, folder)], month, out, strict=a.strict)
                r["label"] = name
                results.append(r)
                self.write_results(name, folder, month, a.market, version, r)
        if not v2 and not a.no_combined:
            out = os.path.join(a.out_dir,
                               report_kit.combined_workbook_name(country, self.code, month))
            print("\n汇总报表：%d 家店铺" % len(usable))
            r = self.safe_build(usable, month, out, strict=a.strict)
            r["label"] = "全店汇总"
            results.append(r)
        results += self.finalize_step(a, month, country, [n for n, _f in usable])
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
        code = exit_code(results, usable, unusable, per_store=not a.no_per_store)
        unpublished = getattr(a, "publish_failed", None) or []
        if unpublished:
            print("  %d 个文件没能更新进月份文件夹（正被打开？关闭后重出）：%s"
                  % (len(unpublished), "、".join(os.path.basename(p) for p in unpublished)))
            code = max(code, 1)
        # 机器可读的最后一行（runlog.MARKERS）：run_monthly 用它把报表这一步的问题说准。
        print("[REPORT_SUMMARY_JSON] " + json.dumps({
            "written": [r["label"] for r in results if r["outcome"] in WRITTEN],
            "missing": [r["label"] for r in results if r["outcome"] not in WRITTEN],
            "unusable": [[n, why] for n, why in unusable],
            "attention": [r["label"] for r in results
                          if r["outcome"] in ("checks_failed", "formula_errors")]
            + ["未能更新：%s" % os.path.basename(p) for p in unpublished],
            "incomplete": []}, ensure_ascii=False))
        return code
