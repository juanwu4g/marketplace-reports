# -*- coding: utf-8 -*-
"""
One Falabella store, one month: open the store, check the language, fetch the
seven forms, explain anything missing.

    python -m platforms.falabella.run_downloads --store "DAB_CL-Fala (Coco)" --month 2026-09
    python -m platforms.falabella.run_downloads --store ... --only billing stock

Output: downloads/<store>/<YYYYMMDD_HHMMSS>_m<YYYYMM>/ with the raw files,
run_meta.json (machine-readable: market, month, expected / missing / flags) and
REPORT.md (for people). Same layout and keys as MercadoLibre's, so the runner,
the window and the report side read both the same way.

Raises only when the store cannot be worked at all (not allowed, wrong
platform, browser did not open, wrong language). Route failures are recorded
in result["errors"] and as cause codes, and the store carries on.
"""
import argparse
import datetime
import io
import json
import os
import sys
import time

from core import console  # noqa: F401  中文输出编码保护
from core import download_diag
from core import site_time
from core import store_config
from core.paths import ROOT
from platforms.falabella import forms
from platforms.falabella import page_kit as pk
from platforms.falabella import periods
from platforms.falabella import site as fa_site

PLATFORM = "FA"
ROUTE_NAMES = ("order", "billing", "stock", "billing1", "billing2")


class Collector(download_diag.Collector):
    """download_diag's collector, but the page language comes from the seller
    centre's data-localefull: <html lang> says "en" on these pages even when the
    screen is Spanish, and download_diag would call every failure LANGUAGE_WRONG."""

    locale = None

    def give_up(self, driver, form, signal, **kw):
        ev = download_diag.Collector.give_up(self, driver, form, signal, **kw)
        if isinstance(ev, dict):
            ev["html_lang"] = ev.get("lang")
            ev["lang"] = self.locale or ""
        return ev


def run_folder(store_name, month, stamp=None, base=None):
    """Absolute folder (CDP resolves a relative download path against the
    browser's own working directory, and files would vanish)."""
    stamp = stamp or datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    y, m = periods.parse_month(month)
    root = base or os.path.join(ROOT, "downloads")
    return os.path.abspath(os.path.join(root, store_name, "%s_m%04d%02d" % (stamp, y, m)))


def expected(month, only=None):
    forms_ = periods.form_ids(month)
    if only:
        keep = set()
        for r in only:
            keep.add(r)
            if r == "billing1":
                keep.add("invoice1")
            if r == "billing2":
                keep.add("invoice2")
        forms_ = [f for f in forms_ if periods.route_of(f) in keep]
    return forms_


def run_store(client, store_name, month, out_dir=None, only=None, collect_timeout=240,
              close_when_done=True, log=print):
    store_config.require_allowed(store_name)
    store_config.require_platform([store_name], PLATFORM)
    country = store_config.country_of(store_name)
    facts = fa_site.get(country)
    site_time.set_site(facts["tz"], facts["utc_offset"])
    out_dir = out_dir or run_folder(store_name, month)
    os.makedirs(out_dir, exist_ok=True)
    log("输出目录：%s" % out_dir)
    began = time.time()
    exp = expected(month, only)
    result = {"store": store_name, "platform": PLATFORM,
              "market": store_config.market_of(store_name), "country": country,
              "month": month, "out_dir": out_dir, "files": {}, "errors": [],
              "pending": [], "expected": exp, "missing": [], "flags": {},
              "cautions": [], "seconds": 0}
    diag = Collector(out_dir, store_name)

    store = client.find_store(store_name)
    log("店铺：%s（%s）" % (store.get("browserName"), store.get("platform_name")))
    store_config.check_ziniao_platform(store_name, store.get("platform_name"), PLATFORM)

    with client.open_store(store, close_when_done=close_when_done) as session:
        d = session.driver
        pk.set_download_dir(d, out_dir)
        ctx = forms.Ctx(d, out_dir, diag, month, facts, today=site_time.site_today(),
                        collect_timeout=collect_timeout, log=log)
        try:
            forms.check_language(ctx)
            diag.locale = ctx.flags.get("locale")
            log("卖家：%s %s，界面 %s" % (ctx.flags.get("seller"), ctx.flags.get("seller_name"),
                                     ctx.flags.get("locale")))
            want = set(periods.route_of(f) for f in exp)
            if "order" in want:
                forms.orders(ctx)
            if "billing" in want:
                forms.guarded(ctx, "billing", ["billing"], forms.billing)
            if "stock" in want:
                forms.guarded(ctx, "stock", ["stock"], forms.stock)
            if "billing1" in want or "invoice1" in want:
                forms.billing1(ctx)
            if "billing2" in want or "invoice2" in want:
                forms.billing2(ctx)
        except forms.LanguageWrong as e:
            ctx.errors.append("language: %s" % e)
            ev = download_diag.exception_evidence(e, step="language")
            ev["lang"] = ctx.flags.get("locale") or ""
            for f in exp:
                diag.items.setdefault(f, {"evidence": dict(ev), "period": None,
                                          "required": True})
        except Exception as e:
            # the browser died or the first page never rendered: the whole store
            ctx.errors.append("store: %s: %s" % (e.__class__.__name__, str(e)[:200]))
            ev = download_diag.exception_evidence(e, step="store")
            for f in exp:
                if f not in ctx.files:
                    diag.items.setdefault(f, {"evidence": dict(ev), "period": None,
                                              "required": True})
        result["files"] = ctx.files
        result["errors"] = ctx.errors
        result["pending"] = [dict(p, store=store_name, out_dir=out_dir, platform=PLATFORM,
                                  report=p["form"])
                             for p in ctx.pending]
        result["flags"] = dict(ctx.flags)
        result["no_data"] = dict(ctx.no_data)

    result["missing"] = diagnose(result, diag)
    result["seconds"] = int(time.time() - began)
    write_run_meta(out_dir, result, began)
    write_report(out_dir, result)
    return result


def diagnose(result, diag):
    recs = download_diag.diagnose_store(
        result["expected"], collector=diag, entries=result["pending"],
        collected=result["files"], out_dir=result["out_dir"])
    for r in recs:
        r["in_report"] = periods.in_report(r["form"])
    return recs


def write_run_meta(out_dir, result, began):
    meta = {"platform": PLATFORM, "market": result["market"], "country": result["country"],
            "site": result["country"].lower(), "month": result["month"],
            "store": result["store"],
            "downloaded_at": datetime.datetime.fromtimestamp(began).isoformat(timespec="seconds"),
            "files": result["files"], "expected": result["expected"],
            "missing": result["missing"], "flags": result["flags"],
            "cautions": result.get("cautions") or [], "errors": result["errors"]}
    try:
        with io.open(os.path.join(out_dir, "run_meta.json"), "w", encoding="utf-8") as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)
    except Exception as e:
        print("  [警告] run_meta.json 写入失败：%s" % str(e)[:80])
    return meta


def read_run_meta(out_dir):
    try:
        with io.open(os.path.join(out_dir, "run_meta.json"), encoding="utf-8") as fh:
            return json.load(fh) or {}
    except Exception:
        return {}


def update_run_meta(out_dir, **fields):
    meta = read_run_meta(out_dir)
    meta.update(fields)
    try:
        with io.open(os.path.join(out_dir, "run_meta.json"), "w", encoding="utf-8") as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)
    except Exception as e:
        print("  [警告] run_meta.json 更新失败：%s" % str(e)[:80])
    return meta


FORM_ZH = {"order": "订单（Lista de Órdenes）", "billing": "交易明细（Transacciones）",
           "stock": "库存（Fulfillment 产品）", "billing1": "平台费用发票明细（Factura falabella.com）",
           "billing2": "代开销售单据（Liquidación factura）",
           "invoice1": "平台费用发票 PDF", "invoice2": "代开销售单据 PDF"}


def write_report(out_dir, result):
    L = ["# %s · Falabella %s · %s" % (result["store"], result["country"], result["month"]), "",
         "下载时间：%s，用时 %d 秒" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                                result.get("seconds") or 0), ""]
    sw = result["flags"].get("language_switched")
    if sw:
        L.append("⚠ 开始时界面语言是 %s，脚本已在页面最下方把它改回 %s。有人在用这家店时改了语言；"
                 "反复出现的话，请提醒用这家店的人保持西班牙语。" % (sw.get("from"), sw.get("to")))
        L.append("")
    if result["flags"].get("seller"):
        L.append("卖家：%s %s（界面 %s）" % (result["flags"].get("seller"),
                                      result["flags"].get("seller_name") or "",
                                      result["flags"].get("locale") or "?"))
        L.append("")
    L.append("## 文件")
    L.append("")
    for f in result["expected"]:
        files = result["files"].get(f) or []
        L.append("- %s `%s`：%s" % (FORM_ZH.get(periods.route_of(f), f), f,
                                   "、".join(files) if files else "—"))
    L.append("")
    if result["errors"]:
        L.append("## 路线报错")
        L.append("")
        L.extend("- %s" % e for e in result["errors"])
        L.append("")
    L.append("## 未下载的报表与原因")
    L.extend(download_diag.format_section([{"store": result["store"],
                                            "missing": result["missing"]}])[4:])
    try:
        with io.open(os.path.join(out_dir, "REPORT.md"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(L) + "\n")
    except Exception as e:
        print("  [警告] REPORT.md 写入失败：%s" % str(e)[:80])


def count_files(result):
    return sum(len(v) for v in (result.get("files") or {}).values())


def print_summary(result):
    print("\n" + "=" * 60)
    print("汇总 —— %s（%s）%s" % (result["store"], result["market"], result["month"]))
    print("=" * 60)
    for f in result["expected"]:
        files = result["files"].get(f) or []
        print("  %-36s %s" % (f, "、".join(files) if files else "—"))
    for line in download_diag.format_section([{"store": result["store"],
                                               "missing": result["missing"]}]):
        print(line)


def main(argv=None):
    from core.ziniao_client import ZiniaoClient
    ap = argparse.ArgumentParser(description="one Falabella store, one month")
    ap.add_argument("--store", required=True)
    ap.add_argument("--month", default=None, metavar="YYYY-MM",
                    help="the month the report is about (default: last month)")
    ap.add_argument("--only", nargs="+", choices=ROUTE_NAMES, default=None)
    ap.add_argument("--out", default=None)
    ap.add_argument("--collect-timeout", type=int, default=240)
    ap.add_argument("--keep-open", action="store_true")
    a = ap.parse_args(argv)
    month = a.month or periods.previous_month(datetime.date.today())
    client = ZiniaoClient()
    if not client.is_up():
        client.start(restart=True)
    r = run_store(client, a.store, month, out_dir=a.out, only=a.only,
                  collect_timeout=a.collect_timeout, close_when_done=not a.keep_open)
    print_summary(r)
    bad = [m for m in r["missing"] if download_diag.needs_attention(m)]
    return 1 if (bad or r["errors"]) else 0


if __name__ == "__main__":
    sys.exit(main())
