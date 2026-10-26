# -*- coding: utf-8 -*-
"""
One Liverpool store, one month: open the store (the supplier-portal user), sign
in to the portal, fetch its forms; then open the store's marketplace user (a second
Ziniao store, config.json stores.catalog[store].marketplace_store) and fetch the
marketplace forms into the same folder; explain anything missing.

    python -m platforms.liverpool.run_downloads --store "<store>" --month YYYY-MM

Output: layout.run_folder(store, "<YYYYMMDD_HHMMSS>_m<YYYYMM>") (core/layout.py) with the raw files,
run_meta.json (market, month, expected / missing / flags) and REPORT.md. Same
layout and keys as the other platforms, so the runner, the window and the
report side read all of them the same way.
"""
import argparse
import datetime
import io
import json
import os
import sys
import time

from core import browser_kit as bk
from core import console  # noqa: F401  中文输出编码保护
from core import download_diag
from core import layout
from core import site_time
from core import store_config
from core.paths import ROOT
from platforms.liverpool import forms
from platforms.liverpool import periods
from platforms.liverpool import site as platform_site

PLATFORM = "LV"
ROUTE_NAMES = tuple(periods.ROUTES)
# False until the routes in forms.py are written: a real run stops before Ziniao
# is touched, with a message that says so.
IMPLEMENTED = True


def require_implemented():
    if not IMPLEMENTED:
        raise forms.NotBuiltYet("利物浦 的下载还没实现（platforms/liverpool/forms.py），"
                                "现在只能 --dry-run")


def run_folder(store_name, month, stamp=None, base=None):
    """Absolute folder (CDP resolves a relative download path against the
    browser's own working directory, and files would vanish)."""
    stamp = stamp or datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    y, m = periods.parse_month(month)
    name = "%s_m%04d%02d" % (stamp, y, m)
    if base:
        return os.path.abspath(os.path.join(base, store_name, name))
    return layout.run_folder(store_name, name)


def expected(month, only=None):
    forms_ = periods.form_ids(month)
    if only:
        forms_ = [f for f in forms_ if periods.route_of(f) in set(only)]
    return forms_


# the catalog option naming the store's marketplace user (a second Ziniao store)
MARKETPLACE_OPTION = "marketplace_store"


def _site(client, ziniao_name, owner, routes, exp, diag, month, facts, out_dir, opener, host,
          result, close_when_done, collect_timeout, log):
    """Open one Ziniao store, sign in to its site (opener), run its routes. Every
    store-level failure gives up the site's remaining forms with the page as evidence."""
    mine = [f for f in exp if periods.route_of(f) in routes]
    if not mine:
        return
    store = client.find_store(ziniao_name)
    log("店铺：%s（%s）" % (store.get("browserName"), store.get("platform_name")))
    store_config.check_ziniao_platform(owner, store.get("platform_name"), PLATFORM)
    with client.open_store(store, close_when_done=close_when_done) as session:
        d = session.driver
        bk.set_download_dir(d, out_dir)
        ctx = forms.Ctx(d, out_dir, diag, month, facts, today=site_time.site_today(),
                        collect_timeout=collect_timeout, log=log, host=host)
        try:
            opener(ctx)
            for route in [r for r in ROUTE_NAMES if r in routes]:
                forms_ = [f for f in mine if periods.route_of(f) == route]
                if forms_:
                    forms.guarded(ctx, route, forms_, forms.ROUTES[route])
        except Exception as e:
            # a sign-in page, a wrong language, a dead browser: the rest of this site
            what = {"NeedsLogin": "login", "Blocked": "blocked", "LanguageWrong": "language"}.get(
                e.__class__.__name__, "store")
            ctx.errors.append("%s %s: %s: %s" % (ziniao_name, what, e.__class__.__name__, str(e)[:200]))
            log("  [%s] %s" % (ziniao_name, e))
            ev = download_diag.exception_evidence(e, step=what, driver=d)
            for f in mine:
                if f not in ctx.files and f not in ctx.no_data:
                    diag.items.setdefault(f, {"evidence": dict(ev), "period": None, "required": True})
        result["files"].update(ctx.files)
        result["errors"].extend(ctx.errors)
        result["flags"].update(ctx.flags)
        result["no_data"].update(ctx.no_data)


def run_store(client, store_name, month, out_dir=None, only=None, collect_timeout=240,
              close_when_done=True, log=print):
    """The store is the supplier-portal user; its marketplace user (config.json
    stores.catalog[store].marketplace_store) is opened after it, into the same folder."""
    require_implemented()
    store_config.require_allowed(store_name)
    store_config.require_platform([store_name], PLATFORM)
    country = store_config.country_of(store_name)
    facts = platform_site.get(country)
    site_time.set_site(facts["tz"], facts["utc_offset"])
    market_user = store_config.companion_store(store_name, MARKETPLACE_OPTION)
    out_dir = out_dir or run_folder(store_name, month)
    os.makedirs(out_dir, exist_ok=True)
    log("输出目录：%s" % out_dir)
    began = time.time()
    exp = expected(month, only)
    result = {"store": store_name, "platform": PLATFORM,
              "market": store_config.market_of(store_name), "country": country,
              "month": month, "out_dir": out_dir, "files": {}, "errors": [],
              "pending": [], "expected": exp, "missing": [],
              "flags": {"marketplace_store": market_user}, "no_data": {},
              "cautions": [], "seconds": 0}
    diag = download_diag.Collector(out_dir, store_name)

    _site(client, store_name, store_name, periods.PORTAL_ROUTES, exp, diag, month, facts, out_dir,
          lambda ctx: (forms.portal_login(ctx), forms.check_language(ctx)), facts["portal"],
          result, close_when_done, collect_timeout, log)
    if market_user:
        _site(client, market_user, store_name, periods.MARKETPLACE_ROUTES, exp, diag, month, facts,
              out_dir, forms.marketplace_open, facts["host"], result, close_when_done, collect_timeout, log)
    else:
        why = ("config.json stores.catalog[%r] 没有写 %s（这家店的 Marketplace 子账号），"
               "订单、物流费明细、补充物流费没法下载" % (store_name, MARKETPLACE_OPTION))
        log("  [配置] %s" % why)
        result["errors"].append("config: %s" % why)
        for f in exp:
            if periods.route_of(f) in periods.MARKETPLACE_ROUTES:
                diag.items.setdefault(f, {"evidence": {"signal": "exception", "step": "config",
                                                        "exception": why,
                                                        "exception_class": "StoreConfigError"},
                                          "period": None, "required": True})

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


# route -> Chinese name of the form, for REPORT.md
FORM_ZH = {"ledger": "账本（Cobros y Pagos Marketplace）", "sales": "销售单据（Ventas Marketplace）",
           "payments": "付款（Detalle de Pagos）", "invoices": "平台费用发票（Notas de Cargo）",
           "orders": "订单列表", "logistics": "物流费明细（Costos logísticos）",
           "balance": "补充物流费（Balance de cuenta）"}


def write_report(out_dir, result):
    L = ["# %s · 利物浦 %s · %s" % (result["store"], result["country"], result["month"]), "",
         "下载时间：%s，用时 %d 秒" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                                result.get("seconds") or 0), "", "## 文件", ""]
    for f in result["expected"]:
        files = result["files"].get(f) or []
        L.append("- %s `%s`：%s" % (FORM_ZH.get(periods.route_of(f), f), f,
                                   "、".join(files) if files else "—"))
    L.append("")
    if result["errors"]:
        L += ["## 路线报错", ""] + ["- %s" % e for e in result["errors"]] + [""]
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
    ap = argparse.ArgumentParser(description="one Liverpool store, one month")
    ap.add_argument("--store", required=True)
    ap.add_argument("--month", default=None, metavar="YYYY-MM",
                    help="the month the report is about (default: last month)")
    ap.add_argument("--only", nargs="+", choices=ROUTE_NAMES, default=None)
    ap.add_argument("--out", default=None)
    ap.add_argument("--collect-timeout", type=int, default=240)
    ap.add_argument("--keep-open", action="store_true")
    a = ap.parse_args(argv)
    require_implemented()
    from core.ziniao_client import ZiniaoClient
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
