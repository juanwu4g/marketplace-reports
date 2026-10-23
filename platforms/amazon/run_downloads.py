# -*- coding: utf-8 -*-
"""
One Amazon store, one month: open the store, check the language, fetch the
forms, explain anything missing.

    python -m platforms.amazon.run_downloads --store "<store>" --month YYYY-MM

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
from core import site_time
from core import store_config
from core import layout
from core.paths import ROOT
from platforms.amazon import forms
from platforms.amazon import periods
from platforms.amazon import site as platform_site

PLATFORM = "AZ"
ROUTE_NAMES = tuple(periods.ROUTES)
# False until the routes in forms.py are written: a real run stops before Ziniao
# is touched, with a message that says so.
IMPLEMENTED = True


def require_implemented():
    if not IMPLEMENTED:
        raise forms.NotBuiltYet("亚马逊 的下载还没实现（platforms/amazon/forms.py），"
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


def run_store(client, store_name, month, out_dir=None, only=None, collect_timeout=240,
              close_when_done=True, log=print):
    require_implemented()
    store_config.require_allowed(store_name)
    store_config.require_platform([store_name], PLATFORM)
    country = store_config.country_of(store_name)
    facts = platform_site.get(country)
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
    diag = download_diag.Collector(out_dir, store_name)

    store = client.find_store(store_name)
    log("店铺：%s（%s）" % (store.get("browserName"), store.get("platform_name")))
    store_config.check_ziniao_platform(store_name, store.get("platform_name"), PLATFORM)

    with client.open_store(store, close_when_done=close_when_done) as session:
        d = session.driver
        bk.set_download_dir(d, out_dir)
        ctx = forms.Ctx(d, out_dir, diag, month, facts, today=site_time.site_today(),
                        collect_timeout=collect_timeout, log=log)
        try:
            forms.check_language(ctx)
            log("界面语言 %s" % ctx.flags.get("locale"))
            want = [r for r in ROUTE_NAMES if any(periods.route_of(f) == r for f in exp)]
            for route in want:
                mine = [f for f in exp if periods.route_of(f) == route]
                forms.guarded(ctx, route, mine, forms.ROUTES[route])
        except forms.Blocked as e:
            # sign-in / verification / captcha: a person has to look (never solved here)
            ctx.errors.append("blocked: %s" % e)
            for f in exp:
                if f not in ctx.files and f not in diag.items:
                    ctx.give_up(f, "exception", step="blocked", exception=str(e)[:200],
                                exception_class="Blocked")
        except forms.LanguageWrong as e:
            ctx.errors.append("language: %s" % e)
            ev = download_diag.exception_evidence(e, step="language", driver=d)
            for f in exp:
                diag.items.setdefault(f, {"evidence": dict(ev), "period": None,
                                          "required": True})
        except Exception as e:
            # the browser died or the first page never rendered: the whole store
            ctx.errors.append("store: %s: %s" % (e.__class__.__name__, str(e)[:200]))
            ev = download_diag.exception_evidence(e, step="store", driver=d)
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


# route -> Chinese name of the form, for REPORT.md
FORM_ZH = {"order": "交易明细（Transacciones 月报 CSV）", "summary": "月度汇总（Resumen PDF）",
           "invoice": "平台费用发票（CFDI XML + PDF）", "stock": "FBA 库存快照",
           "taxcert": "月度预扣税证明（Certificado de retención）"}


def write_report(out_dir, result):
    L = ["# %s · 亚马逊 %s · %s" % (result["store"], result["country"], result["month"]), "",
         "下载时间：%s，用时 %d 秒" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                                result.get("seconds") or 0), ""]
    sw = result["flags"].get("language_switched")
    if sw:
        L += ["⚠ 开始时界面语言是 %s，脚本已把它改回 %s。有人在用这家店时改了语言；"
              "反复出现的话，请提醒用这家店的人保持西班牙语。" % (sw.get("from"), sw.get("to")), ""]
    owed = [p for p in result.get("pending") or []]
    if owed:
        L += ["## 欠账（之后补取）", ""] + ["- %s：%s" % (FORM_ZH.get(p["form"], p["form"]),
                                                    p.get("note") or p.get("last_state"))
                                          for p in owed] + [""]
    L += ["## 文件", ""]
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
    ap = argparse.ArgumentParser(description="one Amazon store, one month")
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
