# -*- coding: utf-8 -*-
"""
Amazon downloads for several stores in one Ziniao session.

    python -m platforms.amazon.run_batch --month YYYY-MM --market AZ-MX --stores "<store>"
    python -m platforms.amazon.run_batch --collect-only --month YYYY-MM --stores ...
    python -m platforms.amazon.run_batch --dry-run --market AZ-MX

The shared contract (docs/ARCHITECTURE.md §5, tests/test_platform_contract.py),
so core/run_monthly treats every platform alike:

* which stores: --stores, else config.json batches["AZ-<country>"], else
  every allowed store of that market. Every name goes through the allowlist and
  must be a store of this platform (config.json stores.catalog) — refused before
  Ziniao is asked.
* serial, one browser at a time; nothing one store does ends the batch.
* exit codes: 0 all fine · 1 finished, needs attention · 2 no store reached ok /
  partial · 3 store not allowed · 4 config unreadable or no stores.
* logs/batch_AZ_<stamp>.json summary; the very last stdout line is
  `[BATCH_SUMMARY_JSON] <path>` for the runner.
* reports requested but not ready go to logs/pending_AZ.json and are
  fetched by --collect-only (or the next run).
"""
import argparse
import datetime
import io
import json
import os
import sys
import time
import traceback

from core import browser_kit as bk
from core import console  # noqa: F401  中文输出编码保护
from core import download_diag
from core import markets
from core import pending_store
from core import store_config
from core.paths import ROOT
from core.store_config import StoreConfigError, StoreNotAllowed
from platforms.amazon import forms
from platforms.amazon import periods
from platforms.amazon import run_downloads
from platforms.amazon import site as platform_site

LOG_DIR = os.path.join(ROOT, "logs")
PENDING = os.path.join(LOG_DIR, "pending_AZ.json")
PLATFORM = "AZ"
DEFAULT_MARKET = "AZ-MX"


def resolve_stores(args):
    if args.stores:
        names = [store_config.require_allowed(n) for n in args.stores]
    else:
        names = store_config.stores_for_market(args.market)
    if not names:
        raise StoreConfigError(
            "没有要跑的 亚马逊 店铺：config.json 里 %s 没有店（stores.catalog / "
            "stores.batches），也没有给 --stores。" % args.market)
    store_config.require_platform(names, PLATFORM)
    return names


def store_needs_attention(row):
    if row.get("errors"):
        return True
    return any(download_diag.needs_attention(m) for m in row.get("missing") or [])


def batch_verdict(rows):
    failed = [r for r in rows if r.get("status") not in ("ok", "partial")]
    attention = [r for r in rows if r not in failed and store_needs_attention(r)]
    records = [m for r in rows for m in (r.get("missing") or [])]
    if rows and len(failed) == len(rows):
        code = 2
    elif failed or attention:
        code = 1
    else:
        code = 0
    return {"platform": PLATFORM,
            "stores_total": len(rows),
            "stores_ok": len(rows) - len(failed) - len(attention),
            "stores_attention": len(attention),
            "stores_failed": len(failed),
            "missing_by_cause": download_diag.count_by_cause(records),
            "cleaning": "skipped",
            "transform": {"ok": True, "partial": False, "skipped": True,
                          "reason": "platform", "note": "亚马逊 没有清洗这一步"},
            "exit_code": code}


def write_summary(path, stamp, rows, extra):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = {"started": stamp,
           "stores": [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows]}
    doc.update(extra or {})
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(json.dumps(doc, ensure_ascii=False, indent=2))


def row_of(result):
    bad = [m for m in result["missing"] if download_diag.needs_attention(m)]
    status = "partial" if (result["errors"] or bad) else "ok"
    if not run_downloads.count_files(result) and result["errors"]:
        status = "failed"
    return {"store": result["store"], "market": result["market"], "status": status,
            "files": run_downloads.count_files(result), "out_dir": result["out_dir"],
            "missing": result["missing"], "errors": result["errors"],
            "pending": len(result["pending"]), "seconds": result["seconds"],
            "_result": result}


def failed_row(name, market, exc, month):
    rec = download_diag.classify(
        dict(download_diag.exception_evidence(exc, step="store-open"), signal="store-open"),
        {"form": "store", "now": time.time()})
    rec["in_report"] = True
    return {"store": name, "market": market, "status": "failed", "files": 0,
            "out_dir": None, "missing": [rec],
            "errors": ["%s: %s" % (exc.__class__.__name__, str(exc)[:300])],
            "pending": 0, "seconds": 0}


def collect_owed(client, names, month, log=print):
    """Fetch owed async reports (logs/pending_AZ.json) for these stores."""
    owed = [e for e in pending_store.load(PENDING)
            if e.get("store") in names and (not month or e.get("range") == month)]
    if not owed:
        return {}
    log("\n补取欠账：%d 条" % len(owed))
    done, keep = [], []
    by_store = {}
    for e in owed:
        by_store.setdefault(e["store"], []).append(e)
    for name, entries in by_store.items():
        try:
            store = client.find_store(name)
            store_config.check_ziniao_platform(name, store.get("platform_name"), PLATFORM)
            facts = platform_site.get(store_config.country_of(name))
            with client.open_store(store) as session:
                for e in entries:
                    bk.set_download_dir(session.driver, e["out_dir"])
                    ctx = forms.Ctx(session.driver, e["out_dir"], None, e["range"], facts,
                                    log=log)
                    try:
                        files = forms.collect_owed(ctx, e)
                    except Exception as x:
                        log("  %s %s 补取失败：%s" % (name, e["form"], x))
                        files = []
                    e["probes"] = int(e.get("probes") or 0) + 1
                    e["last_probe_at"] = time.time()
                    if files:
                        meta = run_downloads.read_run_meta(e["out_dir"])
                        got = meta.get("files") or {}
                        got[e["form"]] = files
                        missing = [m for m in meta.get("missing") or []
                                   if m.get("form") != e["form"]]
                        run_downloads.update_run_meta(e["out_dir"], files=got, missing=missing)
                        done.append(e)
                    else:
                        keep.append(e)
        except Exception as x:
            log("  %s 打不开，欠账留到下次：%s" % (name, x))
            keep.extend(entries)
    if done:
        pending_store.remove(done, path=PENDING)
    if keep:
        pending_store.add(keep, path=PENDING)
    return {"done": len(done), "still_owed": len(keep)}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Amazon downloads, several stores")
    ap.add_argument("--month", default=None, metavar="YYYY-MM")
    ap.add_argument("--stores", nargs="+", default=None)
    ap.add_argument("--market", type=str.upper, default=DEFAULT_MARKET)
    ap.add_argument("--region", type=str.upper, default=None,
                    help="country code; same as --market AZ-<code> (runner compatibility)")
    ap.add_argument("--only", nargs="+", choices=run_downloads.ROUTE_NAMES, default=None)
    ap.add_argument("--collect-only", action="store_true")
    ap.add_argument("--collect-timeout", type=int, default=240)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args(argv)
    if a.region:
        a.market = markets.market_id(PLATFORM, a.region)
    try:
        a.market = markets.normalize(a.market)
        if markets.split(a.market)[0] != PLATFORM:
            raise StoreConfigError("--market %s 不是 亚马逊 的市场" % a.market)
        names = resolve_stores(a)
    except StoreNotAllowed as e:
        print("[错误] %s" % e)
        return 3
    except (StoreConfigError, markets.UnknownMarket) as e:
        print("[错误] %s" % e)
        return 4
    month = a.month or periods.previous_month(datetime.date.today())
    print("亚马逊 批量下载：%s，%s，%d 家店：%s"
          % (markets.label(a.market), month, len(names), "、".join(names)))
    if a.dry_run:
        for n in names:
            print("  %s -> %s" % (n, ", ".join(run_downloads.expected(month, a.only))))
        return 0
    try:
        run_downloads.require_implemented()
    except NotImplementedError as e:
        print("[未实现] %s" % e)
        return 4

    from core.ziniao_client import ZiniaoClient
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(LOG_DIR, "batch_%s_%s.json" % (PLATFORM, stamp))
    rows = []
    try:
        client = ZiniaoClient()
        if not client.is_up():
            client.start(restart=True)
    except Exception as e:
        print("[错误] 紫鸟客户端起不来：%s" % e)
        rows = [failed_row(n, a.market, e, month) for n in names]
        write_summary(log_path, stamp, rows, batch_verdict(rows))
        print("[BATCH_SUMMARY_JSON] %s" % os.path.abspath(log_path))
        return 2

    if a.collect_only:
        late = collect_owed(client, names, month)
        verdict = {"platform": PLATFORM, "collect_only": True, "late": late,
                   "exit_code": 1 if late.get("still_owed") else 0}
        write_summary(log_path, stamp, [], verdict)
        print("[BATCH_SUMMARY_JSON] %s" % os.path.abspath(log_path))
        return verdict["exit_code"]

    for i, name in enumerate(names, 1):
        print("\n" + "#" * 60)
        print("# [%d/%d] %s" % (i, len(names), name))
        print("#" * 60)
        try:
            result = run_downloads.run_store(client, name, month, only=a.only,
                                             collect_timeout=a.collect_timeout)
            run_downloads.print_summary(result)
            if result["pending"]:
                pending_store.add(result["pending"], path=PENDING)
            rows.append(row_of(result))
        except StoreNotAllowed as e:
            print("[错误] %s" % e)
            return 3
        except Exception as e:
            traceback.print_exc()
            rows.append(failed_row(name, a.market, e, month))
        write_summary(log_path, stamp, rows, batch_verdict(rows))

    late = collect_owed(client, names, month) if any(r.get("pending") for r in rows) else {}
    verdict = batch_verdict(rows)
    if late:
        verdict["late"] = late
    write_summary(log_path, stamp, rows, verdict)
    print("\n" + "=" * 60)
    print("批次结论：%d 家店，正常 %d，要看 %d，失败 %d"
          % (verdict["stores_total"], verdict["stores_ok"], verdict["stores_attention"],
             verdict["stores_failed"]))
    for line in download_diag.format_section(rows):
        print(line)
    print("[BATCH_SUMMARY_JSON] %s" % os.path.abspath(log_path))
    return verdict["exit_code"]


if __name__ == "__main__":
    sys.exit(main())
