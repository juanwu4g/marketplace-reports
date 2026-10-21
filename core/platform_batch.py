# -*- coding: utf-8 -*-
"""
The shared batch download runner of every platform except MercadoLibre.

A platform's `platforms/<pkg>/run_batch.py` is a thin module: it builds one
`BatchRunner` with what is genuinely its own and hands `main` to it:

    RUNNER = platform_batch.BatchRunner(
        "FA", default_market="FA-CL",
        forms=forms, periods=periods, run_downloads=run_downloads, site=site,
        log_collected=True)

What a platform declares (the hooks)
    code            two-letter platform code (core/markets.py); names the files
                    logs/batch_<code>_<stamp>.json and logs/pending_<code>.json
    default_market  --market when none is given (e.g. "FA-CL")
    forms           module with Ctx(driver, out_dir, None, range, facts, log=) and
                    collect_owed(ctx, entry) -> [files]
    periods         module with previous_month(date)
    run_downloads   module with ROUTE_NAMES, expected(month, only), run_store(…),
                    print_summary, count_files, read_run_meta, update_run_meta.
                    If it has require_implemented(), a real run calls it before
                    Ziniao is started and stops with exit 4 on NotImplementedError
                    (the scaffold's "还没实现" gate).
    site            module with get(country) -> the site facts forms.Ctx takes
    log_collected   print "  <store> <form> -> [files]" for every owed report the
                    catch-up collected (Falabella does; Amazon and the scaffold
                    did not — kept as it was)
    name_zh/name_en labels; default: the platform's names in core/markets.py

The contract (docs/ARCHITECTURE.md §5, tests/test_platform_contract.py), so
core/run_monthly treats every platform alike:

* which stores: --stores, else config.json batches["<code>-<country>"], else
  every allowed store of that market. Every name goes through the allowlist and
  must be a store of this platform (config.json stores.catalog) — refused before
  Ziniao is asked.
* serial, one browser at a time; nothing one store does ends the batch.
* exit codes: 0 all fine · 1 finished, needs attention · 2 no store reached ok /
  partial · 3 store not allowed · 4 config unreadable or no stores.
* logs/batch_<code>_<stamp>.json summary; the very last stdout line is
  `[BATCH_SUMMARY_JSON] <path>` for the runner.
* reports requested but not ready go to logs/pending_<code>.json and are
  fetched by --collect-only (or the next run).

Nothing here names a platform: codes and labels come from the arguments and
core/markets.py. Paths are read from core.paths.ROOT when used, not at import.
"""
import argparse
import datetime
import io
import json
import os
import time
import traceback

from core import browser_kit as bk
from core import download_diag
from core import markets
from core import paths
from core import pending_store
from core import store_config
from core.store_config import StoreConfigError, StoreNotAllowed


def store_needs_attention(row):
    if row.get("errors"):
        return True
    return any(download_diag.needs_attention(m) for m in row.get("missing") or [])


def write_summary(path, stamp, rows, extra):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = {"started": stamp,
           "stores": [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows]}
    doc.update(extra or {})
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(json.dumps(doc, ensure_ascii=False, indent=2))


def failed_row(name, market, exc, month):
    rec = download_diag.classify(
        dict(download_diag.exception_evidence(exc, step="store-open"), signal="store-open"),
        {"form": "store", "now": time.time()})
    rec["in_report"] = True
    return {"store": name, "market": market, "status": "failed", "files": 0,
            "out_dir": None, "missing": [rec],
            "errors": ["%s: %s" % (exc.__class__.__name__, str(exc)[:300])],
            "pending": 0, "seconds": 0}


class BatchRunner(object):
    """One platform's batch runner. See the module docstring for the hooks."""

    def __init__(self, code, default_market, forms, periods, run_downloads, site,
                 log_collected=False, name_zh=None, name_en=None):
        self.code = code
        self.default_market = default_market
        self.forms = forms
        self.periods = periods
        self.run_downloads = run_downloads
        self.site = site
        self.log_collected = log_collected
        self._name_zh = name_zh
        self._name_en = name_en

    # ------------------------------------------------------------ names and paths

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

    @property
    def log_dir(self):
        return os.path.join(paths.ROOT, "logs")

    @property
    def pending_path(self):
        return os.path.join(self.log_dir, "pending_%s.json" % self.code)

    # ------------------------------------------------------------ stores

    def resolve_stores(self, args):
        if args.stores:
            names = [store_config.require_allowed(n) for n in args.stores]
        else:
            names = store_config.stores_for_market(args.market)
        if not names:
            raise StoreConfigError(
                "没有要跑的 %s 店铺：config.json 里 %s 没有店（stores.catalog / "
                "stores.batches），也没有给 --stores。" % (self.name_zh, args.market))
        store_config.require_platform(names, self.code)
        return names

    # ------------------------------------------------------------ rows and verdict

    def batch_verdict(self, rows):
        failed = [r for r in rows if r.get("status") not in ("ok", "partial")]
        attention = [r for r in rows if r not in failed and store_needs_attention(r)]
        records = [m for r in rows for m in (r.get("missing") or [])]
        if rows and len(failed) == len(rows):
            code = 2
        elif failed or attention:
            code = 1
        else:
            code = 0
        return {"platform": self.code,
                "stores_total": len(rows),
                "stores_ok": len(rows) - len(failed) - len(attention),
                "stores_attention": len(attention),
                "stores_failed": len(failed),
                "missing_by_cause": download_diag.count_by_cause(records),
                "cleaning": "skipped",
                "transform": {"ok": True, "partial": False, "skipped": True,
                              "reason": "platform",
                              "note": "%s 没有清洗这一步" % self.name_zh},
                "exit_code": code}

    def row_of(self, result):
        rd = self.run_downloads
        bad = [m for m in result["missing"] if download_diag.needs_attention(m)]
        status = "partial" if (result["errors"] or bad) else "ok"
        if not rd.count_files(result) and result["errors"]:
            status = "failed"
        return {"store": result["store"], "market": result["market"], "status": status,
                "files": rd.count_files(result), "out_dir": result["out_dir"],
                "missing": result["missing"], "errors": result["errors"],
                "pending": len(result["pending"]), "seconds": result["seconds"],
                "_result": result}

    # ------------------------------------------------------------ owed reports

    def collect_owed(self, client, names, month, log=print):
        """Fetch owed async reports (logs/pending_<code>.json) for these stores.
        Returns {} when nothing is owed, else {"done": n, "still_owed": n}."""
        pending = self.pending_path
        owed = [e for e in pending_store.load(pending)
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
                store_config.check_ziniao_platform(name, store.get("platform_name"), self.code)
                facts = self.site.get(store_config.country_of(name))
                with client.open_store(store) as session:
                    for e in entries:
                        bk.set_download_dir(session.driver, e["out_dir"])
                        ctx = self.forms.Ctx(session.driver, e["out_dir"], None, e["range"],
                                             facts, log=log)
                        try:
                            files = self.forms.collect_owed(ctx, e)
                        except Exception as x:
                            log("  %s %s 补取失败：%s" % (name, e["form"], x))
                            files = []
                        e["probes"] = int(e.get("probes") or 0) + 1
                        e["last_probe_at"] = time.time()
                        if files:
                            if self.log_collected:
                                log("  %s %s -> %s" % (name, e["form"], files))
                            rd = self.run_downloads
                            meta = rd.read_run_meta(e["out_dir"])
                            got = meta.get("files") or {}
                            got[e["form"]] = files
                            missing = [m for m in meta.get("missing") or []
                                       if m.get("form") != e["form"]]
                            rd.update_run_meta(e["out_dir"], files=got, missing=missing)
                            done.append(e)
                        else:
                            keep.append(e)
            except Exception as x:
                log("  %s 打不开，欠账留到下次：%s" % (name, x))
                keep.extend(entries)
        if done:
            pending_store.remove(done, path=pending)
        if keep:
            pending_store.add(keep, path=pending)
        return {"done": len(done), "still_owed": len(keep)}

    # ------------------------------------------------------------ command line

    def parser(self):
        ap = argparse.ArgumentParser(description="%s downloads, several stores" % self.name_en)
        ap.add_argument("--month", default=None, metavar="YYYY-MM")
        ap.add_argument("--stores", nargs="+", default=None)
        ap.add_argument("--market", type=str.upper, default=self.default_market)
        ap.add_argument("--region", type=str.upper, default=None,
                        help="country code; same as --market %s-<code> (runner compatibility)"
                        % self.code)
        ap.add_argument("--only", nargs="+", choices=self.run_downloads.ROUTE_NAMES,
                        default=None)
        ap.add_argument("--collect-only", action="store_true")
        ap.add_argument("--collect-timeout", type=int, default=240)
        ap.add_argument("--dry-run", action="store_true")
        return ap

    def main(self, argv=None):
        rd = self.run_downloads
        a = self.parser().parse_args(argv)
        if a.region:
            a.market = markets.market_id(self.code, a.region)
        try:
            a.market = markets.normalize(a.market)
            if markets.split(a.market)[0] != self.code:
                raise StoreConfigError("--market %s 不是 %s 的市场" % (a.market, self.name_zh))
            names = self.resolve_stores(a)
        except StoreNotAllowed as e:
            print("[错误] %s" % e)
            return 3
        except (StoreConfigError, markets.UnknownMarket) as e:
            print("[错误] %s" % e)
            return 4
        month = a.month or self.periods.previous_month(datetime.date.today())
        print("%s 批量下载：%s，%s，%d 家店：%s"
              % (self.name_zh, markets.label(a.market), month, len(names), "、".join(names)))
        if a.dry_run:
            for n in names:
                print("  %s -> %s" % (n, ", ".join(rd.expected(month, a.only))))
            return 0
        require = getattr(rd, "require_implemented", None)
        if require is not None:
            try:
                require()
            except NotImplementedError as e:
                print("[未实现] %s" % e)
                return 4

        from core.ziniao_client import ZiniaoClient
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(self.log_dir, "batch_%s_%s.json" % (self.code, stamp))
        rows = []
        try:
            client = ZiniaoClient()
            if not client.is_up():
                client.start(restart=True)
        except Exception as e:
            print("[错误] 紫鸟客户端起不来：%s" % e)
            rows = [failed_row(n, a.market, e, month) for n in names]
            write_summary(log_path, stamp, rows, self.batch_verdict(rows))
            print("[BATCH_SUMMARY_JSON] %s" % os.path.abspath(log_path))
            return 2

        if a.collect_only:
            late = self.collect_owed(client, names, month)
            verdict = {"platform": self.code, "collect_only": True, "late": late,
                       "exit_code": 1 if late.get("still_owed") else 0}
            write_summary(log_path, stamp, [], verdict)
            print("[BATCH_SUMMARY_JSON] %s" % os.path.abspath(log_path))
            return verdict["exit_code"]

        for i, name in enumerate(names, 1):
            print("\n" + "#" * 60)
            print("# [%d/%d] %s" % (i, len(names), name))
            print("#" * 60)
            try:
                result = rd.run_store(client, name, month, only=a.only,
                                      collect_timeout=a.collect_timeout)
                rd.print_summary(result)
                if result["pending"]:
                    pending_store.add(result["pending"], path=self.pending_path)
                rows.append(self.row_of(result))
            except StoreNotAllowed as e:
                print("[错误] %s" % e)
                return 3
            except Exception as e:
                traceback.print_exc()
                rows.append(failed_row(name, a.market, e, month))
            write_summary(log_path, stamp, rows, self.batch_verdict(rows))

        late = (self.collect_owed(client, names, month)
                if any(r.get("pending") for r in rows) else {})
        verdict = self.batch_verdict(rows)
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
