# -*- coding: utf-8 -*-
"""
Run the download pipeline over several stores in one client session.

    python run_batch.py                          # every store in stores.batch
    python run_batch.py --stores BOCINA_SM EWTTO_SM
    python run_batch.py --only ventas --limit 2
    python run_batch.py --dry-run                # just show what would run

Which stores
------------
`--stores` if given, else `stores.batch` from config.json, else the whole
`stores.allowed` list. Every name still goes through the allowlist, so a typo
in `--stores` is rejected locally without a request reaching Ziniao.

If the allowlist is empty (= unrestricted) there is nothing to iterate and the
batch refuses to start. "Every store on the account" is 93 stores and is never
what anyone means.

Why serial
----------
One browser at a time. Whether Ziniao will hold two stores open at once is
untested, and each store has its own proxy and fingerprint - opening several in
parallel multiplies both load and the chance of tripping a verification
challenge. The client itself starts ONCE and is reused, so the per-store cost
is just startBrowser/stopBrowser, not a full client restart.

Failure isolation
-----------------
Nothing one store does can end the batch. A route that breaks is recorded by
run_store(); a store that cannot even be opened is caught here. Both show up in
the summary and in the exit code:

    0   every store ok and nothing expected is missing
    1   finished, needs attention: some store partial/failed, or reports still
        owed, or cleaning partially failed. A store "needs attention" when it
        has route errors, or is missing any expected report for a reason other
        than NO_DATA / NOT_AVAILABLE (those two mean the platform has nothing
        to give - see download_diag.CAUSES).
    2   hard failure: Ziniao client error, or NO store at all reached
        ok/partial, so nothing fresh was downloaded
    3   a store is not in the allowlist
    4   config.json is unreadable, or the batch list is empty

Why something is missing
------------------------
After the late pass every report that was expected and did not arrive gets a
cause (download_diag): printed under "未下载的报表与原因", written to each
store's REPORT.md and run_meta.json, and to logs/batch_<stamp>.json. The very
last stdout line is `[BATCH_SUMMARY_JSON] <path>` for the runner.
"""
import argparse
import datetime
import io
import json
import os
import sys
import time
import traceback

from ziniao_client import ZiniaoClient, ZiniaoError, StoreNotAllowed
from store_config import StoreConfigError
import console  # noqa: F401  中文输出编码保护
import store_config
import run_downloads
import transform
import pending_store
import meli_forms
import run_report
import download_diag

ROOT = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(ROOT, "logs")
NEWLINE = chr(10)


def resolve_stores(args):
    """The store list for this batch, already checked against the allowlist."""
    if args.stores:
        names = [store_config.require_allowed(n) for n in args.stores]
    else:
        names = store_config.batch_stores()
    if not names:
        raise StoreConfigError(
            "no stores to run: config.json -> stores.allowed is empty "
            "(unrestricted), so there is no list to walk. Name the stores in "
            "stores.allowed, or pass --stores.")
    if args.limit:
        names = names[:args.limit]
    return names


def write_summary(path, stamp, rows, transform_result=None, extra=None):
    """Machine-readable record of the run.

    A scheduled batch throws stdout away, so without this a silent failure is
    invisible until someone notices missing data.

    `extra` 是 batch_verdict() 的结果：按原因的计数、各状态的店铺数、清洗状态、
    退出码。行里以下划线开头的键是进程内的东西（完整的 result、补收的中间
    结果），不写进文件。
    """
    folder = os.path.dirname(path) or LOG_DIR
    if not os.path.isdir(folder):
        os.makedirs(folder)
    doc = {"started": stamp,
           "stores": [{k: v for k, v in r.items() if not k.startswith("_")}
                      for r in rows]}
    if transform_result is not None:
        doc["transform"] = transform_result
    doc.update(extra or {})
    # utf-8 explicitly: store names are Chinese on some accounts and the
    # default Windows encoding raises UnicodeEncodeError (GUIDE.md 6.4).
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(json.dumps(doc, ensure_ascii=False, indent=2))


def rel_to_downloads(out_dir, name):
    """A downloaded file as a path relative to downloads/.

    That is what the cleaner's matcher wants. A bare filename also matches, but
    against every copy of that name in the tree.
    """
    return os.path.relpath(os.path.join(out_dir, os.path.basename(name)),
                           transform.PIPELINE_ROOT).replace(chr(92), "/")


def make_workbook(row, stamp, args):
    """Write this store's accounting workbook, now that its ventas is in hand.

    Per store, as soon as that store is done - not batched to the end. A store
    that finishes at 09:10 has a usable workbook at 09:10 rather than after the
    twelfth store, which matters when a run is interrupted: work already done is
    already delivered.

    `validate` reads the RAW export and pairs it with the facturacion files in
    the SAME run folder, so it needs neither the Parquet build nor any other
    store. That independence is what makes per-store generation correct rather
    than merely convenient.
    """
    if row.get("workbook") or not row.get("ventas"):
        return
    out = os.path.join(transform.PIPELINE_ROOT, "data", "reports",
                       "%s_%s.xlsx" % (row["store"], stamp))
    res = transform.validate(row["ventas"][0], out)
    if res["ok"]:
        row["workbook"] = out
    else:
        row["errors"] = list(row.get("errors", [])) + [
            "workbook: %s" % (res["output"] or "").strip()[-200:]]
        if row["status"] == "ok":
            row["status"] = "partial"


def owed_label(entry, now=None):
    """还欠着的一条，汇总里怎么写。纯函数。

    以前这里一律写成 "mercadopago/%s (%s) - still generating"：补收阶段剩下的
    Ventas 于是变成 `mercadopago/None (None) - still generating`，而"还在生成"
    也只是猜的 —— 条目上明明留着最后一次探测看到了什么。
    """
    rec = download_diag.record_from_entry(entry, now=now)
    return "%s%s - [%s] %s" % (
        rec["form"], " (%s)" % rec["period"] if rec["period"] else "",
        rec["cause"], download_diag.CAUSES[rec["cause"]][0])


def cleaning_state(tr):
    """清洗这一步的状态，单独一个字段。None = 这次没有这一步。"""
    if tr is None:
        return None
    if tr.get("skipped"):
        return "skipped"
    if tr.get("ok"):
        return "ok"
    return "partial" if tr.get("partial") else "failed"


def store_needs_attention(row):
    """这家店要不要人看一眼。纯函数。

    平台不提供 / 该区间没有数据（NO_DATA、NOT_AVAILABLE）不算：缺了是对的。
    """
    if row.get("status") != "ok":
        return True
    if row.get("pending"):
        return True
    return any(download_diag.needs_attention(r) for r in row.get("missing") or [])


def batch_verdict(rows, tr=None, expired=()):
    """批次的结论：各状态的店铺数、按原因的计数、清洗状态、退出码。纯函数。

    退出码（和 run_monthly / run_reports 约定好的）：
        0  每家店都 ok，应有的报表都在
        1  跑完了但要人看：有店 partial/failed、有报表还欠着、或清洗部分失败
        2  没有任何一家店到达 ok/partial —— 这一趟没有下到任何新数据

    清洗失败**不算店铺失败**：下载都在盘上，坏的是后一步。它在 cleaning 字段
    里单独报，退出码仍是 1。
    """
    failed = [r for r in rows if r.get("status") not in ("ok", "partial")]
    attention = [r for r in rows if r not in failed and store_needs_attention(r)]
    records = [m for r in rows for m in (r.get("missing") or [])]
    cleaning = cleaning_state(tr)
    if rows and len(failed) == len(rows):
        code = 2
    elif failed or attention or cleaning in ("partial", "failed"):
        code = 1
    else:
        code = 0
    return {"stores_total": len(rows),
            "stores_ok": len(rows) - len(failed) - len(attention),
            "stores_attention": len(attention),
            "stores_failed": len(failed),
            "missing_by_cause": download_diag.count_by_cause(
                records + list(expired or [])),
            "cleaning": cleaning,
            "expired": list(expired or []),
            "exit_code": code}


def refresh_folder(folder, live, no_data, row=None, late=None, args=None,
                   expired=()):
    """补收之后，把一个下载目录的最终结论写回它的 run_meta.json 和 REPORT.md。

    返回这个目录的最终缺失记录。`row` 只在这是本批次刚跑的那个目录时给：那时
    手里有完整的 result，整份报告重写（补收回来的文件也进文件表）；否则
    （更早那一趟的目录、--collect-only）只换掉"未下载 —— 原因"那一段。

    用的全是补收时已经采到的证据，不为了诊断重开店铺。
    """
    result = row.get("_result") if row else None
    usable = bool(folder) and os.path.isdir(folder)
    meta = run_downloads.read_run_meta(folder) if usable else {}
    previous = result.get("missing") if result is not None else meta.get("missing")
    missing = download_diag.merge_final(previous, live, no_data)
    missing.extend(r for r in expired
                   if not any(m["form"] == r["form"] and m["period"] == r["period"]
                              for m in missing))
    got = (late or {}).get("by_dir", {}).get(folder) or {}
    late_files = list(got.get("sales") or []) + list(got.get("mp") or [])
    if not usable or (missing == (previous or []) and not late_files):
        return missing
    if result is not None:
        result["missing"] = missing
        for key in ("sales", "mp"):
            result[key] = list(result.get(key) or []) + [
                f for f in got.get(key) or [] if f not in (result.get(key) or [])]
        if late_files:
            result["late_files"] = late_files
        result["pending"] = list(row.get("pending") or [])
        result["empty"] = list(row.get("empty") or [])
        result["errors"] = list(row.get("errors") or [])
        run_downloads.update_run_meta(folder, missing=missing)
        run_report.write(result, args=args, started=row.get("_started"))
    else:
        run_downloads.update_run_meta(folder, missing=missing)
        run_report.update_missing(folder, missing, cautions=meta.get("cautions"),
                                  flags=meta.get("flags"), late_files=late_files)
    return missing


def same_dir(a, b):
    return bool(a) and bool(b) and (os.path.normcase(os.path.normpath(a))
                                    == os.path.normcase(os.path.normpath(b)))


def newest_run(folders):
    """这家店磁盘上最新的那一趟：(目录, 它应有的报表)。找不到返回 (None, None)。

    目录名以时间戳开头，按名字排序就是按时间排序。只认写了 expected 的目录
    （这个版本之前的目录没有，没法拿来判断"是不是已经下到了"）。
    """
    parents = download_diag.dedupe(
        os.path.dirname(os.path.normpath(f)) for f in folders if f)
    best = None
    for parent in parents:
        try:
            names = [n for n in os.listdir(parent)
                     if os.path.isdir(os.path.join(parent, n)) and n[:8].isdigit()]
        except OSError:
            continue
        for n in names:
            if best is None or n > os.path.basename(best):
                best = os.path.join(parent, n)
    if not best:
        return None, None
    expected = (run_downloads.read_run_meta(best) or {}).get("expected")
    return (best, expected) if expected else (None, None)


def settle_carried(row, recs, carried, live, ref, expected):
    """更早几趟留下的欠账，和这一趟的结果对一下。就地改 row / carried。

    * 这一趟已经下到了同一种报表 → 旧条目作废（superseded），从登记簿里撤掉。
      出报表只用最新那个目录；留着它，之后每一趟的补收都要为它再开一次店、
      白等几分钟，直到 7 天后过期。
    * 这一趟也没下到 → 这一趟自己有一条记录，旧的只是重复，不再另算。

    所以旧条目**从不**影响这家店这一趟的状态；row["pending"] 也按这一趟自己
    的记录重写，免得里面混着别的目录的条目。
    """
    lacking = set(r.get("form") for r in recs)
    delivered = set(expected or []) - lacking
    moot = [e for e in live
            if e.get("out_dir") and not same_dir(e.get("out_dir"), ref)
            and download_diag.form_of_entry(e) in delivered]
    if moot:
        pending_store.remove(moot)
        print("    %-16s 最新一趟已下到 %s，撤销更早目录里的 %d 条旧欠账"
              % (row.get("store"),
                 "、".join(sorted(set(download_diag.form_of_entry(e)
                                     for e in moot))), len(moot)))
    for r in carried:
        r["superseded"] = r.get("form") in delivered
    row["pending"] = [
        "%s%s - [%s] %s" % (r["form"],
                            " (%s)" % r["period"] if r.get("period") else "",
                            r["cause"], r.get("explain") or "")
        for r in recs if r.get("owed")]


def finalize_diagnosis(rows, args=None, expired=()):
    """批次收尾的诊断这一步：给每张没到的报表下最终结论。

    在补收之后跑。逐店把登记簿里现在还欠着的条目（带着补收时最后一次探测的
    证据）和店内诊断的结果合起来，写回各自的目录，并放进 row["missing"]。
    """
    everything = pending_store.load()
    for row in rows:
        name = row["store"]
        late = row.get("_late") or {}
        live = [e for e in everything if e.get("store") == name]
        no_data = list(late.get("no_data") or [])
        folders = download_diag.dedupe(
            [row.get("out_dir")]
            + [e.get("out_dir") for e in live + no_data]
            + list((late.get("by_dir") or {}).keys()))
        # 店铺根本没打开时没有目录，它的结论在主循环里已经下好了
        # 以哪个目录为准：这一趟刚跑的那个；--collect-only 没有自己的目录，就以
        # 这家店磁盘上最新的那一趟为准（出报表用的也是它）。
        ref, ref_expected = row.get("out_dir"), row.get("expected")
        if not ref:
            ref, ref_expected = newest_run(folders)
            if ref and not any(same_dir(ref, f) for f in folders):
                folders.append(ref)
        recs = [] if ref else list(row.get("missing") or [])
        carried = []
        for folder in folders:
            if not folder and not any(not e.get("out_dir") for e in live + no_data):
                continue
            got = refresh_folder(
                folder,
                [e for e in live if e.get("out_dir") == folder],
                [e for e in no_data if e.get("out_dir") == folder],
                row=row if folder and folder == row.get("out_dir") else None,
                late=late, args=args)
            if folder and ref and not same_dir(folder, ref):
                # 更早某一趟的目录。它缺什么是那一趟的事，结论已经写回它自己
                # 的 REPORT.md；不能算到这一趟头上 —— 2026-10-13 实测：这一趟
                # 20 张全到，却因为上一趟留下的 3 条被标成 partial、还发了告警。
                carried.extend(dict(r, carried_from=os.path.basename(folder))
                               for r in got)
            else:
                recs.extend(got)
        if ref:
            settle_carried(row, recs, carried, live, ref, ref_expected)
        row["missing"] = recs
        row["carried"] = carried
        if row.get("status") == "ok" and store_needs_attention(row):
            # 没有路线报错，但应有的报表不全：不能再叫 ok
            row["status"] = "partial"
    # 过期的条目属于更早的某一趟：结论写回它自己的目录，不记到这一趟头上
    by_folder = {}
    for r in expired or ():
        by_folder.setdefault(r.get("out_dir"), []).append(r)
    for folder, recs in by_folder.items():
        if folder and not any(folder == row.get("out_dir") for row in rows):
            refresh_folder(folder,
                           [e for e in everything if e.get("out_dir") == folder],
                           [], args=args, expired=recs)
    return rows


def drop_superseded(entries):
    """补收**之前**先撤掉已经作废的旧欠账，返回剩下的。

    一条欠账属于更早的某一趟；这家店磁盘上最新的那一趟要是已经下到了同一种
    报表，它就没用了（出报表只用最新那一趟）。以前要等补收完、诊断时才撤 ——
    2026-10-13 实测一次补取为这样一条旧欠账开店、白等了 7 分钟才发现。
    """
    by_store = {}
    for e in entries:
        by_store.setdefault(e.get("store"), []).append(e)
    moot = []
    for store, es in by_store.items():
        ref, expected = newest_run([e.get("out_dir") for e in es if e.get("out_dir")])
        if not ref:
            continue
        meta = run_downloads.read_run_meta(ref) or {}
        lacking = set(r.get("form") for r in meta.get("missing") or []
                      if isinstance(r, dict))
        delivered = set(expected or []) - lacking
        for e in es:
            if (e.get("out_dir") and not same_dir(e["out_dir"], ref)
                    and download_diag.form_of_entry(e) in delivered):
                moot.append(e)
    if moot:
        pending_store.remove(moot)
        for e in moot:
            print("    %-16s 撤销旧欠账 %s（来自 %s）：最新一趟 %s 已经下到"
                  % (e.get("store"), download_diag.form_of_entry(e),
                     os.path.basename(os.path.normpath(e["out_dir"])),
                     os.path.basename(os.path.normpath(newest_run(
                         [e["out_dir"]])[0] or "?"))))
    return [e for e in entries if e not in moot]


def collect_pass(client, names, args, rows):
    """Reopen each store once and collect the reports deferred during its run.

    This is where the reordering pays off. The slow MercadoPago reports keep
    generating server-side whether or not the store's browser is open, so the
    per-store run no longer waits on them - by the time this pass starts, the
    first store's reports are hours old and the last store's have had the whole
    batch to build.

    Stores are visited in the SAME ORDER they were downloaded, which is what
    gives the last store its head start: it is collected last, so it gains the
    duration of every other store's collection for free.

    The budget per store is short on purpose (--late-timeout, default 420s).
    Anything still not ready stays in the registry and a later run picks it up;
    waiting longer here just re-creates the serial stall this pass removes.
    """
    outstanding = drop_superseded(pending_store.load())
    mine = [e for e in outstanding if e.get("store") in set(names)]
    if not mine:
        print("%s>>> 补收阶段：没有待取回的报表" % NEWLINE)
        return
    others = len(outstanding) - len(mine)

    print("%s%s" % (NEWLINE, "=" * 60))
    print("补收阶段 —— %d 张报表，涉及 %d 家店铺"
          % (len(mine), len({e["store"] for e in mine})))
    if others:
        print("  （另有 %d 张属于本批次以外的店铺，不处理）" % others)
    print("=" * 60)

    by_store = {}
    for e in mine:
        by_store.setdefault(e["store"], []).append(e)

    for name in names:                      # same order as the download loop
        entries = by_store.get(name)
        if not entries:
            continue
        row = next((r for r in rows if r["store"] == name), None)
        print("%s--- 补收：%s（%d 张报表）"
              % (NEWLINE, name, len(entries)))

        result = {"store": name, "sales": [], "mp": [], "pending": [],
                  "empty": [], "errors": [], "retry_later": [], "no_data": []}
        late_dir = None
        keep = []
        by_dir = {}                 # 目录 -> 这一趟补收落在那里的文件
        try:
            store = client.find_store(name)
            with client.open_store(store, close_when_done=True) as session:
                d = session.driver
                d.get(session.launcher_page)
                # Entries from different runs land in different folders, and the
                # download directory is a per-SESSION CDP setting - so it has to
                # be re-applied here, once per folder. Forgetting this is how a
                # route ends up reporting "got 0 files" while clicking correctly.
                folders = {}
                for e in entries:
                    folders.setdefault(e.get("out_dir"), []).append(e)
                for home, group in folders.items():
                    out_dir = home
                    if not out_dir or not os.path.isdir(out_dir):
                        # the run folder is gone; put the file with this batch
                        out_dir = run_downloads.run_folder(name, args._stamp,
                                                           args.out)
                        if not os.path.isdir(out_dir):
                            os.makedirs(out_dir)
                        print("    原下载目录已不存在，改存到 %s" % out_dir)
                    meli_forms.set_download_dir(d, out_dir)
                    late_dir = out_dir
                    n_sales, n_mp = len(result["sales"]), len(result["mp"])
                    left = run_downloads.collect_pending(
                        d, group, out_dir, result,
                        budget=args.late_timeout, poll=args.collect_poll)
                    keep.extend(left)
                    by_dir[home] = {"sales": result["sales"][n_sales:],
                                    "mp": result["mp"][n_mp:]}
                d.get(session.launcher_page)
        except Exception as e:
            # A store that will not open must not cost the other stores' reports.
            # Its entries stay in the registry, so nothing is lost.
            print("    [失败] %s: %s" % (e.__class__.__name__, e))
            result["errors"].append("late collect: %s: %s"
                                    % (e.__class__.__name__, e))
            # 已经确认没有数据的不必再欠着；其余全部照旧欠着
            keep = [x for x in entries
                    if not any(x is y for y in result.get("no_data", []))]
            # 没有浏览器可截图：异常类、卡在哪一步、紫鸟的原话就是全部现场
            for x in keep:
                download_diag.note_failure(x, e)

        # Still generating when the budget ran out. Separated from retry_later
        # because that one already has its own summary line ("datos en proceso")
        # and would otherwise be counted twice.
        waiting = list(keep)
        keep.extend(x for x in result.get("retry_later", [])
                    if not any(x is k for k in keep))
        collected = len(entries) - len(keep)
        print("    已取回 %d 张，仍欠 %d 张" % (collected, len(keep)))

        # Everything not kept is done with - downloaded, or a range that will
        # never produce a report. Removing it is what stops the registry growing
        # without bound.
        done = [e for e in entries if not any(e is k for k in keep)]
        if done:
            pending_store.remove(done)
        if keep:
            # 把这一趟探测到的东西（次数、首末时间、最后状态、证据、原因）
            # 写回登记簿：下一趟、以及批次末尾的诊断，都从这里读。
            for e in keep:
                e["cause"] = download_diag.record_from_entry(e)["cause"]
            pending_store.add(keep)

        if row is not None:
            if result["sales"] and late_dir:
                # ventas arrived late, so this store has no workbook yet
                row["ventas"] = list(row.get("ventas", [])) + [
                    rel_to_downloads(late_dir, f) for f in result["sales"]]
            row["files"] = row.get("files", 0) + len(result["mp"]) + len(result["sales"])
            row["late_collected"] = len(result["mp"]) + len(result["sales"])
            row["empty"] = list(row.get("empty", [])) + result["empty"]
            # Report what is still owed. The registry knowing about it is not
            # enough: a summary that says pending=[] while a report is still
            # generating is the same "silence reads as success" failure that
            # once let six failed requests pass as `ok, 13 files`.
            row["pending"] = (
                [p for p in row.get("pending", [])
                 if run_downloads.DELAYED_MARK not in p]
                + result["pending"]
                + [owed_label(e) for e in waiting])
            row["_late"] = {"no_data": list(result.get("no_data") or []),
                            "by_dir": by_dir}
            if result["errors"]:
                row["errors"] = list(row.get("errors", [])) + result["errors"]
                if row["status"] == "ok":
                    row["status"] = "partial"



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--stores", nargs="+", default=None,
                    help="store names to run; default is config stores.batch")
    ap.add_argument("--limit", type=int, default=0,
                    help="run only the first N stores (0 = all)")
    ap.add_argument("--out", default=None,
                    help="base output dir; each store gets <base>/<STORE>/<stamp>/")
    ap.add_argument("--pace", type=int, default=5,
                    help="seconds between stores (default 5)")
    ap.add_argument("--dry-run", action="store_true",
                    help="print the store list and exit without opening anything")
    ap.add_argument("--no-restart", action="store_true",
                    help="assume the client is already in webdriver mode")
    ap.add_argument("--collect-only", action="store_true",
                    help="download nothing; just reopen the stores that still "
                         "owe reports (logs/pending.json) and collect them. This "
                         "is how you pick up what a previous run left generating.")
    ap.add_argument("--no-late-collect", action="store_true",
                    help="do not reopen stores to collect the reports their own "
                         "run deferred; they stay in logs/pending.json for a "
                         "later run")
    ap.add_argument("--workbooks", action="store_true",
                    help="after cleaning, also write one accounting workbook per "
                         "store into downloads/data/reports/. With this the run "
                         "goes download -> clean -> report with no human input.")
    ap.add_argument("--no-transform", action="store_true",
                    help="skip the cleaning pipeline in downloads/mx_sales, "
                         "which otherwise runs once after every store is done")
    ap.add_argument("--stop-on-error", action="store_true",
                    help="abort the batch on the first store that fails "
                         "(default is to carry on)")
    run_downloads.add_route_args(ap)
    args = ap.parse_args()

    names = resolve_stores(args)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    stamp = run_downloads.month_stamp(stamp, args)

    stale = pending_store.prune()
    expired = [download_diag.expired_record(e) for e in stale]
    if stale:
        print("已清理 %d 张超过 %d 天的过期待取记录（留档在 %s）："
              % (len(stale), pending_store.MAX_AGE_DAYS,
                 pending_store.expired_path()))
        for e, rec in zip(stale, expired):
            print("    %s —— 申请于 %s，探测过 %d 次，最后状态：%s"
                  % (pending_store.describe_entry(e), rec.get("requested_at"),
                     e.get("probes") or 0, e.get("last_state") or "从未探测"))

    print("=" * 60)
    print("批次 %s —— %d 家店铺：%s" % (stamp, len(names), ", ".join(names)))
    print("路线：%s | 白名单：%s" % (args.only, store_config.describe()))
    print("=" * 60)
    if args.dry_run:
        for i, n in enumerate(names, 1):
            print("  %2d. %-16s -> %s"
                  % (i, n, run_downloads.run_folder(n, stamp, args.out)))
        return 0

    client = ZiniaoClient()
    client.start(restart=not args.no_restart)

    rows = []
    began = time.time()

    if args.collect_only:
        # Nothing to download: this run exists only to finish what an earlier
        # one started. Restrict to stores that actually owe something, so a
        # 12-store allowlist does not open 12 browsers to do nothing.
        owed = {e.get("store") for e in pending_store.load()}
        names = [n for n in names if n in owed]
        if not names:
            print("%s这些店铺没有待取回的报表，无需补收"
                  % NEWLINE)
            return 0
        print("%s仅补收模式：%d 家店铺有待取回的报表：%s"
              % (NEWLINE, len(names), ", ".join(names)))
        rows = [{"store": n, "status": "ok", "files": 0, "errors": [],
                 "pending": [], "empty": [], "out_dir": None, "seconds": 0}
                for n in names]

    for i, name in enumerate([] if args.collect_only else names, 1):
        print("\n" + "#" * 60)
        print("# [%d/%d] %s" % (i, len(names), name))
        print("#" * 60)
        row = {"store": name, "status": "failed", "files": 0,
               "errors": [], "pending": [], "empty": [],
               "out_dir": None, "seconds": 0}
        t0 = time.time()
        try:
            out_dir = run_downloads.run_folder(name, stamp, args.out)
            # close_when_done is not negotiable in a batch: the next store
            # needs the browser slot, and a store left open would still be
            # holding its proxy.
            result = run_downloads.run_store(client, name, args,
                                             out_dir=out_dir,
                                             close_when_done=True,
                                             defer_slow=True)
            run_downloads.print_summary(result)
            # Keep the ventas export's path RELATIVE TO downloads/, which is
            # what the cleaner's matcher wants. A bare filename would also match,
            # but against every copy in the tree.
            row["ventas"] = [rel_to_downloads(result["out_dir"], f)
                             for f in result.get("sales", [])]
            deferred = result.get("deferred", []) + result.get("retry_later", [])
            if deferred:
                entries = []
                for p in deferred:
                    e = dict(p)
                    e["store"] = name
                    e["out_dir"] = result["out_dir"]
                    e.setdefault("requested_at", time.time())
                    e["cause"] = download_diag.record_from_entry(e)["cause"]
                    entries.append(e)
                # Written NOW, while the store is fresh in hand. A crash between
                # this store and the late pass would otherwise leave reports
                # generating on MercadoPago that nobody knows to collect.
                pending_store.add(entries)
            row.update(
                status="partial" if result["errors"] else "ok",
                files=run_downloads.count_files(result),
                errors=result["errors"],
                pending=result["pending"],
                empty=result.get("empty", []),
                out_dir=result["out_dir"],
                expected=result.get("expected", []),
                missing=result.get("missing", []),
                flags=result.get("flags", {}),
                cautions=result.get("cautions", []),
                _result=result,
                _started=datetime.datetime.fromtimestamp(t0))
        except StoreNotAllowed as e:
            row["status"] = "not-allowed"
            row["errors"] = [str(e)]
            print("  [跳过] %s" % e)
        except Exception as e:
            # Deliberately broad. The point of a batch is that store 7 dying in
            # a way nobody predicted still leaves stores 8-25 to run.
            row["errors"] = ["%s: %s" % (e.__class__.__name__, e)]
            print("  [失败] %s: %s" % (e.__class__.__name__, e))
            traceback.print_exc()
            # 整家店没跑成：应有的报表全部缺失，原因是同一个。没有浏览器可
            # 截图，现场就是异常本身。
            row["expected"] = download_diag.expected_forms(args)
            row["missing"] = download_diag.store_failure_records(
                e, row["expected"])
        row["seconds"] = int(time.time() - t0)
        # The workbook is written HERE, per store, not batched to the end.
        if args.workbooks:
            make_workbook(row, stamp, args)
        rows.append(row)

        if row["status"] not in ("ok", "partial") and args.stop_on_error:
            print("\n--stop-on-error：在 %s 之后中止批次" % name)
            break
        if i < len(names):
            time.sleep(args.pace)

    # ---- collect what the per-store runs deferred ------------------------
    args._stamp = stamp
    # 补收可能把文件落进**更早那一轮**的目录（报表是那一轮申请的）。先记下
    # 这些目录，清洗时它们和本轮目录一样算"这一次的文件"。
    touched = set(os.path.basename(e["out_dir"]) for e in pending_store.load()
                  if e.get("out_dir"))
    if not args.no_late_collect:
        collect_pass(client, names, args, rows)
    elif pending_store.load():
        print("%s>>> 已跳过补收阶段（--no-late-collect），仍欠：%s"
              % (NEWLINE, pending_store.describe()))

    # ---- hand the new files to the cleaning pipeline --------------------
    # Once, after every store: it walks the whole download tree, so running it
    # per store would repeat the same work. A failure here is reported but
    # cannot undo the downloads, which are already on disk.
    tr = None
    if not args.no_transform:
        # 清洗每次都扫整棵 downloads 树。以前某一轮留下的坏文件会让之后**每
        # 一轮**都报"清洗部分失败"、整批退出码 1 —— 2026-10-09 那轮 902 个文件
        # 里 28 个失败，全是 9 月的旧导出。告诉清洗哪些目录是这一轮的，旧文件
        # 的失败只作为提示列出，不再算这一轮的账。
        touched.add(stamp)
        touched.update(os.path.basename(r["out_dir"]) for r in rows
                       if r.get("out_dir"))
        tr = transform.build(current_runs=sorted(touched))
        if tr and tr.get("failed_older"):
            print("    另有 %d 个更早下载的文件清洗不了，与这一轮无关，"
                  "不影响本次结果" % len(tr["failed_older"]))
    elif transform.available():
        print("\n>>> 已跳过清洗（--no-transform），之后可手动执行：")
        print("    cd downloads && python -m mx_sales build")

    # ---- catch up a store whose ventas only arrived in the late pass -----
    # Workbooks are normally written the moment their store finishes. This is
    # only for a store whose ventas missed the in-session budget and was
    # collected late; without it that store would silently end up without one.
    if args.workbooks:
        late = [r for r in rows if r.get("ventas") and not r.get("workbook")]
        if late:
            print("%s>>> 对账簿：%d 家店铺的 Ventas 是补收回来的，现在补生成"
                  % (NEWLINE, len(late)))
            for row in late:
                make_workbook(row, stamp, args)
        for row in rows:
            if not row.get("ventas") and row["status"] in ("ok", "partial"):
                print("    %-16s 没有 Ventas 导出文件，不生成对账簿" % row["store"])

    # ---- 诊断这一步：没到的报表，各是为什么 ------------------------------
    # 排在补收之后：店内只探测过一次、只能说"分不清"的，这时已经有了补收
    # 时的证据。只用已有的证据，不为了诊断重开店铺。
    try:
        finalize_diagnosis(rows, args=args, expired=expired)
    except Exception as e:           # 诊断出错不能把一整批下载说成失败
        print("%s  [警告] 诊断这一步自己出错了：%s: %s"
              % (NEWLINE, e.__class__.__name__, e))
        traceback.print_exc()

    verdict = batch_verdict(rows, tr, expired)
    log_path = os.path.join(LOG_DIR, "batch_%s.json" % stamp)
    write_summary(log_path, stamp, rows, transform_result=tr, extra=verdict)

    elapsed = int(time.time() - began)
    print("\n" + "=" * 60)
    print("批次汇总 —— 用时 %d 分 %02d 秒" % (elapsed // 60, elapsed % 60))
    print("=" * 60)
    print("  %-16s %-12s %6s  %s" % ("店铺", "状态", "文件数", "备注"))
    for r in rows:
        # Every fact that applies, not just the first one found. An earlier
        # version stopped at the first match, so a store with one report still
        # generating was reported as "1 still generating" and nothing else -
        # hiding that it had produced its workbook perfectly well.
        parts = []
        if r["errors"]:
            parts.append("%d 处错误：%s" % (len(r["errors"]), r["errors"][0][:44]))
        owed = [m for m in r.get("missing") or [] if m.get("owed")]
        if owed or r["pending"]:
            parts.append("%d 张仍欠" % (len(owed) or len(r["pending"])))
        lost = [m for m in r.get("missing") or []
                if not m.get("owed") and download_diag.needs_attention(m)]
        if lost:
            parts.append("%d 张未下载" % len(lost))
        if r["empty"]:
            # Worth showing, but not a problem: the range simply held nothing.
            parts.append("%d 张无流水" % len(r["empty"]))
        if r.get("workbook"):
            parts.append("已出对账簿")
        print("  %-16s %-12s %6d  %s"
              % (r["store"], r["status"], r["files"], " | ".join(parts)))
    print("\n  汇总已写入 %s" % log_path)

    for line in download_diag.format_section(rows, expired=expired):
        print(line)
    print("")

    if tr is not None:
        if tr.get("skipped"):
            print("  清洗            已跳过（%s）" % tr["skipped"])
        elif tr["ok"]:
            print("  清洗            成功")
        elif tr.get("partial"):
            print("  清洗            部分文件失败，其余已写入")
        else:
            print("  清洗            失败")

    code = verdict["exit_code"]
    bad = verdict["stores_attention"] + verdict["stores_failed"]
    if code == 2:
        print("  %d 家店铺没有一家下载成功 —— 这一趟没有任何新数据" % len(rows))
    elif bad:
        print("  %d/%d 家店铺需要关注（其中 %d 家完全失败）"
              % (bad, len(rows), verdict["stores_failed"]))
    if verdict["cleaning"] in ("partial", "failed"):
        # The downloads succeeded; only the transform did not. Worth a non-zero
        # exit so a scheduled run does not look clean, but say which half broke
        # - and how much of it. "清洗失败" for 1 bad file out of 55 reads as if
        # nothing was cleaned, which sends people looking in the wrong place.
        if verdict["cleaning"] == "partial":
            print("  清洗有文件未通过，其余已写入（这不是店铺下载失败）—— "
                  "上面的 ERROR 行写了是哪个文件")
        else:
            print("  清洗失败（这不是店铺下载失败，下载的文件都在盘上）")
    if code == 0:
        print("  全部 %d 家店铺正常" % len(rows))
    # 最后一行，给 run_monthly 读的：人看的日志里会被它滤掉
    print("[BATCH_SUMMARY_JSON] %s" % os.path.abspath(log_path))
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
    except ZiniaoError as e:
        print("\n紫鸟客户端错误：%s" % e)
        sys.exit(2)
