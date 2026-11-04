# -*- coding: utf-8 -*-
"""Move one machine's data from the old layout to the new one (docs/DATA_LAYOUT.md).

    python -m tools.migrate_layout plan   --data-root D:\\marketplace-data
    python -m tools.migrate_layout copy   --data-root D:\\marketplace-data
    python -m tools.migrate_layout remove-old --manifest D:\\marketplace-data\\logs\\migration\\manifest_….json
    python -m tools.migrate_layout undo

plan        lists what would go where, what cannot be placed, how much space it
            needs. Writes nothing.
copy        copies every file into the new tree, checks each copy (size and
            SHA-256), rewrites the moved folder paths inside logs/pending*.json
            and the moved run folders / workbooks inside results/**.json
            (source_folders, workbook), writes a manifest, then switches config.json to the new layout
            (paths.data_root; a backup of config.json is kept). The old folders
            are left exactly as they were. Safe to run again: files already
            copied and identical are skipped; a different file already at the
            target stops the copy for that file and is reported.
remove-old  deletes the old files listed in a manifest, each only if both the
            old file and its copy still have the recorded checksum; then removes
            the old folders that became empty. Run it after checking the new
            layout works (a report run, the window).
undo        removes paths.data_root from config.json: the machine is back on the
            old layout (as long as remove-old has not been run). The new tree is
            left in place — so are the copies sorted inside the window's report
            folder (see below): their flat originals are still at the top of
            that folder until remove-old, which is where the old layout looks.

Every command refuses to run while a monthly run is active. Run it between
monthly runs with the scheduled task paused.

What goes where (old -> new):
  downloads/<store>/<run>/              raw/<PF>/<CC>/<store>/<YYYY-MM or 其他>/<run>/
  downloads/data/reports/financial/*    reports/monthly/<YYYY-MM>/<PF>-<CC>/  (store and
                                        country workbooks, metrics), reports/monthly/<YYYY-MM>/
                                        (alert files, 本月总览), reports/_旧文件/ (anything else)
  downloads/data/processed/, mx_sales.duckdb, reports/<store>_<stamp>.xlsx
                                        work/mx_sales/…
  downloads/data/results/               results/
  downloads/data/<other>                work/_旧文件/<other>
  logs/                                 logs/

The window's 「报表输出到」 (gui_settings.json out_dir), when it names a folder
other than the old default (e.g. S:\\Finance\\电商数据 on a shared drive): in the
old layout that folder IS the flat report folder, in the new one it is the
reports root (core/layout.reports_root), so its files are sorted IN PLACE, under
that same folder, never under data_root/reports:
  <out_dir>/<store>_<YYYY-MM>*.xlsx|json, <CC>_<PF>_全店汇总_*, 全店_*_指标.json
                                        <out_dir>/monthly/<YYYY-MM>/<PF>-<CC>/
  <out_dir>/！报表异常_*_<YYYYMM>.txt, ！需要处理_<YYYYMM>.txt, 本月总览_<YYYY-MM>.xlsx
                                        <out_dir>/monthly/<YYYY-MM>/
  <out_dir>/期间报表/<kind>/<period>/<market>/*
                                        <out_dir>/<kind>/<period>/<market>/*  (quarterly, annual, custom)
  <out_dir>/_versions/, 版本说明.txt      <out_dir>/_旧文件/   (ours: old versions)
  <out_dir>/<anything else>             LEFT WHERE IT IS — the folder is shared with
                                        people (the finance team's own month folders
                                        and downloads live next to our reports; on
                                        2026-11-04 1,536 of theirs vs 108 of ours)
Excel's ~$… lock files are skipped everywhere. Files already under
<out_dir>/monthly, quarterly, annual, custom or _旧文件 are left alone, so the
sort can run again. A source file that cannot be read during copy (open in
Excel, share hiccup) is reported at the end, the rest is copied, config.json is
not switched; run copy again after closing it. These moves are copies like all the
others (verified, listed in the manifest, originals deleted by remove-old), and
results/**.json "workbook" paths into that folder follow them. A window setting
inside downloads/ or logs/ is handled by the ordinary walk instead.

A store's platform and country come from config.json; failing that from the run
folder's run_meta.json. A folder that cannot be placed is listed and not copied.
"""
import argparse
import datetime
import hashlib
import io
import json
import os
import re
import shutil
import sys

from core import console  # noqa: F401  中文输出编码保护
from core import layout
from core import markets
from core import paths
from core import store_config

NOT_STORES = {"data", "mx_sales", "tests", "__pycache__", ".pytest_cache"}
REPORT_RE = re.compile(r"^(?P<store>.+)_(?P<month>\d{4}-\d{2})(?P<rest>(_账单基准)?(_指标)?)\.(xlsx|json)$")
COMBINED_RE = re.compile(r"^(?P<cc>[A-Z]{2})_(?P<pf>[A-Z]{2})_全店汇总_(?P<month>\d{4}-\d{2})(_账单基准)?\.xlsx$")
ALL_METRICS_RE = re.compile(r"^全店_(?P<month>\d{4}-\d{2})_指标\.json$")
ALERT_RE = re.compile(r"^！(?:报表异常_(?:.*_)?|需要处理_)(?P<ym>\d{6})\.txt$")
OVERVIEW_RE = re.compile(r"^本月总览_(?P<month>\d{4}-\d{2})(?:\(新\))?\.xlsx$")
MX_SALES_BOOK_RE = re.compile(r"^.+_\d{8}_\d{6}.*\.xlsx$")
OLD_FILES = "_旧文件"
LEGACY_PERIOD_FOLDER = "期间报表"            # layout.period_dir, legacy branch
PERIOD_KINDS = ("quarterly", "annual", "custom")
NEW_REPORT_DIRS = ("monthly",) + PERIOD_KINDS + (OLD_FILES,)   # under a reports root: left alone


class Refused(Exception):
    """A reason not to touch anything."""


# ------------------------------------------------------------------ helpers

def legacy_downloads():
    return os.path.join(paths.ROOT, "downloads")


def legacy_logs():
    return os.path.join(paths.ROOT, "logs")


def old_default_reports():
    """The old layout's default flat report folder."""
    return os.path.join(paths.ROOT, "downloads", "data", "reports", "financial")


def _under(path, top):
    """Is `path` inside `top` (case and separators as the OS sees them)?"""
    p, t = os.path.normcase(os.path.abspath(path)), os.path.normcase(os.path.abspath(top))
    return p == t or p.startswith(t.rstrip(os.sep) + os.sep)


def window_out_dir():
    """The window's 「报表输出到」 (absolute) when it names a folder other than
    the old default report folder; None when unset or the default."""
    from core import entrance
    out = read_json(entrance.SETTINGS).get("out_dir")
    if not out or not str(out).strip():
        return None
    out = os.path.abspath(str(out).strip())
    if os.path.normcase(out) == os.path.normcase(os.path.abspath(old_default_reports())):
        return None
    return out


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


LOCK_PREFIX = "~$"          # Excel's lock file next to an open workbook: not a file, never copied


def files_under(top):
    for dirpath, _dirs, files in os.walk(top):
        for f in files:
            if f.startswith(LOCK_PREFIX):
                continue
            yield os.path.join(dirpath, f)


def read_json(path):
    try:
        with io.open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (IOError, OSError, ValueError):
        return {}


def refuse_if_running():
    from core import runlog
    if runlog.active_run(legacy_logs()) or os.path.exists(os.path.join(legacy_logs(), "monthly.lock")):
        raise Refused("有一轮月度运行正在进行（或留下了 logs/monthly.lock）。等它结束、暂停计划任务后再迁移。")


def check_target(data_root):
    root = os.path.normcase(os.path.abspath(paths.ROOT))
    target = os.path.normcase(os.path.abspath(data_root))
    if target == root or target.startswith(root + os.sep):
        raise Refused("数据目录不能放在代码目录 %s 里面（这正是要分开的两样东西）。" % paths.ROOT)


# ------------------------------------------------------------------ where a store belongs

def store_market(store, run_dirs):
    """(PF, CC, note) or (None, None, reason)."""
    try:
        cat = store_config._catalog_entries()
    except store_config.StoreConfigError as e:
        raise Refused(str(e))
    if store in cat:
        return cat[store][0], cat[store][1], ""
    current = store_config.current_name(store)          # a renamed store's old folder (catalog "formerly")
    if current != store and current in cat:
        return cat[current][0], cat[current][1], "旧名，现在叫 %s" % current
    for d in run_dirs:
        meta = read_json(os.path.join(d, "run_meta.json"))
        pf, cc = meta.get("platform"), meta.get("country")
        if pf and cc:
            return pf, cc, "按 run_meta.json"
        if meta.get("site") and not meta.get("platform"):
            return markets.DEFAULT_PLATFORM, str(meta["site"]).upper(), "按 run_meta.json 的 site"
    try:
        if not store_config._uses_catalog():
            pf, cc = store_config.market_parts(store)
            return pf, cc, "旧配置：默认美客多墨西哥"
    except store_config.StoreConfigError:
        pass
    return None, None, "config.json 里没有这家店，下载目录也没写平台/国家"


def run_month(run_dir):
    return layout.month_of_run(run_dir) or read_json(os.path.join(run_dir, "run_meta.json")).get("month")


# ------------------------------------------------------------------ the plan

def build_plan(data_root):
    """{"items": [(kind, src, dst)], "unplaced": [(path, reason)], "notes": [...],
    "folder_map": {old run folder: new run folder}} — files only, nothing written."""
    items, unplaced, notes, folder_map = [], [], [], {}
    dl = legacy_downloads()
    store_markets = {}
    if os.path.isdir(dl):
        for name in sorted(os.listdir(dl)):
            top = os.path.join(dl, name)
            if not os.path.isdir(top) or name in NOT_STORES:
                continue
            runs = [os.path.join(top, r) for r in sorted(os.listdir(top))
                    if r[:2] == "20" and os.path.isdir(os.path.join(top, r))]
            if not runs:
                notes.append("跳过 %s：里面没有下载运行目录" % top)
                continue
            pf, cc, how = store_market(name, runs)
            if not pf:
                unplaced.append((top, how))
                continue
            store_markets[name] = "%s-%s" % (pf, cc)
            if how:
                notes.append("%s → %s-%s（%s）" % (name, pf, cc, how))
            for r in runs:
                dst = os.path.join(data_root, "raw", pf, cc, name, run_month(r) or layout.NO_MONTH,
                                   os.path.basename(r))
                folder_map[r] = dst
                for f in files_under(r):
                    items.append(("raw", f, os.path.join(dst, os.path.relpath(f, r))))
    data = os.path.join(dl, "data")
    if os.path.isdir(data):
        for name in sorted(os.listdir(data)):
            src = os.path.join(data, name)
            if name == "reports":
                _plan_reports(src, data_root, store_markets, items, unplaced)
            elif name == "processed":
                items += [("work", f, os.path.join(data_root, "work", "mx_sales", "processed",
                                                    os.path.relpath(f, src))) for f in files_under(src)]
            elif name == "mx_sales.duckdb":
                items.append(("work", src, os.path.join(data_root, "work", "mx_sales", name)))
            elif name == "results":
                items += [("results", f, os.path.join(data_root, "results", os.path.relpath(f, src)))
                          for f in files_under(src)]
            elif os.path.isdir(src):
                items += [("old", f, os.path.join(data_root, "work", "_旧文件", name, os.path.relpath(f, src)))
                          for f in files_under(src)]
            else:
                items.append(("old", src, os.path.join(data_root, "work", "_旧文件", name)))
    lg = legacy_logs()
    if os.path.isdir(lg):
        items += [("logs", f, os.path.join(data_root, "logs", os.path.relpath(f, lg))) for f in files_under(lg)
                  if os.path.basename(f) != "monthly.lock"]
    out = window_out_dir()
    if out:
        _plan_out_dir(out, store_markets, items, notes)
    return {"items": items, "unplaced": unplaced, "notes": notes, "folder_map": folder_map, "out_dir": out}


def _plan_reports(reports, data_root, store_markets, items, unplaced):
    """downloads/data/reports: financial/ (the flat report folder), the cleaning
    pipeline's workbooks, experiments and anything else."""
    for name in sorted(os.listdir(reports)):
        src = os.path.join(reports, name)
        if name == "financial" and os.path.isdir(src):
            _plan_flat_reports(src, os.path.join(data_root, "reports"), store_markets, items)
        elif os.path.isfile(src) and MX_SALES_BOOK_RE.match(name):
            items.append(("work", src, os.path.join(data_root, "work", "mx_sales", "reports", name)))
        elif os.path.isdir(src):
            items += [("old", x, os.path.join(data_root, "work", "_旧文件", "reports", name, os.path.relpath(x, src)))
                      for x in files_under(src)]
        else:
            items.append(("old", src, os.path.join(data_root, "work", "_旧文件", "reports", name)))


# What this program itself writes next to the reports in the old flat folder
# (core/finalize): the per-store version folders and the final-version note.
OURS_DIRS = ("_versions",)
OURS_FILES = ("版本说明.txt",)


def _plan_flat_reports(flat, reports_root, store_markets, items, in_place=False):
    """The files of one old flat report folder `flat` sorted under
    `reports_root` (data_root/reports for the old default folder; the folder
    itself for the window's own folder, in_place).

    The old default folder is ours entirely: sub-folders and unknown files go
    to _旧文件. The window's folder (in place) is shared with people — on the
    production machine it is the finance team's drive, with their own month
    folders and downloads next to our reports — so there only what this
    program wrote is touched: reports, alerts, overviews, 期间报表/ (sorted
    into quarterly/ annual/ custom/), _versions/ and 版本说明.txt (to _旧文件).
    The new layout's own folders are left alone; everything else is left
    exactly where it is and counted. Returns the number of foreign entries."""
    foreign = 0
    for f in sorted(os.listdir(flat)):
        p = os.path.join(flat, f)
        if os.path.isdir(p):
            if in_place and f in NEW_REPORT_DIRS:
                continue
            if in_place and f == LEGACY_PERIOD_FOLDER:
                foreign += _plan_period_folder(p, reports_root, items, in_place)
                continue
            if in_place and f not in OURS_DIRS:
                foreign += 1
                continue
            items += [("old", x, os.path.join(reports_root, OLD_FILES, f, os.path.relpath(x, p)))
                      for x in files_under(p)]
            continue
        if f.startswith(LOCK_PREFIX):
            continue
        dst = report_target(f, reports_root, store_markets)
        if dst is None and in_place and f not in OURS_FILES:
            foreign += 1
            continue
        items.append(("report", p, dst or os.path.join(reports_root, OLD_FILES, f)))
    return foreign


def _plan_period_folder(folder, reports_root, items, in_place=False):
    """期间报表/<kind>/<period>/<market>/… -> <reports_root>/<kind>/<period>/<market>/…;
    anything else in there -> _旧文件/期间报表/… (left alone, counted, in place).
    Returns the number of foreign files."""
    foreign = 0
    for f in files_under(folder):
        rel = os.path.relpath(f, folder)
        if rel.split(os.sep)[0] in PERIOD_KINDS:
            items.append(("report", f, os.path.join(reports_root, rel)))
        elif in_place:
            foreign += 1
        else:
            items.append(("old", f, os.path.join(reports_root, OLD_FILES, LEGACY_PERIOD_FOLDER, rel)))
    return foreign


def _plan_out_dir(out, store_markets, items, notes):
    """The window's 「报表输出到」 when it is not the old default: the old
    layout's flat report folder AND the new layout's reports root, so its files
    are sorted in place (module docstring). Inside downloads/ or logs/ the
    ordinary walk already copies it (twice would conflict)."""
    if _under(out, legacy_downloads()) or _under(out, legacy_logs()):
        notes.append("窗口里「报表输出到」%s 在代码目录的旧数据里，按普通旧文件搬走，没有就地整理" % out)
        return
    if not os.path.isdir(out):
        raise Refused("窗口里「报表输出到」是 %s，现在打不开（共享盘没连上？）。连上它，或先在窗口里改掉这个设置，再迁移。"
                      % out)
    before = len(items)
    foreign = _plan_flat_reports(out, out, store_markets, items, in_place=True)
    mine = items[before:]
    old = sum(1 for _k, _s, d in mine if _under(d, os.path.join(out, OLD_FILES)))
    notes.append("窗口里「报表输出到」%s 不是旧的默认目录：那里本程序写的 %d 个旧文件就地整理到它自己的 monthly\\ "
                 "quarterly\\ annual\\ custom\\ 子目录里（其中 %d 个旧版本文件放进 _旧文件\\）；已经在这些子目录里的不动"
                 % (out, len(mine), old))
    if foreign:
        notes.append("那里另有 %d 个文件/文件夹不是本程序写的（别人放的），原样不动" % foreign)


def report_target(fname, reports_root, store_markets):
    """Where one file of an old flat report folder goes under `reports_root`
    (data_root/reports, or the window's folder itself), or None (→ _旧文件)."""
    monthly = os.path.join(reports_root, "monthly")
    m = ALERT_RE.match(fname)
    if m:
        ym = m.group("ym")
        return os.path.join(monthly, "%s-%s" % (ym[:4], ym[4:]), fname)
    m = OVERVIEW_RE.match(fname)
    if m:
        return os.path.join(monthly, m.group("month"), fname)
    m = COMBINED_RE.match(fname)
    if m:
        return os.path.join(monthly, m.group("month"), "%s-%s" % (m.group("pf"), m.group("cc")), fname)
    m = ALL_METRICS_RE.match(fname)
    if m:
        return os.path.join(monthly, m.group("month"), "ML-MX", fname)
    m = REPORT_RE.match(fname)
    if m:
        market = store_markets.get(m.group("store"))
        if not market:
            try:
                market = store_config.market_of(m.group("store"))
            except store_config.StoreConfigError:
                market = None
        if market:
            return os.path.join(monthly, m.group("month"), market, fname)
    return None


# ------------------------------------------------------------------ copy

def copy(plan, data_root, say=print):
    """Copy and verify every file. Returns (manifest items, conflicts).

    logs/pending*.json are not copied byte for byte: their out_dir values that
    point at a moved run folder are rewritten to the folder's new place; so are
    the results files' source_folders and workbook (core/finalize decides which
    download folders a final version uses from them — an old path would get the
    final's moved folder deleted). The manifest keeps both checksums (sha256 =
    the old file, dst_sha256 = its copy)."""
    done, conflicts, unreadable = [], [], []
    file_map = dict((src, dst) for kind, src, dst in plan["items"] if kind in ("report", "old"))
    for kind, src, dst in plan["items"]:
        try:
            _copy_one(kind, src, dst, plan, file_map, done, conflicts, say)
        except (IOError, OSError) as e:
            # a workbook open in Excel, a file being written, a share hiccup:
            # reported at the end, the rest of the copy goes on, no switch.
            unreadable.append((src, "%s: %s" % (e.__class__.__name__, e)))
            tmp = dst + ".migrating"
            if os.path.exists(tmp):
                os.remove(tmp)
    return done, conflicts, unreadable


def _copy_one(kind, src, dst, plan, file_map, done, conflicts, say):
    """Copy and verify one file (or skip an identical copy); records it in done
    or conflicts. Raises IOError/OSError when the source cannot be read."""
    digest = sha256(src)
    if _is_pending(kind, src):
        content = _pending_rewritten(src, plan["folder_map"], say)
    elif _is_results(kind, src):
        content = _results_rewritten(src, plan["folder_map"], file_map)
    else:
        content = None
    want = hashlib.sha256(content).hexdigest() if content is not None else digest
    if os.path.exists(dst):
        if sha256(dst) != want:
            conflicts.append((src, dst))
            return
    else:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        tmp = dst + ".migrating"
        if content is None:
            shutil.copy2(src, tmp)
        else:
            with open(tmp, "wb") as fh:
                fh.write(content)
        if sha256(tmp) != want:
            os.remove(tmp)
            raise IOError("复制后校验不一致：%s" % src)
        os.replace(tmp, dst)
    done.append({"kind": kind, "src": src, "dst": dst, "size": os.path.getsize(src), "sha256": digest,
                 "dst_sha256": want})


def _is_pending(kind, path):
    name = os.path.basename(path)
    return kind == "logs" and name.startswith("pending") and name.endswith(".json")


def _is_results(kind, path):
    return kind == "results" and path.endswith(".json")


def _results_rewritten(src, folder_map, file_map):
    """The bytes of a results file whose source_folders / workbook point at
    their new place, written as core/results.write writes (so a second copy
    gives the same bytes), or None when nothing changes."""
    def norm(m):
        return dict((os.path.normcase(os.path.abspath(k)), v) for k, v in m.items())
    folders, files = norm(folder_map), norm(file_map)
    doc = read_json(src)
    if not isinstance(doc, dict) or not doc:
        return None
    changed = False
    srcs = doc.get("source_folders")
    if isinstance(srcs, list):
        new = [folders.get(os.path.normcase(os.path.abspath(f)), f) if isinstance(f, str) and f else f
               for f in srcs]
        if new != srcs:
            doc["source_folders"], changed = new, True
    wb = doc.get("workbook")
    if isinstance(wb, str) and wb:
        to = files.get(os.path.normcase(os.path.abspath(wb)))
        if to and to != wb:
            doc["workbook"], changed = to, True
    if not changed:
        return None
    return json.dumps(doc, ensure_ascii=False, indent=1, sort_keys=True).encode("utf-8")


def _pending_rewritten(src, folder_map, say):
    """The bytes of a pending*.json with moved run folders pointing at their new
    place, or None when nothing changes (then the file is copied as is)."""
    norm = dict((os.path.normcase(os.path.abspath(k)), v) for k, v in folder_map.items())
    doc = read_json(src)
    changed = 0
    for e in doc.get("entries") or []:
        new = norm.get(os.path.normcase(os.path.abspath(e["out_dir"]))) if e.get("out_dir") else None
        if new:
            e["out_dir"] = new
            changed += 1
    if not changed:
        return None
    say("  %s：%d 条待取记录改到了新的下载目录" % (os.path.basename(src), changed))
    return json.dumps(doc, ensure_ascii=False, indent=2).encode("utf-8")


def switch_config(data_root, say=print, sorted_in_place=None):
    """Write paths.data_root into config.json (backup first). sorted_in_place:
    how many files of the window's own report folder copy() sorted in place
    (None when unknown)."""
    cfg_path = store_config.CONFIG_PATH
    with io.open(cfg_path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    backup = "%s.bak_layout_%s" % (cfg_path, datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(cfg_path, backup)
    cfg.setdefault("paths", {})["data_root"] = os.path.abspath(data_root).replace("\\", "/")
    tmp = cfg_path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, cfg_path)
    say("config.json 已切换到新布局（paths.data_root = %s），原文件备份为 %s" % (cfg["paths"]["data_root"], backup))
    _reset_default_out_dir(say, sorted_in_place)


def _reset_default_out_dir(say, sorted_in_place=None):
    """A window setting that names the OLD default report folder would now mean
    "reports root" — drop it so the new default applies. Any other folder is
    kept (it is where the user wants reports); its old files were sorted in
    place by the plan (_plan_out_dir), which the notice says."""
    if _drop_out_dir_if(old_default_reports(), "旧的默认目录", say) is False:
        out = window_out_dir()
        head = ("注意：窗口里「报表输出到」是 %s —— 新布局里它是报表的根目录，monthly\\ quarterly\\ annual\\ custom\\ "
                "建在它下面。" % out)
        if sorted_in_place is None:
            say(head + "那里原有的旧报表请用 plan / copy 就地整理。")
        elif sorted_in_place == 0:
            say(head + "那里没有要整理的旧文件。")
        else:
            say(head + "那里原有的 %d 个旧文件已经按新布局整理（复制）到这些子目录里；旧文件先原样留着，"
                "remove-old 时和其他旧文件一起删。" % sorted_in_place)


def _drop_out_dir_if(default, what, say):
    """Remove the window's out_dir setting when it equals `default` (the window
    saves the folder it shows, even an untouched default, which would pin the
    other layout's folder). Returns True when removed, False when another folder
    is set, None when nothing is set."""
    from core import entrance
    p = entrance.SETTINGS
    doc = read_json(p)
    out = doc.get("out_dir")
    if not out:
        return None
    if os.path.normcase(os.path.abspath(out)) != os.path.normcase(os.path.abspath(default)):
        return False
    del doc["out_dir"]
    with io.open(p, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=2)
    say("窗口里「报表输出到」原来是%s，已清掉，改用当前布局的默认报表目录。" % what)
    return True


def undo(say=print):
    cfg_path = store_config.CONFIG_PATH
    with io.open(cfg_path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    if not (cfg.get("paths") or {}).get("data_root"):
        say("config.json 本来就是旧布局，不用改。")
        return
    shutil.copy2(cfg_path, "%s.bak_undo_%s" % (cfg_path, datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
    new_default = os.path.join(cfg["paths"]["data_root"], "reports")
    del cfg["paths"]["data_root"]
    if not cfg["paths"]:
        del cfg["paths"]
    with io.open(cfg_path, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    say("config.json 已切回旧布局。新目录里的文件没有动。")
    _drop_out_dir_if(new_default, "新布局的默认报表目录", say)


# ------------------------------------------------------------------ remove the old files

def remove_old(manifest_path, say=print):
    doc = read_json(manifest_path)
    if not doc.get("items"):
        raise Refused("清单 %s 读不出来或是空的。" % manifest_path)
    removed, kept = 0, []
    out = doc.get("out_dir")                 # the window's report folder, sorted in place
    roots = [legacy_downloads(), legacy_logs()] + ([out] if out else [])
    for it in doc["items"]:
        src, dst = it["src"], it["dst"]
        if not any(_under(src, r) and os.path.normcase(os.path.abspath(src)) != os.path.normcase(os.path.abspath(r))
                   for r in roots):
            kept.append((src, "不在旧目录里，不删"))
            continue
        if os.path.normcase(os.path.abspath(src)) == os.path.normcase(os.path.abspath(dst)):
            kept.append((src, "新旧位置相同"))
            continue
        if not os.path.exists(src):
            continue
        if not os.path.exists(dst) or sha256(dst) != it.get("dst_sha256", it["sha256"]):
            kept.append((src, "新目录里的副本不在或已不同"))
            continue
        if sha256(src) != it["sha256"]:
            kept.append((src, "旧文件在复制后又改过"))
            continue
        os.remove(src)
        removed += 1
    for top in (legacy_downloads(), legacy_logs()):
        _remove_empty_dirs(top)
    if out and os.path.isdir(out):           # the folder itself and the new tree stay
        for name in os.listdir(out):
            p = os.path.join(out, name)
            if os.path.isdir(p) and name not in NEW_REPORT_DIRS:
                _remove_empty_dirs(p)
    say("删掉了 %d 个旧文件。" % removed)
    for src, why in kept[:50]:
        say("  保留 %s：%s" % (src, why))
    if len(kept) > 50:
        say("  …… 另有 %d 个保留" % (len(kept) - 50))
    return removed, kept


def _remove_empty_dirs(top):
    if not os.path.isdir(top):
        return
    for dirpath, _dirs, _files in sorted(os.walk(top), key=lambda x: -len(x[0])):
        try:
            if not os.listdir(dirpath):
                os.rmdir(dirpath)
        except OSError:
            pass


# ------------------------------------------------------------------ command line

def summarize(plan, say=print):
    by = {}
    size = 0
    for kind, src, _dst in plan["items"]:
        n, s = by.get(kind, (0, 0))
        b = os.path.getsize(src)
        by[kind] = (n + 1, s + b)
        size += b
    names = {"raw": "下载文件", "report": "报表文件", "results": "月度结果", "work": "清洗中间文件",
             "logs": "日志", "old": "其他旧文件（放进 _旧文件）"}
    for kind in ("raw", "report", "results", "work", "logs", "old"):
        if kind in by:
            say("  %-24s %6d 个  %8.1f MB" % (names[kind], by[kind][0], by[kind][1] / 1e6))
    say("  合计 %.1f MB" % (size / 1e6))
    for n in plan["notes"]:
        say("  说明：%s" % n)
    for p, why in plan["unplaced"]:
        say("  ⚠ 放不进新布局（不会复制）：%s —— %s" % (p, why))
    return size


def main(argv=None):
    ap = argparse.ArgumentParser(description="把数据从旧布局搬到新布局（docs/DATA_LAYOUT.md）")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("plan", "copy"):
        p = sub.add_parser(name)
        p.add_argument("--data-root", required=True, help="新的数据根目录，例如 D:\\marketplace-data")
    p = sub.add_parser("remove-old")
    p.add_argument("--manifest", required=True)
    sub.add_parser("undo")
    a = ap.parse_args(argv)
    try:
        refuse_if_running()
        if a.cmd == "undo":
            undo()
            return 0
        if a.cmd == "remove-old":
            if layout.mode() != layout.V2:
                raise Refused("config.json 还是旧布局：先跑 copy，确认新布局能用之后再删旧文件。")
            remove_old(a.manifest)
            return 0
        check_target(a.data_root)
        current = layout.configured_data_root()
        if current and os.path.normcase(current) != os.path.normcase(os.path.abspath(a.data_root)):
            raise Refused("config.json 已经指向另一个数据目录 %s。" % current)
        plan = build_plan(os.path.abspath(a.data_root))
        print("从 %s 搬到 %s：" % (paths.ROOT, os.path.abspath(a.data_root)))
        summarize(plan)
        if a.cmd == "plan":
            print("\n这是预览，什么都没有写。确认无误后把 plan 换成 copy 再运行。")
            return 0
        check_space(plan, os.path.abspath(a.data_root))
        items, conflicts, unreadable = copy(plan, os.path.abspath(a.data_root))
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        mdir = os.path.join(os.path.abspath(a.data_root), "logs", "migration")
        os.makedirs(mdir, exist_ok=True)
        manifest = os.path.join(mdir, "manifest_%s.json" % stamp)
        with io.open(manifest, "w", encoding="utf-8") as fh:
            json.dump({"created": stamp, "root": paths.ROOT, "data_root": os.path.abspath(a.data_root),
                       "out_dir": plan["out_dir"], "items": items, "conflicts": conflicts,
                       "unreadable": unreadable, "unplaced": plan["unplaced"]},
                      fh, ensure_ascii=False, indent=1)
        print("\n复制并校验了 %d 个文件。清单：%s" % (len(items), manifest))
        if unreadable:
            print("⚠ %d 个文件读不了（多半是有人在 Excel 里开着，或共享盘断了一下），这次没复制，也没有切换布局：" % len(unreadable))
            for s, why in unreadable[:20]:
                print("  %s —— %s" % (s, why))
            print("关掉这些文件后再运行一次 copy：已经复制好的会直接跳过。")
            return 1
        if conflicts:
            print("⚠ %d 个文件在新目录里已有不同的版本，没有覆盖，也没有切换布局：" % len(conflicts))
            for s, d in conflicts[:20]:
                print("  %s → %s" % (s, d))
            return 1
        in_place = sum(1 for it in items if plan["out_dir"] and _under(it["src"], plan["out_dir"]))
        switch_config(os.path.abspath(a.data_root), sorted_in_place=in_place)
        print("\n旧目录一个文件都没动。新布局用过一次、确认没问题后，再运行：\n"
              "  python -m tools.migrate_layout remove-old --manifest \"%s\"" % manifest)
        return 0
    except Refused as e:
        print("没有执行：%s" % e)
        return 2


def check_space(plan, data_root):
    """Enough free space where the copies go: data_root's disk, and the window's
    report folder's own disk (a shared drive) for the files sorted in place."""
    out = plan.get("out_dir")
    need = {}
    for _kind, src, dst in plan["items"]:
        top = out if out and _under(dst, out) else data_root
        need[top] = need.get(top, 0) + os.path.getsize(src)
    for top, size in need.items():
        free = shutil.disk_usage(_existing_parent(top)).free
        if free < size * 1.1:
            raise Refused("%s 所在的磁盘剩余 %.1f MB，不够（需要约 %.1f MB）。" % (top, free / 1e6, size * 1.1 / 1e6))


def _existing_parent(p):
    p = os.path.abspath(p)
    while not os.path.exists(p):
        up = os.path.dirname(p)
        if up == p:                          # the drive itself is missing
            raise Refused("目录 %s 所在的磁盘不存在。" % p)
        p = up
    return p


if __name__ == "__main__":
    sys.exit(main())
