# -*- coding: utf-8 -*-
"""Move one machine's data from the old layout to the new one (docs/DATA_LAYOUT.md).

    python -m tools.migrate_layout plan   --data-root D:\\marketplace-data
    python -m tools.migrate_layout copy   --data-root D:\\marketplace-data
    python -m tools.migrate_layout remove-old --manifest D:\\marketplace-data\\logs\\migration\\manifest_….json
    python -m tools.migrate_layout undo

plan        lists what would go where, what cannot be placed, how much space it
            needs. Writes nothing.
copy        copies every file into the new tree, checks each copy (size and
            SHA-256), rewrites the moved folder paths inside logs/pending*.json,
            writes a manifest, then switches config.json to the new layout
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
            left in place.

Every command refuses to run while a monthly run is active. Run it between
monthly runs with the scheduled task paused.

What goes where (old -> new):
  downloads/<store>/<run>/              raw/<PF>/<CC>/<store>/<YYYY-MM or 其他>/<run>/
  downloads/data/reports/financial/*    reports/monthly/<YYYY-MM>/<PF>-<CC>/  (store and
                                        country workbooks, metrics), reports/monthly/<YYYY-MM>/
                                        (alert files), reports/_旧文件/ (anything else)
  downloads/data/processed/, mx_sales.duckdb, reports/<store>_<stamp>.xlsx
                                        work/mx_sales/…
  downloads/data/results/               results/
  downloads/data/<other>                work/_旧文件/<other>
  logs/                                 logs/
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
ALERT_RE = re.compile(r"^！报表异常_(?:.*_)?(?P<ym>\d{6})\.txt$")
MX_SALES_BOOK_RE = re.compile(r"^.+_\d{8}_\d{6}.*\.xlsx$")


class Refused(Exception):
    """A reason not to touch anything."""


# ------------------------------------------------------------------ helpers

def legacy_downloads():
    return os.path.join(paths.ROOT, "downloads")


def legacy_logs():
    return os.path.join(paths.ROOT, "logs")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def files_under(top):
    for dirpath, _dirs, files in os.walk(top):
        for f in files:
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
    return {"items": items, "unplaced": unplaced, "notes": notes, "folder_map": folder_map}


def _plan_reports(reports, data_root, store_markets, items, unplaced):
    """downloads/data/reports: financial/ (the flat report folder), the cleaning
    pipeline's workbooks, experiments and anything else."""
    for name in sorted(os.listdir(reports)):
        src = os.path.join(reports, name)
        if name == "financial" and os.path.isdir(src):
            for f in sorted(os.listdir(src)):
                p = os.path.join(src, f)
                if os.path.isdir(p):
                    items += [("old", x, os.path.join(data_root, "reports", "_旧文件", f, os.path.relpath(x, p)))
                              for x in files_under(p)]
                    continue
                dst = report_target(f, data_root, store_markets)
                items.append(("report", p, dst or os.path.join(data_root, "reports", "_旧文件", f)))
        elif os.path.isfile(src) and MX_SALES_BOOK_RE.match(name):
            items.append(("work", src, os.path.join(data_root, "work", "mx_sales", "reports", name)))
        elif os.path.isdir(src):
            items += [("old", x, os.path.join(data_root, "work", "_旧文件", "reports", name, os.path.relpath(x, src)))
                      for x in files_under(src)]
        else:
            items.append(("old", src, os.path.join(data_root, "work", "_旧文件", "reports", name)))


def report_target(fname, data_root, store_markets):
    """Where one file of the old flat report folder goes, or None (→ _旧文件)."""
    monthly = os.path.join(data_root, "reports", "monthly")
    m = ALERT_RE.match(fname)
    if m:
        ym = m.group("ym")
        return os.path.join(monthly, "%s-%s" % (ym[:4], ym[4:]), fname)
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
    point at a moved run folder are rewritten to the folder's new place. The
    manifest keeps both checksums (sha256 = the old file, dst_sha256 = its copy)."""
    done, conflicts = [], []
    for kind, src, dst in plan["items"]:
        digest = sha256(src)
        content = _pending_rewritten(src, plan["folder_map"], say) if _is_pending(kind, src) else None
        want = hashlib.sha256(content).hexdigest() if content is not None else digest
        if os.path.exists(dst):
            if sha256(dst) != want:
                conflicts.append((src, dst))
                continue
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
    return done, conflicts


def _is_pending(kind, path):
    name = os.path.basename(path)
    return kind == "logs" and name.startswith("pending") and name.endswith(".json")


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


def switch_config(data_root, say=print):
    """Write paths.data_root into config.json (backup first)."""
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
    _reset_default_out_dir(say)


def _reset_default_out_dir(say):
    """A window setting that names the OLD default report folder would now mean
    "reports root" — drop it so the new default applies. Any other folder is
    kept (it is where the user wants reports) and named in the output."""
    from core import entrance
    p = entrance.SETTINGS
    doc = read_json(p)
    out = doc.get("out_dir")
    if not out:
        return
    old_default = os.path.join(paths.ROOT, "downloads", "data", "reports", "financial")
    if os.path.normcase(os.path.abspath(out)) == os.path.normcase(os.path.abspath(old_default)):
        del doc["out_dir"]
        with io.open(p, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False, indent=2)
        say("窗口里「报表输出到」原来是旧的默认目录，已清掉，改用新的报表目录。")
    else:
        say("注意：窗口里「报表输出到」是 %s —— 新布局里它是报表的根目录，monthly\\ quarterly\\ annual\\ "
            "会建在它下面；那里原有的旧报表没有搬。" % out)


def undo(say=print):
    cfg_path = store_config.CONFIG_PATH
    with io.open(cfg_path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    if not (cfg.get("paths") or {}).get("data_root"):
        say("config.json 本来就是旧布局，不用改。")
        return
    shutil.copy2(cfg_path, "%s.bak_undo_%s" % (cfg_path, datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
    del cfg["paths"]["data_root"]
    if not cfg["paths"]:
        del cfg["paths"]
    with io.open(cfg_path, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    say("config.json 已切回旧布局。新目录里的文件没有动。")


# ------------------------------------------------------------------ remove the old files

def remove_old(manifest_path, say=print):
    doc = read_json(manifest_path)
    if not doc.get("items"):
        raise Refused("清单 %s 读不出来或是空的。" % manifest_path)
    removed, kept = 0, []
    roots = [os.path.normcase(os.path.abspath(r)) for r in (legacy_downloads(), legacy_logs())]
    for it in doc["items"]:
        src, dst = it["src"], it["dst"]
        if not any(os.path.normcase(os.path.abspath(src)).startswith(r + os.sep) for r in roots):
            kept.append((src, "不在旧目录里，不删"))
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
        size = summarize(plan)
        if a.cmd == "plan":
            print("\n这是预览，什么都没有写。确认无误后把 plan 换成 copy 再运行。")
            return 0
        free = shutil.disk_usage(os.path.splitdrive(os.path.abspath(a.data_root))[0] + os.sep
                                 if os.name == "nt" else _existing_parent(a.data_root)).free
        if free < size * 1.1:
            raise Refused("目标磁盘剩余 %.1f MB，不够（需要约 %.1f MB）。" % (free / 1e6, size * 1.1 / 1e6))
        items, conflicts = copy(plan, os.path.abspath(a.data_root))
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        mdir = os.path.join(os.path.abspath(a.data_root), "logs", "migration")
        os.makedirs(mdir, exist_ok=True)
        manifest = os.path.join(mdir, "manifest_%s.json" % stamp)
        with io.open(manifest, "w", encoding="utf-8") as fh:
            json.dump({"created": stamp, "root": paths.ROOT, "data_root": os.path.abspath(a.data_root),
                       "items": items, "conflicts": conflicts, "unplaced": plan["unplaced"]},
                      fh, ensure_ascii=False, indent=1)
        print("\n复制并校验了 %d 个文件。清单：%s" % (len(items), manifest))
        if conflicts:
            print("⚠ %d 个文件在新目录里已有不同的版本，没有覆盖，也没有切换布局：" % len(conflicts))
            for s, d in conflicts[:20]:
                print("  %s → %s" % (s, d))
            return 1
        switch_config(os.path.abspath(a.data_root))
        print("\n旧目录一个文件都没动。新布局用过一次、确认没问题后，再运行：\n"
              "  python -m tools.migrate_layout remove-old --manifest \"%s\"" % manifest)
        return 0
    except Refused as e:
        print("没有执行：%s" % e)
        return 2


def _existing_parent(p):
    p = os.path.abspath(p)
    while not os.path.exists(p):
        p = os.path.dirname(p)
    return p


if __name__ == "__main__":
    sys.exit(main())
