# -*- coding: utf-8 -*-
"""Final versions, the month folder and clean-up (docs/RESULTS_AND_PERIODS.md §4).

Called at the end of every report step (platforms/mercadolibre/run_reports.main,
core/platform_reports.ReportRunner.main) for the stores and month of that step:

    finalize.run(Unit(platform, market, month, out_dir, bases, store_name, combined_name),
                 stores, all_stores, build_combined, no_combined)

final(store, month, basis) = the newest countable version (core/results.final).

New layout (core/layout.mode() == V2); out_dir is layout.monthly_dir(month, market):

  * the step wrote each store workbook into <out_dir>/_versions/<version>/
    (version_dir); here the final version's workbook is copied to <out_dir> under
    the usual name; a store-basis without a final gets its newest version's
    workbook as <name>_未定稿.xlsx, removed once a final exists;
  * the country summary (全店汇总) is rebuilt from each store's final version
    (the newest where there is none) — the platform builds it (build_combined)
    from the versions' download folders; when a store is not final the name
    gets _未定稿 (and the other name is removed);
  * <out_dir>/版本说明.txt: per store and basis the final version, its grade, why
    every other version is not final, and what was cleaned;
  * clean-up of a store-month that has a final for every basis it produces:
    the other versions' workbooks in _versions/<version>/ (the folder once
    empty — it is shared by every store of the run), their results files, and
    the store-month's raw run folders that no final uses; logs/pending.json
    entries for deleted folders are dropped. A raw run folder newer than every
    final's that has no version at all is kept (it may still be downloading or
    not reported yet). Months without a final are never cleaned.

    Every deletion is checked first: inside the layout's roots, a name of the
    run / version pattern, and logged.

Legacy layout: nothing here touches files; results files and grades are written by
the runners, workbooks stay flat as before. run() only reports the finals.
"""
import collections
import io
import os
import re
import shutil
import tempfile

from core import layout
from core import results

VERSIONS = "_versions"
NOTES = "版本说明.txt"
DRAFT = "_未定稿"
RUN_NAME_RE = re.compile(r"^\d{8}_\d{6}.*_m\d{6}$")

Unit = collections.namedtuple("Unit", "platform market month out_dir bases store_name combined_name")
Unit.__doc__ = """One report step's unit: platform code, market ("ML-MX"), month, the
month folder (v2) or flat folder (legacy), the bases it produces (("bill",
"order") or ("main",)), store_name(store, month, basis) -> workbook file name,
combined_name(month, basis) -> the summary's file name."""


def is_v2(config_path=None):
    return layout.mode(config_path) == layout.V2


def version_dir(out_dir, version):
    """Where a version's store workbooks go in the new layout."""
    return os.path.join(out_dir, VERSIONS, version)


def draft_name(name):
    stem, ext = os.path.splitext(name)
    return stem + DRAFT + ext


def _inside(path, root):
    path, root = os.path.normcase(os.path.abspath(path)), os.path.normcase(os.path.abspath(root))
    return path != root and path.startswith(root.rstrip(os.sep) + os.sep)


def month_folder_ok(unit, config_path=None):
    """The month folder must look like <reports>/monthly/<month>/<PF-CC>: clean-up
    and copies happen nowhere else."""
    d = os.path.abspath(unit.out_dir)
    return (os.path.basename(d) == layout.market_folder(unit.market)
            and os.path.basename(os.path.dirname(d)) == unit.month
            and os.path.basename(os.path.dirname(os.path.dirname(d))) == "monthly")


# ------------------------------------------------------------------ decisions

def state(store, month, bases, config_path=None):
    """{basis: {"versions": [doc, …] newest first, "final": doc or None,
    "chosen": final or newest version with a workbook on disk, or None}}"""
    out = {}
    for b in bases:
        vs = results.versions(store, month, b, config_path)
        fin = next((v for v in vs if v["grade"]["level"] in results.COUNTABLE), None)
        chosen = fin or next((v for v in vs if v.get("workbook") and os.path.isfile(v["workbook"])), None)
        out[b] = {"versions": vs, "final": fin, "chosen": chosen}
    return out


def why_not_final(v, fin):
    g = v["grade"]
    if g["level"] in results.COUNTABLE:
        return "有更新的完整版本（%s）" % fin["version"] if fin else "完整"
    return "%s：%s" % (results.GRADE_CN.get(g["level"], g["level"]), "；".join(g.get("reasons") or []) or "—")


# ------------------------------------------------------------------ file operations

class Log(object):
    def __init__(self, say=print):
        self.say = say
        self.lines = []          # what was cleaned / copied, for 版本说明.txt

    def __call__(self, text, keep=False):
        self.say("  " + text)
        if keep:
            self.lines.append(text)


def _copy(src, dst, log):
    """Atomic copy (a workbook open in Excel makes the replace fail: reported)."""
    tmp = None
    try:
        fd, tmp = tempfile.mkstemp(prefix=".copy_", suffix=".xlsx", dir=os.path.dirname(dst))
        os.close(fd)
        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
        return True
    except (IOError, OSError) as e:
        log("[未能复制] %s → %s：%s（文件正被打开？关闭后重出）" % (src, os.path.basename(dst), e))
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        return False


def _remove_file(path, log, what):
    if not os.path.isfile(path):
        return False
    try:
        os.remove(path)
        log("删除%s：%s" % (what, path), keep=True)
        return True
    except OSError as e:
        log("[未能删除] %s：%s" % (path, e))
        return False


def remove_version_workbook(unit, doc, log):
    """A non-final version's store workbook, only inside <out_dir>/_versions/<version>/;
    the version folder goes once it is empty."""
    vdir = version_dir(unit.out_dir, doc["version"])
    wb = doc.get("workbook")
    if not results.VERSION_RE.match(doc["version"]) or not wb:
        return
    if os.path.normcase(os.path.dirname(os.path.abspath(wb))) != os.path.normcase(os.path.abspath(vdir)):
        log("[跳过] %s 不在 %s 里，不删" % (wb, vdir))
        return
    _remove_file(wb, log, "旧版本工作簿")
    try:
        if os.path.isdir(vdir) and not os.listdir(vdir):
            os.rmdir(vdir)
            log("删除空的版本目录：%s" % vdir, keep=True)
    except OSError as e:
        log("[未能删除] %s：%s" % (vdir, e))


def remove_results_file(doc, log, config_path=None):
    p = doc.get("_path")
    if not p or not results.FILE_RE.match(os.path.basename(p)):
        return
    if not _inside(p, layout.results_root(config_path)):
        log("[跳过] %s 不在结果目录里，不删" % p)
        return
    _remove_file(p, log, "旧版本结果文件")


def remove_run_folder(store, month, path, log, config_path=None):
    name = os.path.basename(os.path.normpath(path))
    if not RUN_NAME_RE.match(name) or layout.month_of_run(name) != month:
        log("[跳过] %s 不是本月的下载目录名，不删" % path)
        return False
    if not _inside(path, layout.store_dir(store, config_path)) or not _inside(path, layout.raw_root(config_path)):
        log("[跳过] %s 不在本店的下载目录里，不删" % path)
        return False
    try:
        shutil.rmtree(path)
        log("删除没有被最终版本使用的下载目录：%s" % path, keep=True)
        return True
    except OSError as e:
        log("[未能删除] %s：%s" % (path, e))
        return False


def drop_pending(folders, log, config_path=None):
    """Drop logs/pending.json entries whose run folder was deleted."""
    if not folders:
        return 0
    from core import pending_store
    path = os.path.join(layout.logs_dir(config_path), "pending.json")
    if not os.path.isfile(path):
        return 0
    gone = set(os.path.normcase(os.path.abspath(f)) for f in folders)
    entries = pending_store.load(path)
    keep = [e for e in entries if not (e.get("out_dir")
                                       and os.path.normcase(os.path.abspath(e["out_dir"])) in gone)]
    if len(keep) != len(entries):
        pending_store.save(keep, path)
        log("logs/pending.json：去掉 %d 条指向已删除目录的待取报表" % (len(entries) - len(keep)), keep=True)
    return len(entries) - len(keep)


# ------------------------------------------------------------------ the steps

def publish_store(unit, store, st, log):
    """Copy the final (or 未定稿) workbook of each basis into the month folder."""
    for b in unit.bases:
        s = st[b]
        name = unit.store_name(store, unit.month, b)
        dst, draft = os.path.join(unit.out_dir, name), os.path.join(unit.out_dir, draft_name(name))
        if s["final"]:
            src = s["final"].get("workbook")
            if src and os.path.isfile(src):
                if _copy(src, dst, log):
                    _remove_file(draft, log, "未定稿")
            else:
                log("[注意] %s 的最终版本 %s 的工作簿不在了：%s" % (store, s["final"]["version"], src))
        elif s["chosen"]:
            _copy(s["chosen"]["workbook"], draft, log)


def clean_store(unit, store, st, log, config_path=None):
    """Clean one store-month once every basis it produces has a final. Returns the
    deleted run folders."""
    produced = [b for b in unit.bases if st[b]["versions"]]
    if not produced or any(st[b]["final"] is None for b in produced):
        return []
    finals = [st[b]["final"] for b in produced]
    keep_versions = dict((b, st[b]["final"]["version"]) for b in produced)
    for b in produced:
        for v in st[b]["versions"]:
            if v["version"] == keep_versions[b]:
                continue
            remove_version_workbook(unit, v, log)
            remove_results_file(v, log, config_path)
    used = set(os.path.normcase(os.path.abspath(f)) for d in finals for f in d.get("source_folders") or ())
    reported = set(v["run"] for b in produced for v in st[b]["versions"])
    newest_final_run = max(d["run"] for d in finals)
    gone = []
    for name, path in layout.run_folders_for_month(store, unit.month, config_path=config_path):
        if os.path.normcase(os.path.abspath(path)) in used:
            continue
        if name not in reported and name > newest_final_run:
            log("保留 %s：比最终版本新，还没有出过报表" % path)
            continue
        if remove_run_folder(store, unit.month, path, log, config_path):
            gone.append(path)
    drop_pending(gone, log, config_path)
    return gone


def build_summaries(unit, states, build_combined, log):
    """The 全店汇总 of each basis from the chosen versions. Returns the build results."""
    out = []
    for b in unit.bases:
        pairs = [(store, st[b]["chosen"]) for store, st in states if st[b]["chosen"]]
        if not pairs:
            continue
        complete = all(st[b]["final"] is not None for _s, st in states if st[b]["versions"])
        name = unit.combined_name(unit.month, b)
        dst = os.path.join(unit.out_dir, name if complete else draft_name(name))
        other = os.path.join(unit.out_dir, draft_name(name) if complete else name)
        r = build_combined(b, pairs, dst)
        r.setdefault("out", dst)
        r["basis"] = b
        r["stores_final"] = [s for s, st in states if st[b]["final"] is not None]
        r["stores_draft"] = [s for s, st in states if st[b]["versions"] and st[b]["final"] is None]
        if r.get("outcome") in ("ok", "checks_failed", "formula_errors", "data_incomplete") and os.path.isfile(dst):
            _remove_file(other, log, "旧的汇总表" if complete else "已定稿的旧汇总表")
            log("汇总表%s：%s（%d 家店%s）" % ("" if complete else "（未定稿）", dst, len(pairs),
                                          "" if complete else "，未定稿：%s" % "、".join(r["stores_draft"])),
                keep=True)
        out.append(r)
    return out


def notes_text(unit, states, log_lines):
    lines = ["%s %s 版本说明（自动生成，每次出报表后重写）" % (unit.market, unit.month), "",
             "规则：每个版本都有等级；只有“完整”“完整·有提醒”可以定稿；可以定稿的版本里最新生成的那个就是最终版本。",
             "没有最终版本的店铺，月份文件夹里放的是它最新版本的工作簿，文件名带“%s”。" % DRAFT, ""]
    for store, st in states:
        for b in unit.bases:
            s = st[b]
            if not s["versions"]:
                continue
            head = "%s%s" % (store, " · " + results.BASIS_CN[b] if results.BASIS_CN.get(b) else "")
            fin = s["final"]
            if fin:
                g = fin["grade"]
                lines.append("%s：最终版本 %s（%s%s）" % (head, fin["version"], results.GRADE_CN[g["level"]],
                                                     "：" + "；".join(g["reasons"]) if g.get("reasons") else ""))
            else:
                lines.append("%s：还没有最终版本%s" % (head, "（月份文件夹里是 %s 的未定稿）" % s["chosen"]["version"]
                                                if s["chosen"] else ""))
            for v in s["versions"]:
                if fin and v["version"] == fin["version"]:
                    continue
                lines.append("    %s 不是最终版本：%s" % (v["version"], why_not_final(v, fin)))
    lines += ["", "本次整理："] + (["    " + x for x in log_lines] or ["    没有需要复制或删除的文件"])
    return "\n".join(lines) + "\n"


def run(unit, stores, all_stores=(), build_combined=None, no_combined=False, config_path=None, say=print):
    """Decide finals for the stores of this step; in v2 also publish, summarise,
    write 版本说明.txt and clean. Returns {"finals": {(store, basis): doc|None},
    "combined": [build results], "cleaned": [run folders]}; never raises on a
    file problem (each is logged)."""
    log = Log(say)
    order = list(dict.fromkeys(list(all_stores or ()) + list(stores)))
    states = [(s, state(s, unit.month, unit.bases, config_path)) for s in order]
    states = [(s, st) for s, st in states if s in stores or any(st[b]["versions"] for b in unit.bases)]
    out = {"finals": dict(((s, b), st[b]["final"]) for s, st in states for b in unit.bases),
           "combined": [], "cleaned": []}
    say("\n" + "=" * 62)
    say("最终版本（%s %s）" % (unit.market, unit.month))
    say("=" * 62)
    for s, st in states:
        for b in unit.bases:
            fin = st[b]["final"]
            say("  [最终] %-24s %-6s %s" % (s, results.BASIS_CN.get(b) or "", "%s（%s）" % (
                fin["version"], results.GRADE_CN[fin["grade"]["level"]]) if fin else
                ("还没有最终版本" if st[b]["versions"] else "没有任何版本")))
    if not is_v2(config_path):
        return out
    if not month_folder_ok(unit, config_path):
        log("[跳过整理] %s 不是 <报表>/monthly/%s/%s 这样的月份文件夹：不复制、不汇总、不清理"
            % (unit.out_dir, unit.month, layout.market_folder(unit.market)))
        return out
    os.makedirs(unit.out_dir, exist_ok=True)
    for s, st in states:
        publish_store(unit, s, st, log)
    if build_combined is not None and not no_combined:
        out["combined"] = build_summaries(unit, states, build_combined, log)
    for s, st in states:
        if s in stores:
            out["cleaned"] += clean_store(unit, s, st, log, config_path)
    # 版本说明 lists the versions as they are after the clean-up
    states = [(s, state(s, unit.month, unit.bases, config_path)) for s, _st in states]
    try:
        with io.open(os.path.join(unit.out_dir, NOTES), "w", encoding="utf-8") as fh:
            fh.write(notes_text(unit, states, log.lines))
    except (IOError, OSError) as e:
        log("[未能写入] %s：%s" % (NOTES, e))
    return out
