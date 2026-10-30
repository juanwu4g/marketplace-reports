# -*- coding: utf-8 -*-
"""Monthly results: the figures of every monthly report, saved as data.

Every time a store's monthly workbook is built, one JSON file is written next to
the others of that store-month (core/layout.results_dir):

    <results>/<PF>/<CC>/<store>/<YYYY-MM>/<version>.<basis>.json

  version  "<run>__<built>": <run> is the download run folder the report was
           built from (e.g. 20261021_233002_m202609), <built> the time the
           report was generated (YYYYMMDD_HHMMSS). A report rebuilt from the same
           download later is a new version.
  basis    "bill" / "order" for MercadoLibre (账单基准 / 订单基准), "main" for
           every other platform (one basis).

Each file carries a grade (GRADES). The final version of a store-month-basis is
the newest one (by "built_at") whose grade counts (COUNTABLE) — no person picks
it. Period reports (quarterly, annual, custom) are built only from final
versions. docs/RESULTS_AND_PERIODS.md describes the fields and the rules.

This module is the read side and the file format; the report runners write the
files (write()), and core/finalize decides finals, copies and cleans.
"""
import datetime
import io
import json
import os
import re

from core import layout

SCHEMA = 1
BASES = ("bill", "order", "main")
BASIS_CN = {"bill": "账单基准", "order": "订单基准", "main": ""}

COMPLETE, COMPLETE_WARN, PROVISIONAL, GAPS, FAILED = (
    "complete", "complete_warn", "provisional", "gaps", "failed")
GRADES = (COMPLETE, COMPLETE_WARN, PROVISIONAL, GAPS, FAILED)
COUNTABLE = (COMPLETE, COMPLETE_WARN)
GRADE_CN = {COMPLETE: "完整", COMPLETE_WARN: "完整·有提醒", PROVISIONAL: "暂定", GAPS: "有缺口",
            FAILED: "失败"}

REQUIRED = ("schema", "platform", "market", "country", "currency", "decimals", "store", "month",
            "basis", "run", "built_at", "version", "grade", "checks", "pl")
FILE_RE = re.compile(r"^(?P<version>.+__\d{8}_\d{6})\.(?P<basis>bill|order|main)\.json$")
MONTH_RE = re.compile(r"^\d{4}-\d{2}$")


# ------------------------------------------------------------------ names

def official_basis(platform):
    """The basis a period report uses unless told otherwise."""
    return "bill" if platform == "ML" else "main"


def built_stamp(when=None):
    return (when or datetime.datetime.now()).strftime("%Y%m%d_%H%M%S")


def version_id(run, built):
    """run: the download run folder's name (a path is accepted); built: a
    YYYYMMDD_HHMMSS stamp or a datetime."""
    if isinstance(built, datetime.datetime):
        built = built_stamp(built)
    return "%s__%s" % (os.path.basename(os.path.normpath(run)), built)


def path_for(store, month, version, basis, config_path=None):
    return os.path.join(layout.results_dir(store, month, config_path), "%s.%s.json" % (version, basis))


# ------------------------------------------------------------------ write

def validate(doc):
    """Problems with a result document ([] when it is well formed)."""
    out = ["缺字段 %s" % k for k in REQUIRED if k not in doc]
    if out:
        return out
    if doc["schema"] != SCHEMA:
        out.append("schema %r（应为 %d）" % (doc["schema"], SCHEMA))
    if doc["basis"] not in BASES:
        out.append("basis %r" % doc["basis"])
    if not MONTH_RE.match(str(doc["month"])):
        out.append("month %r" % doc["month"])
    g = doc.get("grade") or {}
    if g.get("level") not in GRADES:
        out.append("grade.level %r" % g.get("level"))
    if not isinstance(doc["pl"], list) or any("key" not in r or "value" not in r for r in doc["pl"]):
        out.append("pl 必须是 [{key, label, original, kind, value}, …]")
    return out


def write(doc, config_path=None):
    """Write one result file (atomically). Returns its path. Raises ValueError
    on a malformed document — a bad file must not become a final version."""
    bad = validate(doc)
    if bad:
        raise ValueError("结果文件不合格：%s" % "；".join(bad))
    path = path_for(doc["store"], doc["month"], doc["version"], doc["basis"], config_path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1, sort_keys=True)
    os.replace(tmp, path)
    return path


# ------------------------------------------------------------------ read

def load(path):
    with io.open(path, encoding="utf-8") as fh:
        doc = json.load(fh)
    doc["_path"] = path
    return doc


def versions(store, month, basis=None, config_path=None):
    """Every version of a store-month (one basis or all), newest built first.
    Unreadable or malformed files are left out; problems() lists them."""
    pf, cc = layout.store_market(store)
    return versions_at(pf, cc, store, month, basis, config_path)


def final(store, month, basis, config_path=None):
    """The final version: the newest whose grade counts; None when there is none."""
    for doc in versions(store, month, basis, config_path):
        if doc["grade"]["level"] in COUNTABLE:
            return doc
    return None


def problems(store, month, config_path=None):
    """Files in a store-month folder that cannot be used, with the reason."""
    d = layout.results_dir(store, month, config_path)
    out = []
    if not os.path.isdir(d):
        return out
    for name in sorted(os.listdir(d)):
        if not name.endswith(".json"):
            continue
        if not FILE_RE.match(name):
            out.append((name, "文件名不合规"))
            continue
        try:
            bad = validate(load(os.path.join(d, name)))
        except (ValueError, IOError, OSError) as e:
            bad = ["读不出来：%s" % e]
        if bad:
            out.append((name, "；".join(bad)))
    return out


def stores(platform=None, country=None, config_path=None):
    """Store names that have any results, optionally of one platform / country."""
    root = layout.results_root(config_path)
    out = []
    if not os.path.isdir(root):
        return out
    for pf in sorted(os.listdir(root)):
        if platform and pf != platform:
            continue
        for cc in sorted(os.listdir(os.path.join(root, pf))):
            if country and cc != country:
                continue
            base = os.path.join(root, pf, cc)
            out += sorted(s for s in os.listdir(base) if os.path.isdir(os.path.join(base, s)))
    return out


def months(store, config_path=None):
    """Months ('YYYY-MM') with any result for the store, ascending."""
    base = os.path.dirname(layout.results_dir(store, "2000-01", config_path))
    if not os.path.isdir(base):
        return []
    return sorted(m for m in os.listdir(base) if MONTH_RE.match(m))


# ------------------------------------------------------------------ periods

def months_between(first, last):
    """['2026-07', '2026-08', '2026-09'] for ('2026-07', '2026-09'), inclusive."""
    y, m = int(first[:4]), int(first[5:7])
    ly, lm = int(last[:4]), int(last[5:7])
    if (y, m) > (ly, lm):
        raise ValueError("期间起点 %s 晚于终点 %s" % (first, last))
    out = []
    while (y, m) <= (ly, lm):
        out.append("%04d-%02d" % (y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def quarter_months(quarter):
    """'2026-Q3' (or '2026Q3') -> ['2026-07', '2026-08', '2026-09']."""
    m = re.match(r"^(\d{4})-?Q([1-4])$", str(quarter).strip().upper())
    if not m:
        raise ValueError("季度写法应为 2026-Q3：%r" % quarter)
    y, q = int(m.group(1)), int(m.group(2))
    return months_between("%04d-%02d" % (y, 3 * q - 2), "%04d-%02d" % (y, 3 * q))


def year_months(year):
    return months_between("%04d-01" % int(year), "%04d-12" % int(year))


def quarter_of(month):
    """'2026-09' -> '2026-Q3'."""
    return "%s-Q%d" % (month[:4], (int(month[5:7]) - 1) // 3 + 1)


# ------------------------------------------------------------------ by results folder
#
# versions() / final() find a store's folder through config.json (its market).
# A period report also has to read stores that config.json no longer lists (a
# store closed after the quarter); these read the same files by the folder the
# results tree itself gives (<results>/<PF>/<CC>/<store>/), with the same rule.

def versions_at(platform, country, store, month, basis=None, config_path=None):
    """versions() for a store found in the results tree under <PF>/<CC>."""
    d = os.path.join(layout.results_root(config_path), platform, country, store, month)
    if not os.path.isdir(d):
        return []
    out = []
    for name in os.listdir(d):
        m = FILE_RE.match(name)
        if not m or (basis and m.group("basis") != basis):
            continue
        try:
            doc = load(os.path.join(d, name))
        except (ValueError, IOError, OSError):
            continue
        if not validate(doc):
            out.append(doc)
    return sorted(out, key=lambda x: (x["built_at"], x["version"]), reverse=True)


def has_files_at(platform, country, store, month, basis=None, config_path=None):
    """Is there any results file (by its name) for this store-month and basis?
    Cheap: lists the folder, reads nothing."""
    d = os.path.join(layout.results_root(config_path), platform, country, store, month)
    if not os.path.isdir(d):
        return False
    for name in os.listdir(d):
        m = FILE_RE.match(name)
        if m and (not basis or m.group("basis") == basis):
            return True
    return False


def final_at(platform, country, store, month, basis, config_path=None):
    """final() for a store found in the results tree under <PF>/<CC>."""
    for doc in versions_at(platform, country, store, month, basis, config_path):
        if doc["grade"]["level"] in COUNTABLE:
            return doc
    return None


def period_label(kind, months_):
    """The period's name in folders and file names: '2026-Q3', '2026',
    '2026-07_2026-11' (kind 'quarterly' / 'annual' / 'custom')."""
    if kind == "quarterly":
        return quarter_of(months_[0])
    if kind == "annual":
        return months_[0][:4]
    return "%s_%s" % (months_[0], months_[-1])


# ------------------------------------------------------------------ helpers for the writers
#
# Added for the report runners (they build the documents) and core/finalize.
# Nothing above changes: names, format and the read side stay as they were.

VERSION_RE = re.compile(r"^(?P<run>.+)__(?P<built>\d{8}_\d{6})$")
_CODE = {}


def version_parts(version):
    """'<run>__<YYYYMMDD_HHMMSS>' -> (run, built stamp); (None, None) when it is not one."""
    m = VERSION_RE.match(str(version or ""))
    return (m.group("run"), m.group("built")) if m else (None, None)


RUN_TIME_RE = re.compile(r"^(\d{8})_(\d{6})")


def run_time_text(run):
    """The download time a run folder's name carries, as "YYYY-MM-DD HH:MM":
    '20261021_233002_m202609' -> '2026-10-21 23:30' (a path or a version id is
    accepted: the basename is read). "" when the name does not start with a
    YYYYMMDD_HHMMSS stamp."""
    m = RUN_TIME_RE.match(os.path.basename(os.path.normpath(str(run or ""))))
    if not m:
        return ""
    d, t = m.groups()
    return "%s-%s-%s %s:%s" % (d[:4], d[4:6], d[6:8], t[:2], t[2:4])


def stamp_info(grade, version, run, basis, built_at=None):
    """What core/report_kit.stamp_grade writes on sheet ①: from a grade
    ({"level", "reasons"}), the version id, the download run (folder or name)
    and the basis ("bill" / "order" / "main"). final = the grade counts
    (COUNTABLE); at stamping time the version is the newest of its
    store-month-basis, so that is the whole rule."""
    level = (grade or {}).get("level") or ""
    built_at = built_at or datetime.datetime.now()
    return {"level": level, "reasons": [str(r) for r in (grade or {}).get("reasons") or ()],
            "version": str(version or ""), "downloaded_at": run_time_text(run),
            "built_at": built_at.strftime("%Y-%m-%d %H:%M"), "basis_cn": BASIS_CN.get(basis, ""),
            "final": level in COUNTABLE}


def code_info():
    """{"commit", "fingerprint"} of the code that builds the report (core/runlog);
    computed once per process — the fingerprint reads every code file."""
    if not _CODE:
        from core import runlog
        try:
            _CODE.update(commit=runlog.git_commit(), fingerprint=runlog.code_fingerprint())
        except Exception:                     # noqa: BLE001 — a results file must not fail on this
            _CODE.update(commit=None, fingerprint=None)
    return dict(_CODE)


def pl_rows(spec):
    """Evaluate a P&L the way the workbook's formulas do.

    spec: [{"key", "label", "original", "kind", "value", "formula", "name"}] in row
    order. "d" / "in" rows carry their value; "sum" / "f" rows a formula as in
    core/report_family.PLRow: a tuple of row names added up ("-name" subtracts) or
    a Span(first, last) summing every row between the two (inclusive). "name" is
    what formulas refer to (default: the key). Returns the results file's "pl":
    [{"key", "label", "original", "kind", "value"}].
    """
    rows = [dict(r) for r in spec]
    at = {}
    for i, r in enumerate(rows):
        at[r.get("name") or r["key"]] = i
    vals = {}

    def value(i, seen=()):
        if i in vals:
            return vals[i]
        r = rows[i]
        f = r.get("formula")
        if r["kind"] in ("sum", "f"):
            if not f:
                raise ValueError("P&L row %r (%s) needs a formula" % (r["key"], r["kind"]))
            if i in seen:
                raise ValueError("P&L formula loop at %r" % r["key"])
            if hasattr(f, "first") and hasattr(f, "last"):
                v = sum(value(j, seen + (i,)) for j in range(at[f.first], at[f.last] + 1))
            else:
                v = 0.0
                for name in f:
                    neg = name.startswith("-")
                    x = value(at[name[1:] if neg else name], seen + (i,))
                    v += -x if neg else x
        else:
            v = float(r.get("value") or 0.0)
        vals[i] = v
        return v

    return [{"key": r["key"], "label": r["label"], "original": r.get("original") or "", "kind": r["kind"],
             "value": value(i)} for i, r in enumerate(rows)]


def sku_cost(cost, sku):
    """A SKU's goods cost from a core/sku_costs.Result, as ② shows 商品采购成本
    (positive); None when there is no cost table or the SKU has no usable cost."""
    if cost is None:
        return None
    ln = cost.line(sku)
    if ln is None or getattr(ln, "unit_cost", None) is None:
        return None
    return float(ln.cost)


def document(platform, market, country, currency, decimals, store, month, basis, version, built_at,
             source_folders, workbook, grade, checks=(), missing=(), provisional=(), data_incomplete=False,
             figures=None):
    """A results document (schema 1) from what a report runner has in hand.
    figures: {"pl", "memo", "fees", "returns", "sku", "inventory"} — the workbook's
    own numbers; left out ([] / None) when the version has none (failed)."""
    figures = figures or {}
    run, _stamp = version_parts(version)
    if isinstance(built_at, datetime.datetime):
        built_at = built_at.strftime("%Y-%m-%dT%H:%M:%S")
    return plain({
        "schema": SCHEMA, "platform": platform, "market": market, "country": country,
        "currency": currency, "decimals": int(decimals), "store": store, "month": month, "basis": basis,
        "run": run or "", "built_at": built_at, "version": version,
        "source_folders": [os.path.abspath(f) for f in source_folders or () if f],
        "workbook": os.path.abspath(workbook) if workbook else None,
        "code": code_info(), "grade": grade, "checks": list(checks or ()),
        "missing": [{"form": r.get("form"), "cause": r.get("cause"), "in_report": r.get("in_report")}
                    for r in missing or () if r and r.get("form")],
        "flags": {"provisional": list(provisional or ()), "data_incomplete": bool(data_incomplete)},
        "pl": list(figures.get("pl") or ()), "memo": list(figures.get("memo") or ()),
        "fees": list(figures.get("fees") or ()), "returns": list(figures.get("returns") or ()),
        "sku": list(figures.get("sku") or ()), "inventory": figures.get("inventory"),
    })


def plain(obj):
    """obj as plain JSON values (numpy numbers, dates and the like made plain)."""
    return json.loads(json.dumps(obj, ensure_ascii=False, default=_plain))


def _plain(x):
    if hasattr(x, "item"):
        try:
            return x.item()                   # numpy scalars
        except (ValueError, TypeError):
            pass
    if hasattr(x, "isoformat"):
        return x.isoformat()
    return str(x)
