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
    d = layout.results_dir(store, month, config_path)
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
