# -*- coding: utf-8 -*-
"""Where the data lives: raw downloads, reports, monthly results, work files, logs.

Every module asks here; nothing else joins "downloads" or "logs" onto ROOT.

Two layouts, chosen by config.json:

  legacy   (no "paths.data_root" in config.json — how every machine ran until
           2026-10, and still runs until tools/migrate_layout.py has moved it)

      <ROOT>/downloads/<store>/<run>/                raw downloads
      <ROOT>/downloads/data/reports/financial/       every report, flat
      <ROOT>/downloads/data/results/…                monthly results (same tree as v2)
      <ROOT>/downloads/data/                         cleaning output (processed/, mx_sales.duckdb)
      <ROOT>/logs/

  v2       ("paths": {"data_root": "D:/marketplace-data"} in config.json)

      <data>/raw/<PF>/<CC>/<store>/<YYYY-MM>/<run>/  raw downloads
      <data>/reports/monthly/<YYYY-MM>/<PF>-<CC>/    the month's final reports
      <data>/reports/quarterly/<YYYY>-Q<n>/<PF>-<CC>/
      <data>/reports/annual/<YYYY>/<PF>-<CC>/
      <data>/reports/custom/<YYYY-MM>_<YYYY-MM>/<PF>-<CC>/
      <data>/results/<PF>/<CC>/<store>/<YYYY-MM>/    monthly results, one file per version
      <data>/work/                                   cleaning output and other intermediates
      <data>/logs/

<run> is the run folder's name, unchanged between the layouts: the batch stamp
"YYYYMMDD_HHMMSS" plus "_m<YYYYMM>" when the run is for one month. The name is
also the version id that ties a download, its reports and its results together.

The window's 「报表输出到」 (gui_settings.json out_dir) is the flat report folder
in the legacy layout and the reports root in v2 (see reports_root).

docs/DATA_LAYOUT.md is the reader's version of this docstring.
"""
import io
import json
import os
import re

from core import paths

LEGACY = "legacy"
V2 = "v2"

RUN_MONTH_RE = re.compile(r"_m(\d{4})(\d{2})(?:$|_)")
NO_MONTH = "其他"                # v2 month folder for a run that is not about one month


# ------------------------------------------------------------------ config

def _config_path():
    from core import store_config            # its CONFIG_PATH is what tests point elsewhere
    return store_config.CONFIG_PATH


def configured_data_root(config_path=None):
    """paths.data_root from config.json (absolute), or None for the legacy layout.

    A config.json that exists but cannot be read raises StoreConfigError: data
    must never be written to a guessed place."""
    path = config_path or _config_path()
    if not os.path.exists(path):
        return None
    try:
        with io.open(path, encoding="utf-8") as fh:
            cfg = json.load(fh)
    except (ValueError, IOError, OSError) as e:
        from core.store_config import StoreConfigError
        raise StoreConfigError("config.json 读不出来（%s），不知道数据放在哪里" % e)
    node = (cfg.get("paths") or {}).get("data_root") if isinstance(cfg, dict) else None
    if not node or not isinstance(node, str) or not node.strip():
        return None
    p = os.path.expanduser(node.strip())
    return os.path.normpath(p if os.path.isabs(p) else os.path.join(paths.ROOT, p))


def mode(config_path=None):
    return V2 if configured_data_root(config_path) else LEGACY


def data_root(config_path=None):
    """v2: the configured folder; legacy: <ROOT>/downloads/data (where the
    cleaning output and the flat reports have always been)."""
    return configured_data_root(config_path) or os.path.join(paths.ROOT, "downloads", "data")


# ------------------------------------------------------------------ months and markets

def month_of_run(name):
    """'20261021_233002_m202609' -> '2026-09'; None when the name carries no month."""
    m = RUN_MONTH_RE.search(os.path.basename(os.path.normpath(name)))
    return "%s-%s" % (m.group(1), m.group(2)) if m else None


def market_folder(market):
    """'MX' / 'ML-MX' -> 'ML-MX'; 'FA-CL' -> 'FA-CL' (the report folder of a market)."""
    from core import markets
    return markets.normalize(market)


def store_market(store):
    """(platform code, country) of a store, from config.json (as store_config says)."""
    from core import store_config
    return store_config.market_parts(store)


# ------------------------------------------------------------------ raw downloads

def raw_root(config_path=None):
    root = configured_data_root(config_path)
    return os.path.join(root, "raw") if root else os.path.join(paths.ROOT, "downloads")


def store_dir(store, config_path=None):
    """Where a store's run folders live (legacy) or its month folders (v2)."""
    root = configured_data_root(config_path)
    if not root:
        return os.path.join(paths.ROOT, "downloads", store)
    pf, cc = store_market(store)
    return os.path.join(root, "raw", pf, cc, store)


def run_folder(store, name, config_path=None):
    """Absolute folder for one download run named `name` (see module docstring).

    Absolute on purpose: the browser resolves a relative download path against
    its own working directory and the files vanish (see run_downloads)."""
    if not configured_data_root(config_path):
        return os.path.abspath(os.path.join(paths.ROOT, "downloads", store, name))
    return os.path.abspath(os.path.join(store_dir(store, config_path), month_of_run(name) or NO_MONTH, name))


def run_folders(store, base=None, config_path=None):
    """Every run folder of a store, newest first: [(name, absolute path)].

    `base` is a legacy-shaped root (<base>/<store>/<run>/) — what tests and the
    old --downloads options pass; without it the configured layout is used.
    Only folders whose name starts with a year ("20…") count, as before."""
    if base is not None or not configured_data_root(config_path):
        parent = os.path.join(base if base is not None else os.path.join(paths.ROOT, "downloads"), store)
        found = _runs_in(parent)
    else:
        found = []
        sd = store_dir(store, config_path)
        if os.path.isdir(sd):
            for month in os.listdir(sd):
                found += _runs_in(os.path.join(sd, month))
    return sorted(found, key=lambda x: x[0], reverse=True)


def _runs_in(parent):
    if not os.path.isdir(parent):
        return []
    return [(x, os.path.join(parent, x)) for x in os.listdir(parent)
            if x[:2] == "20" and os.path.isdir(os.path.join(parent, x))]


def run_folders_for_month(store, month, base=None, config_path=None):
    """The run folders of one month ('2026-09'), newest first: those named
    *_m<YYYYMM> (v2 keeps them in the month folder)."""
    tag = "_m%s%s" % (month[:4], month[5:7])
    return [(n, p) for n, p in run_folders(store, base, config_path)
            if RUN_MONTH_RE.search(n) and month_of_run(n) == month and tag in n]


# ------------------------------------------------------------------ reports

def reports_root(override=None, config_path=None):
    """legacy: the flat report folder (the window's choice, else the default);
    v2: the root under which monthly/, quarterly/, annual/, custom/ live (the
    window's choice, else <data>/reports)."""
    if override:
        return os.path.abspath(override)
    root = configured_data_root(config_path)
    return os.path.join(root, "reports") if root else os.path.join(paths.ROOT, "downloads", "data", "reports",
                                                                    "financial")


def monthly_dir(month, market, override=None, config_path=None):
    """The folder a market's monthly reports go to. Flat (= reports_root) in
    the legacy layout."""
    root = reports_root(override, config_path)
    if not configured_data_root(config_path):
        return root
    return os.path.join(root, "monthly", month, market_folder(market))


def alert_dir(month, override=None, config_path=None):
    """Where 「！报表异常_…」 files go: next to the reports (legacy), the month
    folder (v2)."""
    root = reports_root(override, config_path)
    if not configured_data_root(config_path):
        return root
    return os.path.join(root, "monthly", month)


def period_dir(kind, period, market, override=None, config_path=None):
    """kind: 'quarterly' (period '2026-Q3'), 'annual' ('2026'), 'custom'
    ('2026-07_2026-11'). Legacy: <flat report folder>/期间报表/<kind>/<period>/<market>."""
    if kind not in ("quarterly", "annual", "custom"):
        raise ValueError("period kind %r" % kind)
    root = reports_root(override, config_path)
    if not configured_data_root(config_path):
        root = os.path.join(root, "期间报表")
    return os.path.join(root, kind, period, market_folder(market))


# ------------------------------------------------------------------ results, work, logs

def results_root(config_path=None):
    return os.path.join(data_root(config_path), "results")


def results_dir(store, month, config_path=None):
    """<results>/<PF>/<CC>/<store>/<YYYY-MM>/ — the same tree in both layouts."""
    pf, cc = store_market(store)
    return os.path.join(results_root(config_path), pf, cc, store, month)


def work_dir(config_path=None):
    root = configured_data_root(config_path)
    return os.path.join(root, "work") if root else os.path.join(paths.ROOT, "downloads", "data")


def logs_dir(config_path=None):
    """Never raises: logs are needed to report that config.json is broken."""
    try:
        root = configured_data_root(config_path)
    except Exception:                         # noqa: BLE001
        root = None
    return os.path.join(root, "logs") if root else os.path.join(paths.ROOT, "logs")


def mx_sales_paths(config_path=None):
    """The Mexican cleaning pipeline's folders (its MX_SALES_* settings):
    raw tree it walks, Parquet output, DuckDB file, per-store workbooks."""
    root = configured_data_root(config_path)
    if not root:
        legacy = os.path.join(paths.ROOT, "downloads")
        return {"raw": legacy, "processed": os.path.join(legacy, "data", "processed"),
                "warehouse": os.path.join(legacy, "data", "mx_sales.duckdb"),
                "reports": os.path.join(legacy, "data", "reports")}
    work = os.path.join(root, "work", "mx_sales")
    return {"raw": os.path.join(root, "raw", "ML", "MX"), "processed": os.path.join(work, "processed"),
            "warehouse": os.path.join(work, "mx_sales.duckdb"), "reports": os.path.join(work, "reports")}


def describe(config_path=None):
    """One line per folder, for the window and the run log header."""
    m = mode(config_path)
    return [("数据布局", "新布局（数据根目录 %s）" % configured_data_root(config_path) if m == V2
             else "旧布局（数据在代码目录的 downloads\\ 和 logs\\ 下）"),
            ("下载", raw_root(config_path)), ("报表", reports_root(None, config_path)),
            ("月度结果", results_root(config_path)), ("日志", logs_dir(config_path))]
