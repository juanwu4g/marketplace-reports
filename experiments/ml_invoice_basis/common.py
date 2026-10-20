# -*- coding: utf-8 -*-
"""Small shared pieces: the store's site per download folder, folder discovery,
output names, id keys.

The MercadoLibre report reads its country profile from module globals
(platforms.mercadolibre.report.site, switched by set_site). The experiment
takes the site from each folder's run_meta.json (like report_build.sku_list)
and always restores the previous site afterwards, so importing or running it
never leaves another country switched on."""

import contextlib
import glob
import os
import re

import pandas as pd

from core.paths import ROOT
from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import file_month_token
from platforms.mercadolibre.report.loading import load_run_meta

REPORT_SUFFIX = "账单口径实验"
STABILITY_SUFFIX = "口径稳定性"
DEFAULT_OUT_DIR = os.path.join(ROOT, "downloads", "data", "reports", "experiments")
DEFAULT_BASE = os.path.join(ROOT, "downloads")

FAC_GLOB = "Reporte_Facturacion_MercadoLibre_*.xlsx"
NC_GLOB = "Reporte_Notas_Credito_MercadoLibre_*.xlsx"


class ExperimentError(Exception):
    """A problem the person running the command must fix (Chinese message)."""


def folder_site_code(folder):
    """The folder's site from run_meta.json; old folders without it are Mexico."""
    return str(load_run_meta(folder).get("site") or "mx").lower()


@contextlib.contextmanager
def folder_site(folder):
    """Switch the MercadoLibre report to this folder's country; restore after."""
    before = site.SITE
    code = folder_site_code(folder)
    try:
        try:
            site.set_site(code)
        except ValueError:
            raise ExperimentError("下载目录 %s 的 run_meta.json 写的站点是 %r，不认识（只认 mx / cl / pe）。"
                                  % (folder, code))
        yield code
    finally:
        site.set_site(before)


def key(v):
    """An id as a clean string ('2000018…'); '' for missing values.

    pandas 3 keeps NaN through astype(str), so 'nan' and NaN both mean empty."""
    if v is None:
        return ""
    if not isinstance(v, str):
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        if isinstance(v, float) and v == int(v):
            return "%d" % v
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return "" if s.lower() in ("nan", "none", "nat") else s


def parse_month(month):
    m = re.match(r"^(\d{4})-(\d{2})$", str(month or ""))
    if not m or not 1 <= int(m.group(2)) <= 12:
        raise ExperimentError("月份要写成 YYYY-MM，例如 2026-09；收到的是 %r。" % month)
    return int(m.group(1)), int(m.group(2))


def has_invoice(folder, year, month):
    """Does the folder hold the Facturación named for this month?"""
    for p in glob.glob(os.path.join(folder, FAC_GLOB)):
        if not os.path.basename(p).startswith("~$") \
                and file_month_token(os.path.basename(p)) == (year, month):
            return True
    return False


def store_root(store, base=None):
    return os.path.join(base or DEFAULT_BASE, store)


def run_folders(root):
    """Download folders of one store (names start with the date), oldest first."""
    if not os.path.isdir(root):
        return []
    return sorted(os.path.join(root, d) for d in os.listdir(root)
                  if d[:2] == "20" and os.path.isdir(os.path.join(root, d)))


def month_folders(root, month):
    """The store's folders downloaded for this month (`…_m<YYYYMM>`), oldest first."""
    tag = "_m%s" % month.replace("-", "")
    return [f for f in run_folders(root) if os.path.basename(f).endswith(tag)]


def pick_folder(store, month, base=None):
    """The newest folder of the month that holds the month's invoice; else the
    newest folder of any kind that holds it. None when there is none."""
    y, m = parse_month(month)
    root = store_root(store, base)
    for f in reversed(month_folders(root, month)):
        if has_invoice(f, y, m):
            return f
    for f in reversed(run_folders(root)):
        if has_invoice(f, y, m):
            return f
    return None


def out_path(out_dir, store, month, suffix):
    """`<store>_<YYYY-MM>_<suffix>.xlsx` — never the monthly report's name."""
    safe = re.sub(r'[\\/:*?"<>|]', "_", store)
    return os.path.join(out_dir or DEFAULT_OUT_DIR, "%s_%s_%s.xlsx" % (safe, month, suffix))


def snapshot_label(folder):
    """'20261013_golden_m202609' -> '2026-10-13 下载'."""
    b = os.path.basename(os.path.normpath(folder))
    if re.match(r"^\d{8}", b):
        return "%s-%s-%s 下载" % (b[:4], b[4:6], b[6:8])
    return b
