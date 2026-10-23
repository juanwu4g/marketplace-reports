# -*- coding: utf-8 -*-
"""MercadoLibre's monthly results files (core/results, docs/RESULTS_AND_PERIODS.md §2):
one per store workbook built, for each basis.

* order basis (账单 vs 订单: 订单基准): written by report_build.py itself — the
  subprocess that holds S — when run_reports passes --results-version;
* bill basis (账单基准): written by run_reports.bill_reports, which holds S;
* a version whose workbook was not written gets a "failed" file from
  run_reports (failed_document), so 版本说明 can say why.

Grades come from core/grading; the figures from report/figures.py and
bill_report/figures.py (the workbook's own numbers).
"""
import datetime

from core import grading
from core import markets
from core import results

PLATFORM = "ML"


def market_of(site):
    return markets.market_id(PLATFORM, str(site).upper())


def _missing(meta):
    return [r for r in ((meta or {}).get("missing") or []) if r and r.get("form")]


def order_document(S, checks, version, workbook, formula_errors=False, built_at=None):
    """The order basis' results document for one store (S from report.derived)."""
    from platforms.mercadolibre.report import figures as OF
    from platforms.mercadolibre.report import site
    country = site.SITE.upper()
    provisional = list(S.get("bill_open") or [])
    incomplete = bool(S.get("ventas_warning"))
    missing = _missing(S.get("meta"))
    g = grading.grade(PLATFORM, "order", True, checks, missing, incomplete, provisional, formula_errors)
    return results.document(
        PLATFORM, market_of(country), country, site.CUR, site.MONEY_DEC, S["store"], _month(S["p0"]), "order",
        version, built_at or datetime.datetime.now(), [S["folder"]], workbook, g,
        results.plain(checks), missing, provisional, incomplete, OF.figures(S))


def bill_document(S, checks, version, workbook, formula_errors=False, built_at=None):
    """The bill basis' results document for one store (S from bill_report.model)."""
    from platforms.mercadolibre.bill_report import figures as BF
    p = S["prof"]
    provisional = list(S.get("provisional") or [])
    missing = _missing(S.get("meta"))
    g = grading.grade(PLATFORM, "bill", True, checks, missing, False, provisional, formula_errors)
    return results.document(
        PLATFORM, market_of(p.country), p.country, p.currency, p.decimals, S["store"], S["month"], "bill",
        version, built_at or datetime.datetime.now(), [S["folder"]], workbook, g,
        results.plain(checks), missing, provisional, False, BF.figures(S))


def failed_document(store, month, basis, region, version, folder, why, checks=(), meta=None):
    """A version whose workbook was not written: grade failed, no figures."""
    from platforms.mercadolibre.bill_report import profile as P
    p = P.get(str(region).lower())
    g = grading.grade(PLATFORM, basis, False, why_not_written=why)
    return results.document(
        PLATFORM, market_of(region), str(region).upper(), p.currency, p.decimals, store, month, basis, version,
        datetime.datetime.now(), [folder] if folder else [], None, g, results.plain(checks),
        _missing(meta), (), False, None)


def write(doc, log=print):
    """core/results.write; a failure is reported, never raised (the workbook is out)."""
    return write_from(lambda: doc, log)


def write_from(make, log=print):
    """Build a document with make() and write it. Any failure — building the
    figures included — is printed with its traceback and returns None: a results
    file must never stop the reports of the other stores."""
    try:
        doc = make()
        path = results.write(doc)
        log("    [结果] %s · %s · %s" % (doc["basis"], results.GRADE_CN[doc["grade"]["level"]], path))
        return path
    except Exception as e:                    # noqa: BLE001
        import traceback
        log("    [结果文件写入失败] %s: %s" % (e.__class__.__name__, e))
        log(traceback.format_exc())
        return None


def _month(p0):
    return "%04d-%02d" % (p0.year, p0.month)
