# -*- coding: utf-8 -*-
"""Date rules for one month of Amazon data. Pure functions, no browser.

The report is always about one calendar month M (normally last month). What has
to be fetched for it (FORMS.md):

  order     "Transacciones" of M (monthly option of the reports repository, CSV)
  summary   "Resumen" of M (same page, PDF) — an independent control total
  invoice   the seller-fee CFDIs (XML + PDF) of every fee period that overlaps M,
            plus the payment complements dated in M. Amazon's periods are about
            14 days but their start days drift (Aug 9, Aug 12, Aug 17, Aug 31 …),
            so they cannot be predicted: one form, the download records which
            periods it found. The period that straddles month end is listed only
            after it closes — owed, fetched by a later run.
  stock     the FBA inventory snapshot (taken whenever the run happens)
  taxcert   the monthly "Certificado de retención" (CFDI de retenciones) of M,
            stamped around the 6th of M+1

Form ids are "<route>" — one per route.
"""
import calendar
import datetime

ROUTES = ("order", "summary", "invoice", "stock", "taxcert")
# The fee invoices feed a reconciliation (bridge B) but no figure of the report
# depends on them: missing ones are a note, the check is skipped.
NOT_IN_REPORT_ROUTES = ("invoice",)

MONTHS_ES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
             "agosto", "septiembre", "octubre", "noviembre", "diciembre")
# How the summary/transaction file names abbreviate months ("2026SepMonthly…").
MONTHS_EN3 = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep",
              "Oct", "Nov", "Dec")
# Amazon issues the withholding certificate of month M in the first week of M+1.
TAXCERT_DAY = 7


def parse_month(month):
    y, m = int(str(month)[:4]), int(str(month)[5:7])
    if not 1 <= m <= 12:
        raise ValueError("bad month %r" % month)
    return y, m


def month_bounds(month):
    y, m = parse_month(month)
    return (datetime.date(y, m, 1),
            datetime.date(y, m, calendar.monthrange(y, m)[1]))


def previous_month(today):
    first = today.replace(day=1) - datetime.timedelta(days=1)
    return "%04d-%02d" % (first.year, first.month)


def form_ids(month):
    """Every form id expected for the month, in fetch order."""
    return list(ROUTES)


def route_of(form):
    return str(form).split("/")[0]


def in_report(form):
    return route_of(form) not in NOT_IN_REPORT_ROUTES


def file_stem(month):
    """'2026Sep' — how Amazon names the monthly files of the reports repository."""
    y, m = parse_month(month)
    return "%d%s" % (y, MONTHS_EN3[m - 1])


def dropdown_month(month):
    """The month dropdown's value: 0-based ("8" = Septiembre)."""
    return str(parse_month(month)[1] - 1)


def taxcert_label(month):
    """How the Mexican tax invoices page writes a month:
    '1 de septiembre de 2026 30 de septiembre de 2026'."""
    a, b = month_bounds(month)
    es = MONTHS_ES[a.month - 1]
    return "%d de %s de %d %d de %s de %d" % (a.day, es, a.year, b.day, es, b.year)


def taxcert_due(month):
    """First day the certificate of the month can be expected."""
    a, b = month_bounds(month)
    return b + datetime.timedelta(days=TAXCERT_DAY)


# --------------------------------------------------------- fee invoice periods

_INV_DATE = "%a %b %d %H:%M:%S UTC %Y"          # 'Mon Sep 14 19:07:36 UTC 2026'


def parse_invoice_date(text):
    """'Mon Sep 14 19:07:36 UTC 2026' -> datetime (UTC, naive)."""
    return datetime.datetime.strptime(" ".join(str(text).split()), _INV_DATE)


def overlaps(month, start, end):
    """A fee period [start, end) (UTC datetimes) shares days with the month.
    Amazon's periods end at 19:07 UTC of the day the next one begins, so a period
    that ends on the 1st at 19:07 still overlaps the month of the 1st."""
    a, b = month_bounds(month)
    lo = datetime.datetime.combine(a, datetime.time())
    hi = datetime.datetime.combine(b + datetime.timedelta(days=1), datetime.time())
    return start < hi and end > lo


def in_month(month, when):
    a, b = month_bounds(month)
    return a <= when.date() <= b


def straddles_end(month, end):
    """The period runs past the month's last day (listed only after it closes)."""
    return end.date() > month_bounds(month)[1]
