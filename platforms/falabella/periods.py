# -*- coding: utf-8 -*-
"""Date rules for one month of Falabella data. Pure functions, no browser.

The report is always about one calendar month M (normally last month). What has
to be fetched for it:

  order      the order list, in pieces of at most 14 days (the order page's
             calendar allows no more)
  billing    "Transacciones" for transaction dates 1..end of M
  stock      the Fulfillment product snapshot (taken whenever the run happens)
  billing1   "Factura falabella.com" report of every billing period that touches
             M. Periods run 21st -> 20th, so M touches two of them; the second
             is listed only after it closes on the 20th of M+1
  billing2   "Liquidación factura" report of M
  invoice1   the PDFs (facturas, notas de crédito) of each billing1 period
  invoice2   the PDF(s) of the liquidation of M
"""
import calendar
import datetime

MONTHS_ES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
             "agosto", "septiembre", "octubre", "noviembre", "diciembre")

# Forms the monthly report does not read (the PDFs are for the accountant).
NOT_IN_REPORT_ROUTES = ("invoice1", "invoice2")


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


def order_chunks(month, max_days=14):
    """[(start, end)] covering the month, each at most max_days long, ends included."""
    a, end = month_bounds(month)
    out = []
    while a <= end:
        b = min(end, a + datetime.timedelta(days=max_days - 1))
        out.append((a, b))
        a = b + datetime.timedelta(days=1)
    return out


def billing_periods(month, cycle_day=21):
    """Billing periods (start, end) that overlap the month, oldest first.

    cycle_day=21: [21 Aug - 20 Sep], [21 Sep - 20 Oct] for September.
    """
    a, b = month_bounds(month)
    out = []
    # the period containing day 1, then step forward
    if a.day >= cycle_day:
        start = a.replace(day=cycle_day)
    else:
        prev = a.replace(day=1) - datetime.timedelta(days=1)
        start = prev.replace(day=cycle_day)
    while start <= b:
        nxt = (start.replace(day=1) + datetime.timedelta(days=32)).replace(day=cycle_day)
        end = nxt - datetime.timedelta(days=1)
        out.append((start, end))
        start = nxt
    return out


def es_date(d, pad=False):
    """date -> '21 agosto 2026' ('01 septiembre 2026' with pad=True)."""
    return "%s %s %d" % (("%02d" if pad else "%d") % d.day, MONTHS_ES[d.month - 1], d.year)


def billing_row_label(start, end):
    """How the Factura falabella.com table writes a period."""
    return "%s - %s" % (es_date(start), es_date(end))


def liquidation_row_label(month):
    """How the Liquidación factura table writes a month (days zero-padded)."""
    a, b = month_bounds(month)
    return "%s - %s" % (es_date(a, pad=True), es_date(b, pad=True))


def iso(d):
    return d.strftime("%Y-%m-%d")


def period_id(start, end):
    return "%s_%s" % (iso(start), iso(end))


def form_ids(month):
    """Every form id expected for the month, in fetch order."""
    out = ["order/%s" % period_id(a, b) for a, b in order_chunks(month)]
    out += ["billing", "stock"]
    for a, b in billing_periods(month):
        out.append("billing1/%s" % period_id(a, b))
    out.append("billing2")
    for a, b in billing_periods(month):
        out.append("invoice1/%s" % period_id(a, b))
    out.append("invoice2")
    return out


def route_of(form):
    return str(form).split("/")[0]


def in_report(form):
    return route_of(form) not in NOT_IN_REPORT_ROUTES


def billing_period_state(start, end, today):
    """'open' (still running), 'closed' (should be listed).

    A period is listed by Falabella once it has closed; a day of grace covers
    the overnight generation.
    """
    if today <= end + datetime.timedelta(days=1):
        return "open"
    return "closed"
