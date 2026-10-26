# -*- coding: utf-8 -*-
"""Date rules for one month of Liverpool data. Pure functions, no browser.

The report is about one calendar month M. What a store yields for it (FORMS.md):

  supplier portal (the store's own user, e.g. "Liverpool Report")
    ledger    "Cobros y Pagos Marketplace": every ledger line posted in M
              (Fecha de contabilización), both accounts — the report's baseline
    sales     "Ventas Marketplace": every sale / return document of M with its
              order number, Liverpool article and commission — the link from a
              ledger document to an order and a SKU
    payments  "Detalle de Pagos": the payments made in M and the open items
              (evidence for the cash bridge; the ledger already holds the money)
    invoices  "Notas de Cargo": Liverpool's fee CFDIs (XML + PDF) of M and M+1
              (the tax documents of the fees; the last week of M is invoiced in M+1)
  marketplace seller centre (the store's marketplace user, config marketplace_store)
    orders    the order list (the last four months, as the page offers it): order
              date, seller SKU, status — order-date KPIs and the SKU of each order
    logistics "Costos logísticos" charged in M, one row per order: the shipping
              charge of each order (the ledger has them per day)
    balance   "Balance de cuenta" registered in M: extra logistics charges and the
              digital fee, per order where they have one

Form ids are "<route>" — one per route.
"""
import calendar
import datetime

PORTAL_ROUTES = ("ledger", "sales", "payments", "invoices")
MARKETPLACE_ROUTES = ("orders", "logistics", "balance")
ROUTES = PORTAL_ROUTES + MARKETPLACE_ROUTES
# The payments and the fee CFDIs are evidence: no figure of the report depends on them.
NOT_IN_REPORT_ROUTES = ("payments", "invoices")


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


def next_month(month):
    y, m = parse_month(month)
    return "%04d-%02d" % (y + (m == 12), 1 if m == 12 else m + 1)


def form_ids(month):
    """Every form id expected for the month, in fetch order."""
    return list(ROUTES)


def route_of(form):
    return str(form).split("/")[0]


def in_report(form):
    return route_of(form) not in NOT_IN_REPORT_ROUTES


def sap_date(d):
    """The supplier portal's date fields: '01.09.2026'."""
    return d.strftime("%d.%m.%Y")


def web_date(d):
    """The marketplace's date fields (typed digit by digit): '01092026'."""
    return d.strftime("%d%m%Y")


def stamp(month):
    """'202609' — in the names the downloads are saved under."""
    y, m = parse_month(month)
    return "%04d%02d" % (y, m)
