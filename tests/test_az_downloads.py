# -*- coding: utf-8 -*-
"""Amazon download layer: the pure parts (dates, row parsing, which invoices
belong to a month) and the blocked-page rules. No browser, no store."""
import datetime

from core import download_diag
from platforms.amazon import forms
from platforms.amazon import periods
from platforms.amazon import site

DT = datetime.datetime


def test_month_helpers():
    assert periods.form_ids("2026-09") == ["order", "summary", "invoice", "stock", "taxcert"]
    assert periods.in_report("order") and not periods.in_report("invoice")
    assert periods.file_stem("2026-09") == "2026Sep"
    assert periods.dropdown_month("2026-09") == "8"
    assert periods.dropdown_month("2026-01") == "0"
    assert periods.taxcert_label("2026-09") == \
        "1 de septiembre de 2026 30 de septiembre de 2026"
    assert periods.taxcert_label("2026-02") == "1 de febrero de 2026 28 de febrero de 2026"
    assert periods.taxcert_due("2026-09") == datetime.date(2026, 10, 7)
    assert periods.previous_month(datetime.date(2026, 1, 9)) == "2025-12"


def test_invoice_period_overlap():
    p = periods.parse_invoice_date
    assert p("Mon Sep 14 19:07:36 UTC 2026") == DT(2026, 9, 14, 19, 7, 36)
    m = "2026-09"
    assert periods.overlaps(m, DT(2026, 8, 31, 19, 7), DT(2026, 9, 14, 19, 7))
    assert periods.overlaps(m, DT(2026, 9, 28, 19, 7), DT(2026, 10, 12, 19, 7))
    assert not periods.overlaps(m, DT(2026, 8, 17, 19, 7), DT(2026, 8, 31, 19, 7))
    assert not periods.overlaps(m, DT(2026, 10, 12, 19, 7), DT(2026, 10, 26, 19, 7))
    assert periods.straddles_end(m, DT(2026, 10, 12, 19, 7))


ROWS = [
    "Complemento de pago de la tarifa de vendedor para México PDF 52547EF6-C70C-11F1-90AF-0D0ABC5EA6E5 "
    "TIENDA EJEMPLO SERVICIOS COMERCIALES AMAZON MEXICO Amazon.com.mx Mon Oct 12 19:24:37 UTC 2026 "
    "Mon Oct 12 19:24:37 UTC 2026 Ver",
    "Factura de la tarifa de vendedor para México XML 681374E4-C707-11F1-B3C3-1BD6BE5B3A4C "
    "TIENDA EJEMPLO SERVICIOS COMERCIALES AMAZON MEXICO Amazon.com.mx Mon Sep 28 19:07:36 UTC 2026 "
    "Mon Oct 12 19:07:36 UTC 2026 Ver",
    "Complemento de pago de la tarifa de vendedor para México XML 008BFEC7-BC0C-11F1-8890-0386568DDFCB "
    "TIENDA EJEMPLO SERVICIOS COMERCIALES AMAZON MEXICO Amazon.com.mx Mon Sep 28 19:24:37 UTC 2026 "
    "Mon Sep 28 19:24:37 UTC 2026 Ver",
    "Factura de la tarifa de vendedor para México PDF A2D5BDEE-B107-11F1-A239-879AB14969DD "
    "TIENDA EJEMPLO SERVICIOS COMERCIALES AMAZON MEXICO Amazon.com.mx Mon Aug 31 19:07:36 UTC 2026 "
    "Mon Sep 14 19:07:36 UTC 2026 Ver",
    "Factura de la tarifa de vendedor para México XML 6E8DD725-A606-11F1-8D9F-D73B7AE8BDC2 "
    "TIENDA EJEMPLO SERVICIOS COMERCIALES AMAZON MEXICO Amazon.com.mx Mon Aug 17 19:07:36 UTC 2026 "
    "Mon Aug 31 19:07:36 UTC 2026 Ver",
    "Spft Settlement Note XML 11111111-2222-3333-4444-555555555555 "
    "TIENDA EJEMPLO SERVICIOS COMERCIALES AMAZON MEXICO Amazon.com.mx Mon Sep 14 19:07:36 UTC 2026 "
    "Mon Sep 28 19:07:36 UTC 2026 Ver",
    "Invoice Type Formato de la factura Invoice Number Payer Name",
]


def parsed():
    out = []
    for t in ROWS:
        m = forms.INVOICE_ROW.match(t)
        if m:
            out.append(({"type": m.group("type"), "format": m.group("fmt"),
                         "number": m.group("num"),
                         "start": periods.parse_invoice_date(m.group("start")),
                         "end": periods.parse_invoice_date(m.group("end"))}, None))
    return out


def test_invoice_rows_parse():
    rows = parsed()
    assert len(rows) == 6                               # the header row is not one
    assert rows[1][0]["type"] == "Factura de la tarifa de vendedor para México"
    assert rows[1][0]["format"] == "XML"
    assert rows[1][0]["number"] == "681374E4-C707-11F1-B3C3-1BD6BE5B3A4C"


def test_wanted_invoices_for_september():
    got = forms.wanted_invoices("2026-09", parsed())
    nums = [(r["number"][:8], r["format"]) for r, _ in got]
    assert ("681374E4", "XML") in nums                 # Sep 28 -> Oct 12 overlaps
    assert ("A2D5BDEE", "PDF") in nums                 # Aug 31 -> Sep 14 overlaps
    assert ("008BFEC7", "XML") in nums                 # payment complement in September
    assert ("11111111", "XML") in nums                 # settlement note of a September period
    assert ("52547EF6", "PDF") not in nums             # complement paid in October
    assert ("6E8DD725", "XML") not in nums             # August period


def test_request_list_row_match():
    pat = forms.range_pattern("2026-09")
    assert pat.search("19 oct 2026, GMT-6 1 sept 2026, 0:00:00 GMT-6 30 sept 2026, 23:59:59 "
                      "GMT-6 Pedidos estándar Transacciones Informe de transacciones")
    assert not pat.search("19 oct 2026, GMT-6 1 ago 2026, 0:00:00 GMT-6 31 ago 2026, 23:59:59 GMT-6")
    assert forms.range_pattern("2026-02").search(
        "3 mar 2026, GMT-6 1 feb 2026, 0:00:00 GMT-6 28 feb 2026, 23:59:59 GMT-6 Resumen")


def test_sign_in_and_captcha_pages_are_blocked():
    for url in ("https://sellercentral.amazon.com.mx/ap/signin?openid.return_to=x",
                "https://sellercentral.amazon.com.mx/ap/mfa?arb=1",
                "https://www.amazon.com.mx/errors/validateCaptcha"):
        assert download_diag.BLOCKED_URL.search(url)
    assert download_diag.BLOCKED_TITLE.search("Amazon Sign-In")
    assert not download_diag.BLOCKED_URL.search(
        "https://sellercentral.amazon.com.mx/payments/reports-repository")
    ev = {"signal": "exception", "url": "https://sellercentral.amazon.com.mx/ap/signin",
          "title": "Amazon Sign-In"}
    assert download_diag.classify(ev)["cause"] == "PAGE_BLOCKED"


def test_service_down_text():
    assert forms.SERVICE_DOWN.search("Tax Document Library No invoices available There are no")
    assert not forms.SERVICE_DOWN.search("Factura de la tarifa de vendedor para México")


def test_site_facts():
    s = site.get("mx")
    assert s["host"].startswith("https://sellercentral.amazon.com.mx")
    assert s["currency"] == "MXN" and s["vat"] == 0.16 and s["locale"].startswith("es")
