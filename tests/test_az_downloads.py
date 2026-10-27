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


# ------------------------------------------------------------------ robustness (2026-10 review)

import pytest  # noqa: E402

from core import browser_kit as bk  # noqa: E402


class Page(object):
    def __init__(self, url, title):
        self.current_url, self.title = url, title


REPO = "https://sellercentral.amazon.com.mx/payments/reports-repository"


def test_a_5xx_page_is_retryable_not_blocked():
    for title in ("503 Service Unavailable", "Error 500", "502 Bad Gateway",
                  "Internal Server Error"):
        with pytest.raises(forms.ServerDown) as e:
            forms.blocked(Page(REPO, title))
        assert isinstance(e.value, bk.PageError) and not isinstance(e.value, forms.Blocked)
    with pytest.raises(forms.ServerDown):
        forms.blocked(Page("chrome-error://chromewebdata/", "sellercentral.amazon.com.mx"))
    with pytest.raises(forms.ServerDown):
        forms.blocked(Page("about:blank", ""))


def test_real_sign_in_pages_are_still_blocked():
    assert forms.blocked(Page("https://sellercentral.amazon.com.mx/ap/signin?openid.x=1",
                              "Amazon Sign-In"))
    assert forms.blocked(Page("https://sellercentral.amazon.com.mx/ap/mfa?arb=1",
                              "Verificación en dos pasos"))
    assert forms.blocked(Page("https://sellercentral.amazon.com.mx/somewhere", "Iniciar sesión"))
    assert not forms.blocked(Page(REPO, "Biblioteca de documentos de pagos"))


class _Drv(object):
    def __init__(self, titles):
        self.titles, self.gets = list(titles), 0
        self.current_url = REPO

    def get(self, url):
        self.gets += 1

    @property
    def title(self):
        return self.titles[min(self.gets, len(self.titles)) - 1]


def _ctx(d):
    return forms.Ctx(d, ".", None, "2026-09", site.get("mx"), log=lambda *a: None)


def test_go_reloads_once_after_a_5xx_then_raises_server_down(monkeypatch):
    monkeypatch.setattr(forms.time, "sleep", lambda s: None)
    d = _Drv(["503 Service Temporarily Unavailable", "Seller Central"])
    _ctx(d).go(forms.REPOSITORY)
    assert d.gets == 2
    d = _Drv(["503 Service Unavailable"] * 3)
    with pytest.raises(forms.ServerDown):
        _ctx(d).go(forms.REPOSITORY)
    assert d.gets == 2
    # the page left open is diagnosed as the platform's hiccup, not as "needs login"
    ev = {"signal": "exception", "url": REPO, "title": "503 Service Unavailable",
          "exception_class": "ServerDown", "exception": "x"}
    assert download_diag.classify(ev)["cause"] == "PLATFORM_ERROR"


def test_retried_tries_again_on_server_down_but_never_on_blocked(monkeypatch):
    monkeypatch.setattr(bk.time, "sleep", lambda s: None)
    calls = []

    def flaky(ctx):
        calls.append(1)
        if len(calls) == 1:
            raise forms.ServerDown("503")
        return "ok"
    assert forms.retried(_ctx(None), "x", flaky) == "ok" and len(calls) == 2

    def signin(ctx):
        calls.append(2)
        raise forms.Blocked("signin")
    with pytest.raises(forms.Blocked):
        forms.retried(_ctx(None), "x", signin)
    assert calls.count(2) == 1


def _stock_row(when, fmt=".csv", ready=True):
    return [["Inventario de Logística de Amazon", when, when, fmt], object() if ready else None]


def test_owed_stock_takes_only_a_csv_newer_than_the_request():
    old, new = _stock_row("19/10/26, 2:07 a.m."), _stock_row("01/11/26, 0:05 a.m.")
    e = {"route": "stock", "seen_before": forms.stock_row_id(old)}
    assert forms.owed_stock_row(e, [new, old]) is new
    assert forms.owed_stock_row(e, [old]) is None
    assert forms.owed_stock_row(e, [_stock_row("01/11/26", ready=False), old]) is None
    assert forms.owed_stock_row({"route": "stock"}, [old]) is old         # old entry: as before


class _SelfClosingTab(object):
    """'Ver' opens a tab that triggers the download and closes itself at once:
    the handle shows up, then every call on it says the window is gone."""
    def __init__(self):
        self.calls, self.current = 0, "main"
        self.switch_to = self

    @property
    def window_handles(self):
        self.calls += 1
        return ["main", "tab"] if self.calls == 2 else ["main"]   # 1: before; 2: after the click

    @property
    def current_window_handle(self):
        return self.current

    def window(self, h):
        from selenium.common.exceptions import NoSuchWindowException
        if h != "main":
            raise NoSuchWindowException("no such window: target window already closed")
        self.current = h

    @property
    def current_url(self):
        from selenium.common.exceptions import NoSuchWindowException
        raise NoSuchWindowException("no such window: target window already closed")

    def close(self):
        pass


def test_a_document_tab_that_closes_itself_is_a_page_error_not_a_dead_browser(tmp_path, monkeypatch):
    from core import browser_kit as bk
    d = _SelfClosingTab()
    ctx = type("Ctx", (), {})(); ctx.d, ctx.out_dir = d, str(tmp_path)
    monkeypatch.setattr(forms.bk, "real_click", lambda d, b: None)
    monkeypatch.setattr(forms.bk, "wait_for", lambda fn, timeout=0, poll=0: (lambda r: r if r else None)(_try(fn)))
    monkeypatch.setattr(forms.bk, "wait_new_files", lambda *a, **k: [])
    with pytest.raises(bk.PageError) as e:
        forms.fetch_document(ctx, object())
    assert "自己关掉了" in str(e.value)
    assert d.current == "main"                                   # back on the main window
    # the diagnosis must not call this a dead browser
    assert not download_diag.DEAD_SESSION.search(str(e.value))


def _try(fn):
    try:
        return fn()
    except Exception:
        return None
