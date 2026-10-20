# -*- coding: utf-8 -*-
"""A store that fails before any form (its first page never renders) must be
diagnosed from where the browser actually is. 2026-10-20: Falabella's session
had expired, the browser sat on /user/auth/login, and the run reported
INTERNAL_ERROR ("the script is broken") instead of PAGE_BLOCKED ("someone has
to log in") — nobody would have been sent to log in."""
from core import download_diag


class FakeDriver(object):
    def __init__(self, url, title):
        self.current_url, self.title = url, title


class DeadDriver(object):
    @property
    def current_url(self):
        raise RuntimeError("chrome not reachable")


def test_login_page_is_page_blocked_not_a_script_error():
    exc = RuntimeError("payments page never rendered (no data-localefull found)")
    d = FakeDriver("https://sellercenter.falabella.com/user/auth/login#/ordersandtransactionsoverview",
                   "Seller Center")
    ev = download_diag.exception_evidence(exc, step="store", driver=d)
    assert ev["url"].endswith("#/ordersandtransactionsoverview")
    rec = download_diag.classify(ev, {"form": "billing"})
    assert rec["cause"] == "PAGE_BLOCKED"
    # without the page, the same failure looks like a bug in the script
    bare = download_diag.classify(download_diag.exception_evidence(exc, step="store"),
                                  {"form": "billing"})
    assert bare["cause"] != "PAGE_BLOCKED"


def test_amazon_sign_in_page_is_page_blocked():
    ev = download_diag.exception_evidence(
        RuntimeError("Seller Central header never rendered"), step="store",
        driver=FakeDriver("https://sellercentral.amazon.com.mx/ap/signin?x=1", "Amazon Sign-In"))
    assert download_diag.classify(ev, {"form": "order"})["cause"] == "PAGE_BLOCKED"


def test_a_dead_browser_leaves_only_the_exception():
    ev = download_diag.exception_evidence(RuntimeError("boom"), step="store", driver=DeadDriver())
    assert "url" not in ev and ev["exception"] == "boom"
