# -*- coding: utf-8 -*-
"""Falabella automatic sign-in (platforms/falabella/login.py). The browser parts
are proven on the VM; here: the address fix and the places where it must stop."""
from platforms.falabella import login

AUTH = ("https://access-key-corp.falabella.tech/auth/realms/esti/protocol/openid-connect/auth"
        "?scope=openid%20profile%20email&login_hint=a%40b.cl&state=x&client_id=GSC")


def test_fix_adds_fpkt_once():
    fixed = login.with_fix(AUTH)
    assert fixed == AUTH + "&fpkt=true"
    assert login.with_fix(fixed) == fixed


def test_fix_keeps_fragment_and_handles_no_query():
    assert login.with_fix("https://h/p#/x") == "https://h/p?fpkt=true#/x"


def test_login_pages_recognised():
    assert login.on_login_page("https://sellercenter.falabella.com/user/auth/login#/ordersandtransactionsoverview")
    assert login.on_login_page(AUTH)
    assert not login.on_login_page("https://sellercenter.falabella.com/report/transaction/statement#/x")
    assert not login.on_login_page(None)


class Ctx(object):
    def __init__(self, d):
        self.d, self.flags, self.lines = d, {}, []

    def log(self, s):
        self.lines.append(s)


class NoDebugger(object):
    capabilities = {}


class EmptyEmail(object):
    capabilities = {"goog:chromeOptions": {"debuggerAddress": "127.0.0.1:1"}}
    current_url = "https://sellercenter.falabella.com/user/auth/login"
    clicked = False

    def set_page_load_timeout(self, s):
        pass

    def execute_script(self, src, selector, *a):
        return [Field("")] if "email" in selector else [Button(self)]


class Field(object):
    def __init__(self, v):
        self.v = v

    def get_attribute(self, k):
        return self.v


class Button(object):
    def __init__(self, d):
        d.clicked = True


def test_stops_without_debugger():
    c = Ctx(NoDebugger())
    assert login.sign_in(c) is False
    assert c.flags["auto_login"].startswith("failed")


def test_stops_when_email_not_prefilled_and_never_clicks():
    d = EmptyEmail()
    c = Ctx(d)
    assert login.sign_in(c, timeout=1) is False
    assert "邮箱" in c.flags["auto_login"] and not d.clicked


class StillSignedIn(object):
    """Keycloak's session is still valid: "Continuar" goes straight back to
    Seller Center (through one /user/auth/ hop), no sign-in page ever appears."""
    capabilities = {"goog:chromeOptions": {"debuggerAddress": "127.0.0.1:1"}}

    def __init__(self):
        self.urls = ["https://sellercenter.falabella.com/user/auth/login"]

    @property
    def current_url(self):
        return self.urls.pop(0) if len(self.urls) > 1 else self.urls[0]

    def set_page_load_timeout(self, s):
        pass


def test_still_signed_in_after_continuar_counts_as_logged_in(monkeypatch):
    from platforms.falabella import page_kit as pk
    d = StillSignedIn()
    monkeypatch.setattr(login, "_sso_target", lambda addr: None)
    monkeypatch.setattr(login.time, "sleep", lambda s: None)
    monkeypatch.setattr(pk, "find", lambda drv, sel, *a, **k: [Field("a@b.cl")])

    def click(drv, el):
        drv.urls = ["https://sellercenter.falabella.com/user/auth/callback?code=x",
                    "https://sellercenter.falabella.com/report/transaction/statement#/x"]
    monkeypatch.setattr(pk, "real_click", click)
    c = Ctx(d)
    assert login.sign_in(c, timeout=30) is True
    assert c.flags["auto_login"] == "ok"


def test_a_redirect_hop_alone_is_not_back_in():
    class D(object):
        current_url = "https://sellercenter.falabella.com/user/auth/callback?code=x"
    assert not login._back_in(D())
