# -*- coding: utf-8 -*-
"""Liverpool's download runner without a browser: two Ziniao stores into one
folder, in order (supplier portal, then the marketplace user named in config),
and what is recorded when one of them cannot be used."""
import io
import json
import os

import pytest

from core import download_diag
from core import store_config
from platforms.liverpool import forms
from platforms.liverpool import run_downloads as RD

MAIN, MARKET_USER = "Liverpool Report", "Liverpool Order 子账号(Nikki)"


@pytest.fixture
def config(tmp_path, monkeypatch):
    def write(allowed=(MAIN, MARKET_USER), companion=MARKET_USER):
        entry = {"platform": "LV", "country": "MX"}
        if companion:
            entry["marketplace_store"] = companion
        p = tmp_path / "config.json"
        p.write_text(json.dumps({"stores": {"allowed": list(allowed), "catalog": {MAIN: entry}}}),
                     encoding="utf-8")
        monkeypatch.setattr(store_config, "CONFIG_PATH", str(p))
    write()
    return write


class Session(object):
    driver = object()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


class Client(object):
    def __init__(self):
        self.opened = []

    def find_store(self, name):
        return {"browserName": name, "platform_name": "自定义平台"}

    def open_store(self, store, close_when_done=True):
        self.opened.append(store["browserName"])
        return Session()


@pytest.fixture
def routes(monkeypatch):
    """Every route writes one file; the openers do nothing (a test may replace one)."""
    calls = []

    def route(name):
        def fn(ctx):
            calls.append((name, ctx.host))
            fn_ = os.path.join(ctx.out_dir, "%s.xlsx" % name)
            with open(fn_, "w") as fh:
                fh.write("x" * 100)
            ctx.got(name, [os.path.basename(fn_)])
        return fn
    monkeypatch.setattr(forms, "ROUTES", dict((r, route(r)) for r in RD.ROUTE_NAMES))
    monkeypatch.setattr(forms, "portal_login", lambda ctx: None)
    monkeypatch.setattr(forms, "check_language", lambda ctx: None)
    monkeypatch.setattr(forms, "marketplace_open", lambda ctx: None)
    monkeypatch.setattr(RD.bk, "set_download_dir", lambda d, p: None)
    monkeypatch.setattr(download_diag, "give_up", lambda diag, d, form, signal, **kw:
                        diag.items.setdefault(form, {"evidence": dict(kw, signal=signal), "period": None,
                                                     "required": True}))
    return calls


def run(tmp_path, client=None):
    client = client or Client()
    r = RD.run_store(client, MAIN, "2026-09", out_dir=str(tmp_path / "out"), log=lambda *a: None)
    return client, r


def meta(r):
    with io.open(os.path.join(r["out_dir"], "run_meta.json"), encoding="utf-8") as fh:
        return json.load(fh)


def test_two_stores_one_folder(config, routes, tmp_path):
    client, r = run(tmp_path)
    assert client.opened == [MAIN, MARKET_USER]
    assert set(r["files"]) == set(RD.ROUTE_NAMES)
    # the portal store first (its forms), then the marketplace user (its forms)
    assert [name for name, _h in routes] == list(RD.periods.PORTAL_ROUTES) + list(RD.periods.MARKETPLACE_ROUTES)
    m = meta(r)
    assert m["platform"] == "LV" and m["flags"]["marketplace_store"] == MARKET_USER
    assert not [x for x in m["missing"] if download_diag.needs_attention(x)]


def test_one_store_does_both_sites_marketplace_first(config, routes, tmp_path):
    config(companion=None)
    client, r = run(tmp_path)
    assert client.opened == [MAIN]                       # opened once, for both sites
    assert set(r["files"]) == set(RD.ROUTE_NAMES)
    assert [name for name, _h in routes][:3] == list(RD.periods.MARKETPLACE_ROUTES)
    assert not [x for x in r["missing"] if download_diag.needs_attention(x)]


def test_marketplace_user_must_be_allowed(config, routes, tmp_path):
    config(allowed=(MAIN,))
    client = Client()
    with pytest.raises(store_config.StoreNotAllowed):
        run(tmp_path, client)
    assert client.opened == []                  # refused before Ziniao opened anything


def test_portal_login_comes_from_the_catalog_and_is_never_recorded(config, routes, tmp_path, monkeypatch):
    import json as _json
    p = store_config.CONFIG_PATH
    cfg = _json.load(open(p, encoding="utf-8"))
    cfg["stores"]["catalog"][MAIN]["portal_login"] = {"user": "P00000001", "password": "s3cret-pw"}
    open(p, "w", encoding="utf-8").write(_json.dumps(cfg))
    seen = []
    monkeypatch.setattr(forms, "portal_login", lambda ctx: seen.append(ctx.login))
    client, r = run(tmp_path)
    assert seen == [{"user": "P00000001", "password": "s3cret-pw"}]
    meta = open(os.path.join(r["out_dir"], "run_meta.json"), encoding="utf-8").read()
    report = open(os.path.join(r["out_dir"], "REPORT.md"), encoding="utf-8").read()
    assert "s3cret-pw" not in meta + report and '"portal_login_configured": true' in meta


class FakeField(object):
    def __init__(self, value=""):
        self.value = value

    def get_attribute(self, name):
        return self.value

    def clear(self):
        self.value = ""

    def send_keys(self, text):
        self.value += text

    def is_displayed(self):
        return True


class FakePortal(object):
    """The SAP sign-in page: two fields and a button; 'Salir del sistema' shows after a
    submit with the right password."""
    title = "SAP NetWeaver Portal"
    current_url = "https://proveedores.liverpool.com.mx/irj/portal"
    window_handles = ["w1"]

    def __init__(self, password):
        self.user, self.pw, self.ok_pw, self.signed = FakeField(), FakeField(), password, False

    def get(self, url):
        pass

    def find_elements(self, by, value):
        if value == "logonuidfield":
            return [self.user]
        if "Salir del sistema" in str(value):
            return [FakeField()] if self.signed else []
        return []

    def find_element(self, by, value):
        return {"logonpassfield": self.pw}.get(value, FakeField())

    def submit(self):
        self.signed = self.user.value == "P00000001" and self.pw.value == self.ok_pw


@pytest.fixture
def fast(monkeypatch):
    monkeypatch.setattr(forms.time, "sleep", lambda s: None)
    monkeypatch.setattr(forms, "real_window", lambda d: "w1")
    monkeypatch.setattr(forms.bk, "real_click", lambda d, el: d.submit())
    monkeypatch.setattr(forms.bk, "wait_for", lambda fn, timeout=0, poll=0: fn())


def portal_ctx(d, login):
    from platforms.liverpool import site
    return forms.Ctx(d, ".", None, "2026-09", site.get("MX"), login=login)


def test_portal_sign_in_types_the_configured_account(fast):
    d = FakePortal("s3cret-pw")
    ctx = portal_ctx(d, {"user": "P00000001", "password": "s3cret-pw"})
    assert forms.portal_login(ctx) is True
    assert d.signed and ctx.flags["portal_login"] == "config"


def test_portal_sign_in_wrong_password_or_none_needs_a_person(fast):
    with pytest.raises(forms.NeedsLogin):
        forms.portal_login(portal_ctx(FakePortal("other"), {"user": "P00000001", "password": "s3cret-pw"}))
    with pytest.raises(forms.NeedsLogin):
        forms.portal_login(portal_ctx(FakePortal("s3cret-pw"), None))


def test_portal_sign_in_needed_is_page_blocked_and_marketplace_still_runs(config, routes, tmp_path,
                                                                        monkeypatch):
    def needs_login(ctx):
        raise forms.NeedsLogin("供应商门户登录页：紫鸟没有自动填入账号密码")
    monkeypatch.setattr(forms, "portal_login", needs_login)
    client, r = run(tmp_path)
    assert client.opened == [MAIN, MARKET_USER]
    gone = dict((x["form"], x) for x in r["missing"])
    assert set(gone) == {"ledger", "sales", "payments", "invoices"}
    assert gone["ledger"]["cause"] == "PAGE_BLOCKED"
    assert {"orders", "logistics", "balance"} <= set(r["files"])


def test_files_in_a_subfolder_of_the_run_folder_count(tmp_path):
    sub = tmp_path / "Notas_Cargo_202609"
    sub.mkdir()
    (sub / "a.xml").write_text("<x/>")
    assert download_diag._rejected(["Notas_Cargo_202609/a.xml"], str(tmp_path)) == []
    assert download_diag._rejected(["Notas_Cargo_202609/missing.xml"], str(tmp_path)) == ["missing.xml（不在目录里）"]
