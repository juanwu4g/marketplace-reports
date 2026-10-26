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
    hosts = dict(routes)
    assert hosts["ledger"].startswith("https://proveedores") and hosts["orders"].startswith("https://marketplace")
    m = meta(r)
    assert m["platform"] == "LV" and m["flags"]["marketplace_store"] == MARKET_USER
    assert not [x for x in m["missing"] if download_diag.needs_attention(x)]


def test_without_a_marketplace_user_the_portal_still_runs(config, routes, tmp_path):
    config(companion=None)
    client, r = run(tmp_path)
    assert client.opened == [MAIN]
    assert {"ledger", "sales", "payments", "invoices"} <= set(r["files"])
    gone = dict((x["form"], x) for x in r["missing"])
    assert set(gone) == {"orders", "logistics", "balance"}
    assert all(x["in_report"] and download_diag.needs_attention(x) for x in gone.values())
    assert "marketplace_store" in " ".join(r["errors"])


def test_marketplace_user_must_be_allowed(config, routes, tmp_path):
    config(allowed=(MAIN,))
    client = Client()
    with pytest.raises(store_config.StoreNotAllowed):
        run(tmp_path, client)
    assert client.opened == []                  # refused before Ziniao opened anything


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
