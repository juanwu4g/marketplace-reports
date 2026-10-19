# -*- coding: utf-8 -*-
"""The Amazon routes for one store and one month (FORMS.md).

Every route takes a Ctx and fills ctx.files[form] with the downloaded file names.
When a form cannot be had it records why — ctx.give_up(form, signal, ...) —
instead of raising: download_diag turns the evidence into a cause code. Only a
store-level problem (wrong language, dead browser) is raised to run_downloads.

Browser helpers: core/browser_kit (shadow roots, find / need / wait_for,
js_click / real_click, wait_new_files, retried). Write the routes from what
tools/probe_session showed, never from guesses.
"""
import datetime
import os
import time
import traceback

from core import browser_kit as bk
from core import download_diag
from platforms.amazon import periods


class LanguageWrong(RuntimeError):
    """The seller centre is not in the language the routes are written for.
    The message contains 界面语言 so that download_diag classifies it as
    LANGUAGE_WRONG."""


class NotBuiltYet(NotImplementedError):
    pass


def not_built(what):
    raise NotBuiltYet("亚马逊 的%s还没实现：先做探查（tools/probe_session），"
                      "写进 FORMS.md，再写这条路线" % what)


class Ctx(object):
    """Everything one store's routes share. One per store (never module state)."""

    def __init__(self, driver, out_dir, diag, month, site, today=None,
                 collect_timeout=240, log=print):
        self.d = driver
        self.out_dir = out_dir
        self.diag = diag
        self.month = month
        self.site = site
        self.host = site["host"]
        self.today = today or datetime.date.today()
        self.collect_timeout = collect_timeout
        self.log = log
        self.files = {}          # form -> [file names]
        self.no_data = {}        # form -> reason (platform says: nothing there)
        self.pending = []        # requested, not ready: owed entries
        self.errors = []         # route-level failures, one line each
        self.flags = {}

    def go(self, path, settle=0):
        self.d.get(self.host + path)
        if settle:
            time.sleep(settle)

    def got(self, form, names):
        names = [n for n in names if n]
        if names:
            self.files.setdefault(form, []).extend(names)
            download_diag.delivered(self.diag, form, self.files[form])
        return names

    def give_up(self, form, signal, **kw):
        self.log("  [放弃] %s：%s %s" % (form, signal, kw.get("step") or ""))
        return download_diag.give_up(self.diag, self.d, form, signal, **kw)

    def nothing_there(self, form, reason):
        """The platform has nothing for this form in this period (terminal)."""
        self.log("  [无数据] %s：%s" % (form, reason))
        self.no_data[form] = reason
        self.give_up(form, "no-data", row_state="empty", status_text=reason, shot=False)

    def rename(self, name, new):
        src, dst = os.path.join(self.out_dir, name), os.path.join(self.out_dir, new)
        if os.path.exists(dst):
            os.remove(dst)
        os.replace(src, dst)
        return new


def guarded(ctx, route, forms, fn, *a):
    """Run one route; an exception gives up every form of it not yet delivered."""
    try:
        fn(ctx, *a)
    except LanguageWrong:
        raise
    except Exception as e:
        ctx.errors.append("%s: %s: %s" % (route, e.__class__.__name__, str(e)[:200]))
        ctx.log("  [路线失败] %s：%s" % (route, e))
        if not isinstance(e, bk.PageError):
            traceback.print_exc()
        ev = download_diag.exception_evidence(e, step=route)
        for f in forms:
            if f not in ctx.files and f not in ctx.no_data:
                ctx.give_up(f, "exception", **{k: v for k, v in ev.items() if k != "signal"})


def retried(ctx, what, fn, *a, attempts=2):
    """bk.retried that never retries a wrong language."""
    return bk.retried(ctx, what, fn, *a, attempts=attempts, never_retry=(LanguageWrong,))


# ================================================================ language gate

def check_language(ctx):
    """Read the UI language; not the expected one -> raise LanguageWrong."""
    not_built("界面语言检查")


# ================================================================ routes

def orders(ctx):
    not_built("订单路线")


def settlement(ctx):
    not_built("结算路线")


# route name (periods.ROUTES) -> function(ctx)
ROUTES = {"orders": orders, "settlement": settlement}


def collect_owed(ctx, entry):
    """Fetch one owed report (requested earlier, not ready then). Returns file names."""
    not_built("补取欠账")
