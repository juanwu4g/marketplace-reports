# -*- coding: utf-8 -*-
"""The five Amazon routes for one store and one month (FORMS.md).

Every route takes a Ctx and fills ctx.files[form] with the downloaded file names.
When a form cannot be had it records why — ctx.give_up(form, signal, ...) —
instead of raising: download_diag turns the evidence into a cause code. Only a
store-level problem (wrong language, sign-in/captcha page, dead browser) is
raised to run_downloads.

Form ids are those of periods.form_ids(): order, summary, invoice, stock, taxcert.

Seller Central is built from Katal web components (kat-dropdown, kat-button,
kat-table-row …) with open shadow roots: every lookup walks them (core/browser_kit).
"""
import base64
import datetime
import io
import json
import os
import re
import time
import traceback

from core import browser_kit as bk
from core import download_diag
from platforms.amazon import periods
from selenium.common.exceptions import WebDriverException

REPOSITORY = "/payments/reports-repository"
FEE_INVOICES = "/tax/seller-fee-invoices"
TAX_INVOICES = "/sellertaxinvoice/displaySellerTaxInvoices"
INVENTORY = "/reportcentral/MANAGE_INVENTORY_HEALTH/1"

# Reports repository: report type values and the word the request list shows.
REPORT_TYPES = {
    "order": ("SELLER_TRANSACTION_DATE_RANGE", "Transacciones", ".csv"),
    "summary": ("SELLER_SUMMARY_DATE_RANGE", "Resumen", ".pdf"),
}
MONTH_ABBR = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep",
              "oct", "nov", "dic")

# What the fee invoice page says when its service fails (seen 2026-10-19 with
# 250 invoices on the account): never "no invoices", always "try again later".
SERVICE_DOWN = re.compile(r"No invoices available|Unable to connect to service|"
                          r"No hay facturas disponibles|No se puede conectar", re.I)
INVOICE_ROW = re.compile(
    r"^(?P<type>.+?) (?P<fmt>PDF|XML|CSV) (?P<num>[0-9A-Fa-f]{8}-[0-9A-Fa-f-]{27}) .*?"
    r"(?P<start>[A-Z][a-z]{2} [A-Z][a-z]{2} \d{2} \d\d:\d\d:\d\d UTC \d{4}) "
    r"(?P<end>[A-Z][a-z]{2} [A-Z][a-z]{2} \d{2} \d\d:\d\d:\d\d UTC \d{4})")
FEE_INVOICE = re.compile(r"^Factura de la tarifa|Settlement Note", re.I)
PAYMENT_COMPLEMENT = re.compile(r"^Complemento de pago|PaymentComplement", re.I)


class LanguageWrong(RuntimeError):
    """The seller centre is not in Spanish. The message contains 界面语言 so
    that download_diag classifies it as LANGUAGE_WRONG."""


class Blocked(RuntimeError):
    """A sign-in, verification or captcha page instead of Seller Central. A person
    has to deal with it (never solved by the script); every form has already been
    given up with the page as evidence (PAGE_BLOCKED)."""


class ServerDown(bk.PageError):
    """Amazon answered with a 5xx / browser error page instead of Seller Central.
    The platform is having a moment: worth another try (a PageError, so retried()
    tries again), never a reason to send a person to sign in. The page stays
    open, so download_diag reads its title and says PLATFORM_ERROR."""


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

    def go(self, path, settle=0, tries=2, pause=20):
        """Open a page. A 5xx page is loaded once more after a pause (it is
        usually gone by then); a sign-in page raises Blocked at once."""
        for i in range(tries):
            self.d.get(self.host + path)
            if settle:
                time.sleep(settle)
            try:
                if blocked(self.d):
                    raise Blocked("页面被带到了登录/验证页：%s" % self.d.current_url[:120])
                return
            except ServerDown as e:
                if i == tries - 1:
                    raise
                self.log("  %s（%s），%d 秒后重新打开" % (e, path, pause))
                time.sleep(pause)

    def got(self, form, names):
        names = [n for n in names if n]
        if names:
            self.files.setdefault(form, []).extend(names)
            if self.diag is not None:
                download_diag.delivered(self.diag, form, self.files[form])
        return names

    def give_up(self, form, signal, **kw):
        self.log("  [放弃] %s：%s %s" % (form, signal, kw.get("step") or ""))
        if self.diag is None:            # collecting owed reports: nothing to record
            return None
        return download_diag.give_up(self.diag, self.d, form, signal, **kw)

    def nothing_there(self, form, reason):
        """The platform has nothing for this form in this period (terminal)."""
        self.log("  [无数据] %s：%s" % (form, reason))
        self.no_data[form] = reason
        self.give_up(form, "no-data", row_state="empty", status_text=reason, shot=False)

    def owe(self, form, route, requested=None, state="absent", note=""):
        """Requested (or published later by Amazon), not there yet: fetched by
        run_batch --collect-only or the next run."""
        self.log("  %s 还没好（%s），记为欠账%s" % (form, state, "：" + note if note else ""))
        self.pending.append({"form": form, "route": route, "range": self.month,
                             "requested_at": requested or time.time(),
                             "last_state": state, "last_status_text": note,
                             "probes": 1, "note": note})

    def rename(self, name, new):
        src, dst = os.path.join(self.out_dir, name), os.path.join(self.out_dir, new)
        if os.path.abspath(src) == os.path.abspath(dst):
            return new
        if os.path.exists(dst):
            os.remove(dst)
        os.replace(src, dst)
        return new


# 真正要人去处理的：登录、两步验证、验证码。download_diag.BLOCKED_TITLE 还收了
# 404/50x/ERR_ 这些错误页的标题，在这里不能用 —— 2026-10 审查发现一个临时的
# 「503 Service Unavailable」会让整家店变成"需要人工登录"，而且不会重试。
SIGN_IN_TITLE = re.compile(
    r"\bsign-?in\b|\blog ?in\b|iniciar sesi[oó]n|inicia sesi[oó]n|ingresa a tu cuenta|"
    r"verifica|captcha|autenticaci[oó]n|two-step|dos pasos", re.I)
# 页面没打开：浏览器自己的错误页、一片空白。重新打开就好。
NOT_LOADED = re.compile(r"^about:blank|^chrome-error://|^data:,", re.I)


def blocked(d):
    """True on a sign-in / verification / captcha page. Raises ServerDown on a
    5xx or browser error page (retryable, not a sign-in problem)."""
    try:
        url, title = d.current_url or "", d.title or ""
    except Exception:
        return False
    if download_diag.SERVER_ERROR.search(title):
        raise ServerDown("亚马逊返回了错误页「%s」" % title[:60])
    if NOT_LOADED.search(url) or re.search(r"\bERR_", title):
        raise ServerDown("页面没有打开（%s %s）" % (url[:60], title[:60]))
    return bool(download_diag.BLOCKED_URL.search(url) or SIGN_IN_TITLE.search(title))


def guarded(ctx, route, forms, fn, *a):
    """Run one route; an exception gives up every form of it not yet delivered."""
    try:
        fn(ctx, *a)
    except (LanguageWrong, Blocked):
        raise
    except Exception as e:
        ctx.errors.append("%s: %s: %s" % (route, e.__class__.__name__, str(e)[:200]))
        ctx.log("  [路线失败] %s：%s" % (route, e))
        if not isinstance(e, bk.PageError):
            traceback.print_exc()
        ev = download_diag.exception_evidence(e, step=route)
        for f in forms:
            if f not in ctx.files and f not in ctx.no_data and \
                    not any(p["form"] == f for p in ctx.pending):
                ctx.give_up(f, "exception", **{k: v for k, v in ev.items() if k != "signal"})


def retried(ctx, what, fn, *a, attempts=2):
    """bk.retried that never retries a wrong language or a sign-in page."""
    return bk.retried(ctx, what, fn, *a, attempts=attempts,
                      never_retry=(LanguageWrong, Blocked))


def katal_click(d, el):
    """Click a kat-* component: the real <button>/<label> sits in its shadow root."""
    inner = d.execute_script(
        "var r=arguments[0].shadowRoot;"
        "return r?(r.querySelector('button')||r.querySelector('label')"
        "||r.querySelector('input')):null;", el)
    bk.js_click(d, inner or el)


# ================================================================ language gate

_LOCALE = ("var h=document.querySelector('#ngstrim-header');"
           "return h?h.getAttribute('data-locale'):null;")


def page_locale(d):
    try:
        return d.execute_script(_LOCALE)
    except Exception:
        return None


def check_language(ctx):
    """Seller Central's language is an account setting ("#ngstrim-header
    data-locale"). Column names, the summary PDF and every label the routes read
    are Spanish: when someone switched it, switch it back — a page load with
    mons_sel_locale — and record it; only if that fails stop with LanguageWrong."""
    ctx.go(REPOSITORY)
    loc = bk.wait_for(lambda: page_locale(ctx.d), timeout=60, poll=2)
    if blocked(ctx.d):
        raise Blocked("页面被带到了登录/验证页：%s" % ctx.d.current_url[:120])
    if not loc:
        raise bk.PageError("Seller Central header never rendered (#ngstrim-header not found)")
    want = ctx.site["locale"]
    ctx.flags["locale"] = loc
    if not loc.lower().startswith("es"):
        ctx.log("  界面语言是 %s，不是西班牙语：改回 %s" % (loc, want))
        ctx.go("%s?mons_sel_locale=%s" % (REPOSITORY, want), settle=5)
        new = bk.wait_for(lambda: page_locale(ctx.d), timeout=60, poll=2) or ""
        if not new.lower().startswith("es"):
            raise LanguageWrong("界面语言是 %s，不是西班牙语（%s），脚本自动改回也没成功："
                                "请在这家店右上角的语言菜单里选 Español" % (loc, want))
        ctx.flags["language_switched"] = {"from": loc, "to": new}
        ctx.flags["locale"] = new
        ctx.log("  已改回 %s" % new)
    return loc


# ================================================================ 1+2. reports repository

_DROPDOWNS = bk.WALK + """var o=[];W(document,function(e){if(e.tagName!=='KAT-DROPDOWN')return;
 var opts=[];try{(e.options||[]).forEach(function(x){opts.push([String(x.value),String(x.name)]);});}catch(z){}
 o.push([e,String(e.value),opts,e.hasAttribute('disabled')]);});return o;"""
_SETV = ("var e=arguments[0];e.value=arguments[1];"
         "e.dispatchEvent(new CustomEvent('change',{bubbles:true,composed:true,"
         "detail:{value:arguments[1]}}));return String(e.value);")
_REPO_ROWS = bk.WALK + """var o=[];W(document,function(e){if(e.tagName!=='KAT-TABLE-ROW')return;
 var t=(e.innerText||'').replace(/\\s+/g,' ').trim();if(!/GMT/.test(t))return;
 var s=null,b=[];W(e,function(x){if(!s&&x.tagName==='KAT-STATUSINDICATOR')s=x.getAttribute('label');
 if(x.tagName==='KAT-BUTTON')b.push([x,x.getAttribute('label')||'']);});
 e.querySelectorAll('kat-statusindicator').forEach(function(x){if(!s)s=x.getAttribute('label');});
 e.querySelectorAll('kat-button').forEach(function(x){b.push([x,x.getAttribute('label')||'']);});
 o.push([e,t,s||'',b]);});return o;"""


def dropdowns(d):
    try:
        return d.execute_script(_DROPDOWNS) or []
    except Exception:
        return []


def _dropdown(d, test):
    for el, value, opts, disabled in dropdowns(d):
        if not disabled and test(value, opts):
            return el, value, opts
    return None


def range_pattern(month):
    """The requested range as a request-list row writes it:
    '1 sept 2026, 0:00:00 GMT-6 30 sept 2026, 23:59:59 GMT-6'."""
    a, b = periods.month_bounds(month)
    mon = MONTH_ABBR[a.month - 1]
    return re.compile(r"\b1 %s\w*\.? %d, 0:00:00 GMT[-+]\d+ %d %s\w*\.? %d, 23:59:59"
                      % (mon, a.year, b.day, mon, b.year), re.I)


def all_repo_rows(d):
    try:
        return d.execute_script(_REPO_ROWS) or []
    except Exception:
        return []


def is_row_of(row, month, word):
    return bool(range_pattern(month).search(row[1]) and re.search(r"\b%s\b" % word, row[1]))


def repo_rows(d, month, word):
    """[(element, text, status, [(button, label)])] of the request list for this
    month and report kind, newest first (the page lists newest first)."""
    return [r for r in all_repo_rows(d) if is_row_of(r, month, word)]


def ready(row):
    return any(lbl.startswith("Descargar") for _, lbl in row[3])


def _repo_page(ctx):
    ctx.go(REPOSITORY)
    got = bk.wait_for(lambda: _dropdown(ctx.d, lambda v, o: v.startswith("SELLER_")),
                      timeout=60, poll=2)
    if not got:
        raise bk.PageError("report type dropdown (SELLER_…) not found")
    return got


def request_report(ctx, form):
    """Choose type, Mensual, month and year; press the request button.
    Returns the request list (text, status) as it was before the request."""
    d = ctx.d
    type_value, word, _ = REPORT_TYPES[form]
    el, value, opts = _repo_page(ctx)
    if value != type_value:
        d.execute_script(_SETV, el, type_value)
        time.sleep(2)
    radio = bk.wait_for(lambda: bk.find(d, "kat-radiobutton[value=MONTHLY]"), timeout=20)
    if not radio:
        raise bk.PageError("radio button MONTHLY (Mes) not found")
    katal_click(d, radio[0])
    y, m = periods.parse_month(ctx.month)
    month_dd = bk.wait_for(lambda: _dropdown(
        d, lambda v, o: any(n.lower().startswith("enero") for _, n in o)), timeout=20)
    year_dd = bk.wait_for(lambda: _dropdown(
        d, lambda v, o: o and all(re.match(r"^\d{4}$", x) for x, _ in o)), timeout=20)
    if not month_dd or not year_dd:
        raise bk.PageError("month / year dropdowns not found after choosing Mes")
    if str(y) not in [x for x, _ in year_dd[2]]:
        raise bk.PageError("year %d is not offered (%s)" % (y, [x for x, _ in year_dd[2]]))
    d.execute_script(_SETV, year_dd[0], str(y))
    shown = d.execute_script(_SETV, month_dd[0], periods.dropdown_month(ctx.month))
    if shown != periods.dropdown_month(ctx.month):
        raise bk.PageError("month dropdown kept %r" % shown)
    time.sleep(1)
    before = [(r[1], r[2]) for r in all_repo_rows(d)]
    btn = bk.need(d, "kat-button#filter-generate-button", what="request button")
    katal_click(d, btn)
    ctx.log("  已申请 %s %s" % (word, ctx.month))
    return before


def wait_ready(ctx, form, before, timeout):
    """Wait for our new request to be ready. Returns (row, state).

    The list shows a fixed number of rows, newest first, and requests of the same
    kind on the same day read the same: a new request is recognised as "the top
    row is of this kind and month, and the list is no longer what it was" (it
    shifted by one row, or the top row went back to En curso)."""
    d = ctx.d
    _, word, _ = REPORT_TYPES[form]
    end = time.time() + timeout
    n = 0
    seen = False
    while True:
        rows = all_repo_rows(d)
        top = rows[0] if rows else None
        mine = top is not None and is_row_of(top, ctx.month, word)
        if mine and [(r[1], r[2]) for r in rows] != before:
            seen = True
        if seen and mine:
            if ready(top):
                return top, "ready"
            if re.search(r"Vencido|Fall|Error", top[2], re.I):
                return top, "failed"
        if time.time() >= end:
            return (top if seen and mine else None), ("generating" if seen else "absent")
        n += 1
        time.sleep(10)
        if mine and n % 3 == 0:                 # the list does not refresh by itself
            for b, lbl in top[3]:
                if lbl.startswith("Actualizar"):
                    katal_click(d, b)


def download_repo_row(ctx, form, row):
    _, _, ext = REPORT_TYPES[form]
    btn = [b for b, lbl in row[3] if lbl.startswith("Descargar")]
    if not btn:
        return []
    before = bk.snapshot(ctx.out_dir)
    katal_click(ctx.d, btn[0])
    files = bk.wait_new_files(ctx.out_dir, before, timeout=90, start_timeout=40)
    return [f for f in files if f.lower().endswith(ext)]


def repo_report(ctx, form):
    """order (Transacciones CSV) or summary (Resumen PDF) of the month."""
    before = retried(ctx, "申请 %s" % form, request_report, form)
    requested = time.time()
    row, state = wait_ready(ctx, form, before, ctx.collect_timeout)
    if state != "ready":
        ctx.owe(form, form, requested, state, row[1][:120] if row else "")
        return
    files = download_repo_row(ctx, form, row)
    if not files:
        ctx.give_up(form, "download-not-landed", step=row[1][:80])
        return
    stem = periods.file_stem(ctx.month)
    if not all(f.startswith(stem) for f in files):
        raise bk.PageError("downloaded %s, expected a file of %s" % (files, stem))
    ctx.log("  %s -> %s（%d 秒）" % (form, files, time.time() - requested))
    ctx.got(form, files)


def orders(ctx):
    repo_report(ctx, "order")


def summary(ctx):
    repo_report(ctx, "summary")


# ================================================================ 3. fee invoices

_INV_ROWS = bk.WALK + """var o=[];W(document,function(e){if(e.tagName!=='TR')return;
 var b=null;e.querySelectorAll('button').forEach(function(x){if((x.innerText||'').trim()==='Ver')b=x;});
 if(b)o.push([(e.innerText||'').replace(/\\s+/g,' ').trim(),b]);});return o;"""
_FETCH = """var u=arguments[0],cb=arguments[arguments.length-1];
fetch(u,{credentials:'include'}).then(function(r){return r.arrayBuffer().then(function(b){
 var s='',a=new Uint8Array(b);for(var i=0;i<a.length;i+=8192)
 s+=String.fromCharCode.apply(null,a.subarray(i,i+8192));
 cb({status:r.status,type:r.headers.get('content-type')||'',b64:btoa(s)});});})
 .catch(function(e){cb({err:String(e)});});"""


def invoice_rows(d):
    """[(row dict, Ver button)] of the fee invoice table."""
    try:
        raw = d.execute_script(_INV_ROWS) or []
    except Exception:
        return []
    out = []
    for text, btn in raw:
        m = INVOICE_ROW.match(text)
        if not m:
            continue
        out.append(({"type": m.group("type"), "format": m.group("fmt"),
                     "number": m.group("num").upper(),
                     "start": periods.parse_invoice_date(m.group("start")),
                     "end": periods.parse_invoice_date(m.group("end"))}, btn))
    return out


def _page_says_down(d):
    try:
        text = d.execute_script("return document.body.innerText") or ""
    except Exception:
        return False
    return bool(SERVICE_DOWN.search(text))


def _invoice_page(ctx, tries=4, pause=60):
    """The invoice table, waiting out Amazon's service hiccups. [] if it stays down."""
    for i in range(tries):
        ctx.go(FEE_INVOICES)
        rows = bk.wait_for(lambda: invoice_rows(ctx.d) or (_page_says_down(ctx.d) and ["down"]),
                           timeout=60, poll=3)
        if rows and rows != ["down"]:
            return rows
        if i < tries - 1:
            ctx.log("  费用发票页说没有发票/连不上服务（第 %d 次），%d 秒后再试" % (i + 1, pause))
            time.sleep(pause)
    return []


def _load_more(ctx, rows):
    """Older rows are behind "Load More": load until the month is covered."""
    a, _ = periods.month_bounds(ctx.month)
    oldest = datetime.datetime.combine(a, datetime.time())
    for _ in range(6):
        if not rows or min(r["start"] for r, _ in rows) <= oldest:
            return rows
        more = [b for b in bk.find(ctx.d, "button, kat-button, a")
                if re.search(r"Load More|Cargar m[aá]s|Ver m[aá]s",
                             (b.text or b.get_attribute("label") or ""), re.I)]
        if not more:
            return rows
        katal_click(ctx.d, more[0])
        time.sleep(5)
        rows = invoice_rows(ctx.d)
    return rows


def wanted_invoices(month, rows):
    """The rows of the month: fee invoices whose period overlaps it, payment
    complements dated in it. CSV breakdowns are not needed (the XML has the
    amounts)."""
    out = []
    for r, btn in rows:
        if r["format"] not in ("XML", "PDF"):
            continue
        if FEE_INVOICE.search(r["type"]) and periods.overlaps(month, r["start"], r["end"]):
            out.append((r, btn))
        elif PAYMENT_COMPLEMENT.search(r["type"]) and periods.in_month(month, r["start"]):
            out.append((r, btn))
    return out


def fetch_document(ctx, btn):
    """'Ver' either opens a tab on /documents/download/… (fee invoices) or
    downloads the file straight away (payment complements). Returns
    (where, bytes): for a tab, its URL is read, the tab closed and the bytes
    fetched with the page's session."""
    d = ctx.d
    main = d.current_window_handle
    before = set(d.window_handles)
    files_before = bk.snapshot(ctx.out_dir)
    bk.real_click(d, btn)
    new = bk.wait_for(lambda: [h for h in d.window_handles if h not in before]
                      or (bk.snapshot(ctx.out_dir) - files_before and ["file"]),
                      timeout=25, poll=0.5)
    if not new:
        raise bk.PageError("'Ver' opened no tab and downloaded nothing")
    if new == ["file"]:
        got = [f for f in bk.wait_new_files(ctx.out_dir, files_before, timeout=60)
               if not f.endswith(bk.PARTIAL)]
        if not got:
            raise bk.PageError("'Ver' download never completed")
        path = os.path.join(ctx.out_dir, got[0])
        with open(path, "rb") as fh:
            data = fh.read()
        os.remove(path)
        return got[0], data
    url = None
    try:
        d.switch_to.window(new[0])
        url = bk.wait_for(lambda: "/documents/download/" in d.current_url and d.current_url,
                          timeout=20, poll=0.5)
        if not url:
            url = "tab:" + d.current_url
    except WebDriverException as e:
        # 那个标签页自己关掉了（2026-10-26 实测：'Ver' 开一个标签页触发下载，然后
        # 马上关闭）。它关掉不等于浏览器没了：主窗口还在，下面照常回到主窗口。
        # 这里要是把 selenium 的 "no such window" 原样抛出去，会被诊断成
        # "店铺打不开"，整家店判成要人处理。
        # 不把 selenium 的原话（"no such window…"）带进错误信息：诊断会把那几个词当成
        # 浏览器断开（DEAD_SESSION）。
        url = "tab:closed:" + e.__class__.__name__
    finally:
        for h in new:
            try:
                d.switch_to.window(h)
                d.close()
            except Exception:
                pass
        d.switch_to.window(main)
    if url.startswith("tab:"):
        # a tab that shows the file itself (an XML or PDF viewer) has no
        # /documents/download/ address: the click may also have saved it
        got = bk.wait_new_files(ctx.out_dir, files_before, timeout=15, start_timeout=10)
        if got:
            path = os.path.join(ctx.out_dir, got[0])
            with open(path, "rb") as fh:
                data = fh.read()
            os.remove(path)
            return got[0], data
        url = url[4:]
        if url.startswith("closed:"):
            raise bk.PageError("'Ver' 开的标签页在读到地址之前就自己关掉了，也没有文件落盘（%s）"
                               % url[7:])
        if not url.startswith("http"):
            raise bk.PageError("document tab opened %r" % url[:120])
    d.set_script_timeout(90)
    r = d.execute_async_script(_FETCH, url)
    if r.get("err") or r.get("status") != 200:
        raise bk.PageError("document fetch failed: %s" % (r.get("err") or r.get("status")))
    return url, base64.b64decode(r["b64"])


def invoices(ctx):
    form = "invoice"
    rows = _invoice_page(ctx)
    if not rows:
        ctx.owe(form, form, state="delayed",
                note="费用发票页一直说没有发票/连不上服务（亚马逊那边的问题，稍后补取）")
        return
    rows = _load_more(ctx, rows)
    want = wanted_invoices(ctx.month, rows)
    index_path = os.path.join(ctx.out_dir, "fee_invoices.json")
    index = []
    if os.path.exists(index_path):
        with io.open(index_path, encoding="utf-8") as fh:
            index = json.load(fh)
    have = {(x["number"], x["format"]) for x in index}
    names = []
    for r, btn in want:
        ext = r["format"].lower()
        name = "fee_%s.%s" % (r["number"], ext)
        if (r["number"], r["format"]) not in have or \
                not os.path.exists(os.path.join(ctx.out_dir, name)):
            url, data = fetch_document(ctx, btn)
            head = data[:5]
            if (ext == "xml" and not head.startswith(b"<?xml") and b"<cfdi" not in data[:400]) \
                    or (ext == "pdf" and not head.startswith(b"%PDF")):
                raise bk.PageError("%s is not a %s file (%r)" % (name, ext, head))
            with open(os.path.join(ctx.out_dir, name), "wb") as fh:
                fh.write(data)
            index = [x for x in index if (x["number"], x["format"]) != (r["number"], r["format"])]
            index.append({"type": r["type"], "format": r["format"], "number": r["number"],
                          "start": r["start"].isoformat(), "end": r["end"].isoformat(),
                          "file": name})
        names.append(name)
    with io.open(index_path, "w", encoding="utf-8") as fh:
        json.dump(index, fh, ensure_ascii=False, indent=1)
    periods_ = sorted({(r["start"].isoformat(), r["end"].isoformat())
                       for r, _ in want if FEE_INVOICE.search(r["type"])})
    ctx.flags["invoice_periods"] = ["%s/%s" % p for p in periods_]
    ctx.log("  费用发票：%d 个文件，%d 个账期" % (len(names), len(periods_)))
    if names:
        ctx.got(form, names + ["fee_invoices.json"])
    # the period running past month end is listed only after it closes
    last_end = max([r["end"] for r, _ in want if FEE_INVOICE.search(r["type"])] or [None],
                   key=lambda x: x or datetime.datetime.min)
    _, b = periods.month_bounds(ctx.month)
    if last_end is None or last_end.date() <= b:
        ctx.owe(form, form, state="delayed",
                note="跨月末的账期（%s 之后）要等它结束后亚马逊才出发票" % (last_end or b))
    elif not names:
        ctx.nothing_there(form, "这个月没有费用发票")


# ================================================================ 4. stock

_STOCK_ROWS = bk.WALK + """var o=[];W(document,function(e){if(e.tagName!=='KAT-TABLE-ROW')return;
 var c=[];e.querySelectorAll('kat-table-cell').forEach(function(x){c.push((x.innerText||'').trim());});
 var b=null;W(e,function(x){if(!b&&x.tagName==='BUTTON'&&(x.innerText||'').trim()==='Descargar')b=x;});
 if(c.length>=4&&c[0]!=='Tipo de reporte')o.push([c,b]);});return o;"""


def stock_rows(d):
    try:
        return d.execute_script(_STOCK_ROWS) or []
    except Exception:
        return []


def stock(ctx):
    form = "stock"
    d = ctx.d
    ctx.go(INVENTORY)
    btn = bk.need(d, "button[value=CSV]", timeout=60, what="'Solicitar descarga de .csv'")
    bk.wait_for(lambda: stock_rows(d), timeout=20)
    rows = stock_rows(d)
    before = len(rows)
    seen = stock_row_id(rows[0]) if rows else None
    bk.real_click(d, btn)
    requested = time.time()
    ready = bk.wait_for(lambda: [r for r in stock_rows(d)[:1]
                                 if len(stock_rows(d)) > before and r[1] is not None
                                 and r[0][3].lower() == ".csv"],
                        timeout=ctx.collect_timeout, poll=5)
    if not ready:
        ctx.owe(form, form, requested, "generating")
        ctx.pending[-1]["seen_before"] = seen
        return
    files = _download_stock(ctx, ready[0][1])
    if not files:
        ctx.give_up(form, "download-not-landed", step="Inventario de Logística de Amazon")
        return
    ctx.log("  库存 -> %s" % files)
    ctx.got(form, files)


def stock_row_id(row):
    """What tells one request apart: Tipo de reporte, Rango de fechas cubierto
    (the request time), Fecha solicitada, Formato — the fixed cells, not the button."""
    return " | ".join(row[0][:4])


def owed_stock_row(entry, rows):
    """The ready CSV of an owed stock request, or None: the newest one listed
    above the row that was newest before the request — a CSV older than the
    request is an old snapshot. Old entries without "seen_before": the newest."""
    if "seen_before" in entry and entry["seen_before"]:
        for i, r in enumerate(rows):
            if stock_row_id(r) == entry["seen_before"]:
                rows = rows[:i]
                break
    rows = [r for r in rows if r[1] is not None and r[0][3].lower() == ".csv"]
    return rows[0] if rows else None


def _download_stock(ctx, button):
    before = bk.snapshot(ctx.out_dir)
    bk.real_click(ctx.d, button)
    files = [f for f in bk.wait_new_files(ctx.out_dir, before, timeout=90, start_timeout=40)
             if f.lower().endswith(".csv")]
    if not files:
        return []
    return [ctx.rename(files[0], "stock_%s.csv" % ctx.today.strftime("%Y%m%d"))]


# ================================================================ 5. withholding certificate

_CERT_ROWS = bk.WALK + """var o=[];W(document,function(e){if(e.tagName!=='TR'||e.id!=='invoiceDownloadRequestRow')return;
 var l={};e.querySelectorAll('a').forEach(function(a){var t=(a.innerText||'').trim().toLowerCase();if(t==='pdf'||t==='xml')l[t]=a;});
 o.push([(e.innerText||'').replace(/\\s+/g,' ').trim(),l]);});return o;"""
UUID = re.compile(r"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}")


def cert_rows(d):
    try:
        return d.execute_script(_CERT_ROWS) or []
    except Exception:
        return []


def taxcert(ctx):
    form = "taxcert"
    d = ctx.d
    ctx.go(TAX_INVOICES)
    rows = bk.wait_for(lambda: cert_rows(d), timeout=60, poll=3)
    if rows is None:
        if _page_says_down(d):
            ctx.owe(form, form, state="delayed", note="墨西哥税务发票页连不上服务（稍后补取）")
            return
        raise bk.PageError("Mexican tax invoice table (invoiceDownloadRequestRow) not found")
    label = periods.taxcert_label(ctx.month)
    hit = [(t, links) for t, links in rows
           if t.startswith("Certificado de retenci") and label in t]
    if not hit:
        due = periods.taxcert_due(ctx.month)
        ctx.owe(form, form, state="delayed",
                note="%s 的预扣税证明还没出（一般在 %s 前后出）" % (ctx.month, due))
        return
    text, links = hit[0]
    uid = UUID.search(text)
    uid = uid.group(0).upper() if uid else ctx.month
    names = []
    for ext in ("xml", "pdf"):
        if ext not in links:
            raise bk.PageError("certificate link %s not found" % ext)
        before = bk.snapshot(ctx.out_dir)
        bk.real_click(d, links[ext])
        got = [f for f in bk.wait_new_files(ctx.out_dir, before, timeout=60, start_timeout=30)
               if f.lower().endswith("." + ext)]
        if not got:
            ctx.give_up(form, "download-not-landed", step="Certificado de retención %s" % ext)
            return
        names.append(ctx.rename(got[0], "retenciones_%s.%s" % (uid, ext)))
    ctx.log("  预扣税证明 -> %s" % names)
    ctx.got(form, names)


# route name (periods.ROUTES) -> function(ctx)
ROUTES = {"order": orders, "summary": summary, "invoice": invoices, "stock": stock,
          "taxcert": taxcert}


# ================================================================ owed reports

def collect_owed(ctx, entry):
    """Fetch one owed report (requested earlier, not ready then). Returns the
    form's file names, or [] if still not there."""
    route = entry["route"]
    if route in REPORT_TYPES:
        _repo_page(ctx)
        _, word, _ = REPORT_TYPES[route]
        rows = [r for r in repo_rows(ctx.d, ctx.month, word) if ready(r)]
        return download_repo_row(ctx, route, rows[0]) if rows else []
    if route == "stock":
        ctx.go(INVENTORY)
        rows = bk.wait_for(lambda: stock_rows(ctx.d), timeout=60, poll=3) or []
        row = owed_stock_row(entry, rows)
        if not row:
            return []
        if entry.get("requested_at"):
            # the snapshot is of the day it was requested, not of the day it is fetched
            ctx.today = datetime.date.fromtimestamp(entry["requested_at"])
        return _download_stock(ctx, row[1])
    if route in ("invoice", "taxcert"):
        # these pages publish documents later; the route itself fetches what is
        # missing. Still owed afterwards = not complete yet.
        ctx.pending = []
        ROUTES[route](ctx)
        if ctx.pending:
            return []
        return ctx.files.get(route, [])
    return []
