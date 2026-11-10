# -*- coding: utf-8 -*-
"""The Liverpool routes for one store and one month (FORMS.md).

Two sites, two Ziniao stores, one download folder:

  supplier portal  proveedores.liverpool.com.mx (SAP NetWeaver), the store itself:
                   ledger, sales, payments, invoices — the pages are Web Dynpro apps
                   on itsproves.liverpool.com.mx inside the portal's frames
  marketplace      marketplace.liverpool.com.mx, the store's marketplace user
                   (config.json stores.catalog[store].marketplace_store):
                   orders, logistics, balance

Every route takes a Ctx and fills ctx.files[form] with the downloaded file names.
When a form cannot be had it records why — ctx.give_up(form, signal, ...) — instead
of raising: download_diag turns the evidence into a cause code. Store-level
problems (a sign-in page, a wrong language) are raised to run_downloads.

The supplier portal's sign-in: Ziniao never fills saved passwords in a browser
started by automation (its password script returns early in WebDriver mode), so the
account configured for the store in config.json (stores.catalog[store].portal_login,
owner's decision 2026-10-27) is typed into the portal's own sign-in form — nowhere
else, and it is never logged. Without it the store is reported as needing a person
(NeedsLogin -> PAGE_BLOCKED).

The marketplace's sign-in (Entrada Única, Auth0): its session lives in the Ziniao
store, and when it has expired the store lands on the sign-in page with the e-mail
and password already filled by Ziniao (the store's own saved sign-in — a Chrome
autofill preview that a real click commits). The program types nothing there; it
only presses "Iniciar sesión" (marketplace_sign_in, proven in the test VM
2026-11-05). Nothing filled, or the sign-in refused -> Blocked (PAGE_BLOCKED).

The marketplace's language: before any route, marketplace_language reads the
header control ("es" / "en") and switches it back to Spanish through its menu
when someone left it in English (the routes find every button by its Spanish
text); the run records the language at entry (flags.marketplace_locale) and a
"language:" line when it had to change it. Unreadable or unswitchable ->
LanguageWrong (LANGUAGE_WRONG).
"""
import base64
import datetime
import os
import re
import time
import traceback

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from core import browser_kit as bk
from core import download_diag
from platforms.liverpool import periods
from platforms.liverpool import readers as R


class LanguageWrong(RuntimeError):
    """The site is not in Spanish. The message contains 界面语言 so that
    download_diag classifies it as LANGUAGE_WRONG."""


class NeedsLogin(RuntimeError):
    """A sign-in page the script must not fill in (download_diag: PAGE_BLOCKED)."""


class Blocked(RuntimeError):
    """A sign-in / verification page recognised by its address or title."""


class NotBuiltYet(NotImplementedError):
    pass


class Ctx(object):
    """Everything one store's routes share on one site. One per site (never module state)."""

    def __init__(self, driver, out_dir, diag, month, site, today=None,
                 collect_timeout=240, log=print, host=None, login=None):
        self.d = driver
        # the supplier portal's {"user", "password"} from config.json, or None; never logged
        self.login = login
        self.out_dir = out_dir
        self.diag = diag
        self.month = month
        self.site = site
        self.host = host or site["host"]
        self.today = today or datetime.date.today()
        self.collect_timeout = collect_timeout
        self.log = log
        self.files = {}          # form -> [file names]
        self.no_data = {}        # form -> reason (platform says: nothing there)
        self.pending = []        # nothing is asynchronous on Liverpool (kept for the contract)
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
    except (LanguageWrong, NeedsLogin, Blocked):
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


# ================================================================ shared page helpers

def real_window(d):
    """Ziniao opens with extra handles (an omnibox pop-up) that are not windows:
    using one, the site thinks it is a phone. Switch to the first real window."""
    for h in d.window_handles:
        d.switch_to.window(h)
        try:
            d.get_window_size()
            return h
        except Exception:
            continue
    raise bk.PageError("浏览器里没有可用的窗口")


def visible(d, xpath):
    return [e for e in d.find_elements(By.XPATH, xpath) if e.is_displayed()]


def by_text(d, text, exact=True):
    q = "normalize-space(text())='%s'" % text if exact else "contains(normalize-space(.),'%s')" % text
    return visible(d, "//*[%s]" % q)


def need_text(d, text, timeout=20, exact=True, what=None):
    els = bk.wait_for(lambda: by_text(d, text, exact), timeout=timeout, poll=1)
    if not els:
        raise bk.PageError("找不到「%s」%s" % (text, "（%s）" % what if what else ""))
    return els


def type_slowly(el, text):
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.BACKSPACE)
    for ch in text:
        el.send_keys(ch)
        time.sleep(0.06)


def download(ctx, click, new_name, timeout=120):
    """click() starts a download into ctx.out_dir; the new file is renamed."""
    before = bk.snapshot(ctx.out_dir)
    click()
    files = bk.wait_new_files(ctx.out_dir, before, timeout=timeout, start_timeout=60)
    if not files:
        return None
    return ctx.rename(files[-1], new_name)


# ================================================================ supplier portal

PORTAL_MENU = ("Facturación y Pagos", "Pagos MKP")


def portal_login(ctx):
    """Open the supplier portal and make sure the session is signed in."""
    d = ctx.d
    real_window(d)
    d.get(ctx.site["warmup"])                 # Akamai cookies (else "Access Denied")
    time.sleep(8)
    for attempt in range(2):
        d.get(ctx.site["portal"] + "/")
        time.sleep(8)
        if "access denied" not in (d.title or "").lower():
            break
        d.get(ctx.site["warmup"])
        time.sleep(10)
    if "access denied" in (d.title or "").lower():
        raise Blocked("门户拒绝访问（Access Denied）：%s" % d.current_url[:100])
    if by_text(d, "Salir del sistema"):
        return True
    user = d.find_elements(By.ID, "logonuidfield")
    if not user:
        raise bk.PageError("门户首页既不是登录页也不是已登录的页面")
    if not (user[0].get_attribute("value") or "").strip():
        # Ziniao never fills saved passwords in a browser started by automation:
        # the account configured for this store (config.json portal_login) is typed in.
        if not ctx.login:
            raise NeedsLogin("供应商门户登录页：没有可用的账号密码（config.json 这家店的 portal_login 没有写）")
        user[0].clear()
        user[0].send_keys(ctx.login["user"])
        pw = d.find_element(By.ID, "logonpassfield")
        pw.clear()
        pw.send_keys(ctx.login["password"])
        ctx.flags["portal_login"] = "config"
    else:
        ctx.flags["portal_login"] = "ziniao"
    bk.real_click(d, d.find_element(By.NAME, "uidPasswordLogon"))
    ok = bk.wait_for(lambda: by_text(d, "Salir del sistema"), timeout=40, poll=2)
    if not ok:
        raise NeedsLogin("供应商门户登录没有成功（config.json 里的密码可能已改或过期）")
    return True


def check_language(ctx):
    """The portal greets in Spanish ("Bienvenido …"); the routes are written for it."""
    if not by_text(ctx.d, "Bienvenido", exact=False):
        raise LanguageWrong("门户界面语言不是西班牙语（找不到 Bienvenido）")
    ctx.flags["locale"] = "es"


def work_area(ctx):
    """Into the frame that holds the page (portal desktop -> inner page -> work area)."""
    d = ctx.d
    d.switch_to.default_content()
    frames = d.find_elements(By.CSS_SELECTOR, "iframe")
    inner = [f for f in frames if "Área interna" in (f.get_attribute("name") or "") or
             "desktop" in (f.get_attribute("name") or "").lower()]
    d.switch_to.frame(inner[0] if inner else frames[3])
    d.switch_to.frame(d.find_element(By.NAME, "isolatedWorkArea"))


def nav(ctx, item):
    """Facturación y Pagos -> Pagos MKP -> the left-hand item; ends inside the work area."""
    d = ctx.d
    d.switch_to.default_content()
    for label in PORTAL_MENU:
        bk.real_click(d, need_text(d, label, what="门户菜单")[0])
        time.sleep(6)
    d.switch_to.default_content()
    frames = d.find_elements(By.CSS_SELECTOR, "iframe")
    inner = [f for f in frames if "Área interna" in (f.get_attribute("name") or "")] or frames[3:4]
    d.switch_to.frame(inner[0])
    links = visible(d, "//a[contains(normalize-space(.),'%s')]" % item)
    if not links:
        raise bk.PageError("门户左侧菜单里找不到「%s」" % item)
    bk.real_click(d, links[0])
    time.sleep(12)
    work_area(ctx)
    body = d.execute_script("return document.body.innerText || ''")
    if "bloqueado" in body.lower() or "chrome-error" in d.current_url:
        raise Blocked("「%s」打不开（紫鸟拦截了 itsproves.liverpool.com.mx？请确认已加白，并重启紫鸟客户端）" % item)


def inputs_after(d, label, n=2):
    """The first n visible text inputs after the label text (Web Dynpro ids change)."""
    xp = "//*[contains(normalize-space(text()),'%s')]/following::input[@type='text']" % label
    els = visible(d, xp)
    if len(els) < n:
        raise bk.PageError("「%s」后面的输入框不够 %d 个" % (label, n))
    return els[:n]


def set_field(ctx, el, value):
    bk.real_click(ctx.d, el)
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(value)
    el.send_keys(Keys.TAB)
    time.sleep(2)


def export_excel(ctx, opener, new_name):
    """A Web Dynpro table's export menu: opener -> "Exportar a Microsoft Excel"."""
    d = ctx.d

    def click():
        bk.real_click(d, opener)
        time.sleep(2)
        opt = by_text(d, "Exportar a Microsoft Excel")
        if not opt:
            raise bk.PageError("导出菜单里没有「Exportar a Microsoft Excel」")
        bk.real_click(d, opt[-1])
    return download(ctx, click, new_name)


def no_rows(d):
    return "La tabla no contiene datos" in (d.execute_script("return document.body.innerText || ''") or "")


def ledger(ctx):
    """Cobros y Pagos Marketplace: Todas las Partidas, Fecha de Contabilización = the month."""
    form = "ledger"
    d = ctx.d
    a, b = periods.month_bounds(ctx.month)
    nav(ctx, "Cobros y Pagos Marketplace")
    bk.real_click(d, need_text(d, "Todas las Partidas", what="科目选择")[0])
    time.sleep(3)
    work_area(ctx)
    lo, hi = inputs_after(d, "Fecha de Contabiliz")
    set_field(ctx, lo, periods.sap_date(a))
    work_area(ctx)
    lo, hi = inputs_after(d, "Fecha de Contabiliz")
    set_field(ctx, hi, periods.sap_date(b))
    work_area(ctx)
    got = [e.get_attribute("value") for e in inputs_after(d, "Fecha de Contabiliz")]
    if got != [periods.sap_date(a), periods.sap_date(b)]:
        raise bk.PageError("入账日期没有填进去：%s" % got)
    ctx.flags["ledger_range"] = got
    run = visible(d, "//div[@ct='B' and contains(@lsdata,'Concluido')]")
    el = run[0] if run else d.execute_script("return document.elementFromPoint(16, 17)")
    bk.real_click(d, el)
    if not bk.wait_for(lambda: by_text(d, "Saldos Proveedor", exact=False) or no_rows(d), timeout=90, poll=2):
        raise bk.PageError("执行后没有出现 Saldos Proveedor")
    time.sleep(3)
    if no_rows(d):
        return ctx.nothing_there(form, "本月账本没有任何行")
    name = export_excel(ctx, need_text(d, "Export", what="导出")[0], R.LEDGER % periods.stamp(ctx.month))
    if not name:
        return ctx.give_up(form, "download-not-landed", step="export")
    ctx.got(form, [name])


def sales(ctx):
    """Ventas Marketplace: Fecha = the month -> Ejecutar -> Export."""
    form = "sales"
    d = ctx.d
    a, b = periods.month_bounds(ctx.month)
    nav(ctx, "Ventas Marketplace")
    lo, hi = inputs_after(d, "Fecha")
    set_field(ctx, lo, periods.sap_date(a))
    work_area(ctx)
    lo, hi = inputs_after(d, "Fecha")
    set_field(ctx, hi, periods.sap_date(b))
    work_area(ctx)
    bk.real_click(d, need_text(d, "Ejecutar", what="执行")[0])
    if not bk.wait_for(lambda: by_text(d, "Export") or no_rows(d), timeout=120, poll=2):
        raise bk.PageError("执行后没有出现结果表")
    time.sleep(3)
    if not by_text(d, "Export"):
        return ctx.nothing_there(form, "本月没有销售单据")
    name = export_excel(ctx, by_text(d, "Export")[0], R.SALES % periods.stamp(ctx.month))
    if not name:
        return ctx.give_up(form, "download-not-landed", step="export")
    ctx.got(form, [name])


def payments(ctx):
    """Detalle de Pagos: payments made in the month (year/month -> Buscar) and the
    open items, each table exported."""
    form = "payments"
    d = ctx.d
    y, m = periods.parse_month(ctx.month)
    nav(ctx, "Detalle de Pagos")
    year, month = inputs_after(d, "Fecha de pago")
    set_field(ctx, year, str(y))
    work_area(ctx)
    year, month = inputs_after(d, "Fecha de pago")
    set_field(ctx, month, "%02d" % m)
    work_area(ctx)
    bk.real_click(d, need_text(d, "Buscar", what="查询付款")[0])
    time.sleep(10)
    work_area(ctx)
    menus = by_text(d, "Generar Archivo en Excel")
    if len(menus) < 2:
        raise bk.PageError("付款页的导出菜单不够两个")
    names = []
    n = export_excel(ctx, menus[1], R.PAYMENTS % periods.stamp(ctx.month))
    names.append(n)
    work_area(ctx)
    menus = by_text(d, "Generar Archivo en Excel")
    names.append(export_excel(ctx, menus[0], R.PENDING % periods.stamp(ctx.month)))
    if not any(names):
        return ctx.give_up(form, "download-not-landed", step="export")
    names.extend(payment_details(ctx))
    ctx.got(form, names)


def _clearing_links(d):
    """The payments of "Pagos realizados": links whose text is the clearing document."""
    return [e for e in visible(d, "//a") if re.match(r"^\d{8,12}$", (e.text or "").strip())]


def payment_details(ctx):
    """Every payment's breakdown (finance's "佣金和税金 Fees", requirement 2026-11-06):
    click the payment's Doc.comp. -> "Desglose de documento <n>" fills in under
    Detalles -> its own "Generar Archivo en Excel" (the third menu of the page).
    A month without a payment has no link and yields nothing."""
    d = ctx.d
    work_area(ctx)
    docs = [e.text.strip() for e in _clearing_links(d)]
    names = []
    for doc in docs:
        work_area(ctx)
        link = [e for e in _clearing_links(d) if e.text.strip() == doc]
        if not link:
            ctx.errors.append("payments: 付款 %s 的链接找不到了，明细没有下载" % doc)
            continue
        bk.real_click(d, link[0])
        shown = bk.wait_for(lambda: "Desglose de documento %s" % doc in (
            d.execute_script("return document.body.innerText || ''") or ""), timeout=60, poll=2)
        time.sleep(4)
        work_area(ctx)
        menus = by_text(d, "Generar Archivo en Excel")
        if not shown or len(menus) < 3:
            ctx.errors.append("payments: 付款 %s 的明细（Desglose）没有展开，明细没有下载" % doc)
            continue
        n = export_excel(ctx, menus[2], R.PAYMENT_DETAIL % (periods.stamp(ctx.month), doc))
        if n:
            names.append(n)
        else:
            ctx.errors.append("payments: 付款 %s 的明细导出没有落地" % doc)
    ctx.flags["payment_details"] = len(names)
    return names


def _km_next_link(d):
    """The folder pager's ">" (next page) link, or None on the last page. The page
    links read "16-30", "31-45", "46-48" — and, for a last page of one file, not a
    range at all: the October 2026 run that followed only "<n>-<m>" links left the
    46th file (an XML) behind. ">" is a link on every page but the last."""
    for e in visible(d, "//a"):
        if (e.text or "").strip() == ">":
            return e
    return None


def _km_total(d):
    """The folder's file count from the pager's "46-48 / 48", or None."""
    m = re.search(r"\d+\s*-\s*\d+\s*/\s*(\d+)", d.execute_script("return document.body.innerText || ''") or "")
    return int(m.group(1)) if m else None


def invoices(ctx):
    """Notas de Cargo: Liverpool's fee CFDIs (XML + PDF) of the month and of the
    next one (the last week's are issued then), saved into Notas_Cargo_<YYYYMM>/."""
    form = "invoices"
    d = ctx.d
    saved = []
    for mon in (ctx.month, periods.next_month(ctx.month)):
        y, m = periods.parse_month(mon)
        nav(ctx, "Notas de Cargo")
        year = visible(d, "//a[normalize-space(.)='%d']" % y)
        if not year:
            continue
        bk.real_click(d, year[0])
        time.sleep(5)
        work_area(ctx)
        folder = visible(d, "//a[normalize-space(.)='%02d']" % m)
        if not folder:
            continue
        bk.real_click(d, folder[0])
        time.sleep(5)
        work_area(ctx)
        hrefs = set()

        def grab():
            for e in d.find_elements(By.XPATH, "//a[contains(@href,'/Nota_Cargo/')]"):
                h = e.get_attribute("href") or ""
                if "/irj/" in h:
                    hrefs.add(h.split("/irj/", 1)[1])
        grab()
        for _ in range(20):                      # pages of 15; ">" is a link until the last page
            nxt = _km_next_link(d)
            if nxt is None:
                break
            bk.real_click(d, nxt)
            time.sleep(4)
            work_area(ctx)
            grab()
        total = _km_total(d)
        if total is not None and total > len(hrefs):
            ctx.errors.append("invoices: Notas de Cargo %04d/%02d 文件夹显示 %d 个文件，只找到 %d 个"
                              % (y, m, total, len(hrefs)))
        d.switch_to.default_content()
        dest = os.path.join(ctx.out_dir, R.INVOICE_DIR % periods.stamp(mon))
        os.makedirs(dest, exist_ok=True)
        for h in sorted(hrefs):
            b64 = d.execute_async_script(
                "var cb=arguments[arguments.length-1];"
                "fetch('/irj/'+arguments[0],{credentials:'include'}).then(r=>r.arrayBuffer()).then(b=>{"
                "var s='',u=new Uint8Array(b);for(var i=0;i<u.length;i+=32768)"
                "s+=String.fromCharCode.apply(null,u.subarray(i,i+32768));cb(btoa(s));}).catch(e=>cb(''))", h)
            if b64:
                fn = h.rsplit("/", 1)[-1]
                with open(os.path.join(dest, fn), "wb") as fh:
                    fh.write(base64.b64decode(b64))
                saved.append("%s/%s" % (os.path.basename(dest), fn))
    if not saved:
        return ctx.nothing_there(form, "本月没有平台费用发票（Notas de Cargo 里没有这个月的文件夹）")
    ctx.got(form, saved)


# ================================================================ marketplace

def _is_sign_in_page(url):
    return "login" in url or "entradaunica" in url


def marketplace_sign_in(ctx):
    """On the marketplace's sign-in page (Entrada Única: #username, #password,
    button[name=action]) press "Iniciar sesión" once Ziniao has filled the store's
    saved e-mail. The password is Chrome's autofill preview until a real click lands
    on its box, so the box is clicked first. Nothing is typed, nothing is logged.
    Returns True when the marketplace is open again."""
    d = ctx.d
    filled = bk.wait_for(lambda: [e for e in d.find_elements(By.ID, "username")
                                  if (e.get_attribute("value") or "").strip()], timeout=20, poll=2)
    if not filled:
        return False
    pw = [e for e in d.find_elements(By.ID, "password") if e.is_displayed()]
    btn = [b for b in d.find_elements(By.CSS_SELECTOR, "button[name=action]") if b.is_displayed()]
    if not pw or not btn:
        return False
    bk.real_click(d, pw[0])
    time.sleep(1)
    bk.real_click(d, btn[0])
    ok = bk.wait_for(lambda: not _is_sign_in_page(d.current_url or ""), timeout=60, poll=2)
    if not ok:
        return False
    time.sleep(6)                          # the orders page renders after the redirect
    ctx.flags["marketplace_login"] = "ziniao"
    return True


LANG_CONTROL_CSS = "[class*='LanguageSelector'][class*='ContentName']"   # header item showing "es" / "en"
LANG_MENU_CSS = "li[role=menuitem]"                                          # "Español (es)", "Inglés (en)"
WANT_LANG = "es"


def _lang_control(d):
    els = [e for e in d.find_elements(By.CSS_SELECTOR, LANG_CONTROL_CSS) if e.is_displayed()]
    return els[0] if els else None


def _lang_text(d):
    ctl = _lang_control(d)
    return (ctl.text or "").strip().lower() if ctl is not None else None


def _body(d):
    return d.execute_script("return document.body.innerText || ''") or ""


def _menu_click(d, el):
    """Pick an entry of the header's (MUI) menu: move the mouse there, pause, click.
    real_click's scrollIntoView + immediate click leaves the entry unpicked."""
    ActionChains(d).move_to_element(el).pause(0.5).click().perform()


def marketplace_language(ctx):
    """Before any route: the seller centre's language, read from the header control
    ("es" / "en"); when it is not Spanish, switch it back through the control's
    menu — the routes find every button by its Spanish text. Returns (language at
    entry, changed). Raises LanguageWrong when the control cannot be found, the
    menu has no Spanish entry, or the switch did not survive a reload (every page
    load restores the account's own language setting — VM 2026-11-05 — so a
    language that is wrong on a fresh load is the account's setting, which only
    a person can change)."""
    d = ctx.d
    ctl = bk.wait_for(lambda: _lang_control(d), timeout=40, poll=2)
    if ctl is None:
        raise LanguageWrong("Marketplace 顶栏找不到语言切换器（页面没加载出来，或平台改版了）")
    cur = (ctl.text or "").strip().lower()
    if cur == WANT_LANG:
        ctx.flags["marketplace_locale"] = cur
        return cur, False
    ctx.log("  [警告] Marketplace 界面语言是 %s，不是 %s —— 正在改回去" % (cur, WANT_LANG))
    bk.real_click(d, ctl)
    item = bk.wait_for(lambda: [e for e in d.find_elements(By.CSS_SELECTOR, LANG_MENU_CSS)
                                if e.is_displayed() and "(%s)" % WANT_LANG in (e.text or "")],
                       timeout=10, poll=1)
    if not item:
        raise LanguageWrong("Marketplace 语言菜单里找不到 Español (es)（当前 %s）；请手动进店改回 Español" % cur)
    _menu_click(d, item[0])
    now = bk.wait_for(lambda: _lang_text(d) == WANT_LANG and WANT_LANG, timeout=15, poll=1)   # changes in place
    if now:
        ctx.go("/orders", settle=10)                                         # re-rendered in Spanish
        now = bk.wait_for(lambda: _lang_text(d) == WANT_LANG and WANT_LANG, timeout=40, poll=2)
    if now != WANT_LANG:
        now = _lang_text(d)
        raise LanguageWrong("Marketplace 界面语言切换后重新打开页面又是 %s（期望 %s）：顶栏的切换只对当前页面有效，"
                            "每次打开页面都恢复成账号自己的语言设置；请在 Marketplace 的账号设置里把语言改成 Español"
                            % (now, WANT_LANG))
    ctx.flags["marketplace_locale"] = cur
    ctx.errors.append("language: Marketplace 界面语言原本是 %s，已改回 %s。请查清是谁改的" % (cur, WANT_LANG))
    ctx.log("  Marketplace 界面语言已从 %s 改回 %s ✓" % (cur, WANT_LANG))
    return cur, True


def marketplace_open(ctx):
    """The marketplace seller centre, signed in (by the session saved in the Ziniao
    store, or by pressing the sign-in button on the page Ziniao has filled), in
    Spanish (switched back when it is not), with the orders page showing."""
    d = ctx.d
    real_window(d)
    ctx.go("/orders", settle=12)
    url = d.current_url or ""
    if _is_sign_in_page(url):
        if not marketplace_sign_in(ctx):
            raise Blocked("Marketplace 被带到了登录页，紫鸟没有填入账号密码或登录没有成功（请手动进店登录一次）：%s"
                          % url[:100])
    else:
        ctx.flags["marketplace_login"] = "session"
    if "dispositivo móvil" in _body(d):
        raise bk.PageError("Marketplace 以为是手机（窗口不对）")
    marketplace_language(ctx)
    if not bk.wait_for(lambda: "Pedidos" in _body(d), timeout=30, poll=2):
        raise LanguageWrong("Marketplace 界面语言是西班牙语，但找不到 Pedidos（页面没加载出来，或平台改版了）")


def _button(d, startswith):
    els = [e for e in d.find_elements(By.CSS_SELECTOR, "button") if e.is_displayed()
           and (e.text or "").strip().startswith(startswith)]
    if not els:
        raise bk.PageError("找不到按钮「%s」" % startswith)
    return els[0]


def _xlsx(ctx):
    opt = bk.wait_for(lambda: by_text(ctx.d, "XLSX"), timeout=10, poll=1)
    if not opt:
        raise bk.PageError("下载菜单里没有 XLSX")
    bk.real_click(ctx.d, opt[-1])


def _results(d):
    m = re.search(r"Mostrando los resultados [\d\s\-–]+ de ([\d,]+)",
                  d.execute_script("return document.body.innerText || ''") or "")
    return int(m.group(1).replace(",", "")) if m else None


def _apply_filters(ctx, fields):
    """Filtros -> type the dates (DD/MM/YYYY, digit by digit) -> Aplicar filtros."""
    d = ctx.d
    bk.real_click(d, _button(d, "Filtros"))
    time.sleep(3)
    for el, text in fields(d):
        bk.real_click(d, el)
        time.sleep(0.4)
        type_slowly(el, text)
    ActionChains(d).send_keys(Keys.ESCAPE).perform()
    time.sleep(1)
    bk.real_click(d, _button(d, "Aplicar filtros"))
    time.sleep(8)


def orders(ctx):
    """Pedidos: the order list as the page offers it (the last four months) -> XLSX."""
    form = "orders"
    d = ctx.d
    ctx.go("/orders", settle=12)

    def click():
        bk.real_click(d, _button(d, "Descargar"))
        time.sleep(3)
        _xlsx(ctx)
    name = download(ctx, click, R.ORDERS.replace("*", ctx.today.strftime("%Y%m%d")))
    if not name:
        return ctx.give_up(form, "download-not-landed", step="descargar")
    ctx.got(form, [name])


def logistics(ctx):
    """Costos logísticos: Fecha de cobro = the month -> Descargar tabla -> XLSX."""
    form = "logistics"
    d = ctx.d
    a, b = periods.month_bounds(ctx.month)
    ctx.go("/orders/admin/logistic-costs", settle=12)

    def fields(d):
        return [(d.find_element(By.CSS_SELECTOR, "input[name='charged_at_start']"), periods.web_date(a)),
                (d.find_element(By.CSS_SELECTOR, "input[name='charged_at_end']"), periods.web_date(b))]
    _apply_filters(ctx, fields)
    got = [d.find_element(By.CSS_SELECTOR, "input[name='%s']" % n).get_attribute("value")
           for n in ("charged_at_start", "charged_at_end")] if d.find_elements(
               By.CSS_SELECTOR, "input[name='charged_at_start']") else None
    if got and got != [a.strftime("%d/%m/%Y"), b.strftime("%d/%m/%Y")]:
        raise bk.PageError("扣费日期没有填进去：%s" % got)
    if _results(d) == 0:
        return ctx.nothing_there(form, "本月没有物流费")

    def click():
        bk.real_click(d, _button(d, "Descargar"))
        time.sleep(3)
        _xlsx(ctx)
    name = download(ctx, click, R.LOGISTICS % periods.stamp(ctx.month))
    if not name:
        return ctx.give_up(form, "download-not-landed", step="descargar")
    ctx.got(form, [name])


def balance(ctx):
    """Mi tienda -> Balance de cuenta: Fecha real de cobro = the month -> Descargar ->
    XLSX. The drawer has two date pairs with the same input names (from / to):
    the first is "Fecha de registro", the last "Fecha real de cobro" — finance's
    requirement (2026-11-09) is the charge date, like Costos logísticos."""
    form = "balance"
    d = ctx.d
    a, b = periods.month_bounds(ctx.month)
    ctx.go("/stores/balance", settle=12)

    def fields(d):
        fr = [e for e in d.find_elements(By.CSS_SELECTOR, "input[name='from']") if e.is_displayed()]
        to = [e for e in d.find_elements(By.CSS_SELECTOR, "input[name='to']") if e.is_displayed()]
        if not fr or not to:
            raise bk.PageError("找不到日期的输入框")
        if len(fr) < 2 or len(to) < 2:
            raise bk.PageError("筛选里找不到「Fecha real de cobro」（日期输入框只有一组，平台改版了？）")
        return [(fr[-1], periods.web_date(a)), (to[-1], periods.web_date(b))]
    _apply_filters(ctx, fields)
    ctx.flags["balance_filter"] = "fecha_real_de_cobro"
    if _results(d) == 0:
        return ctx.nothing_there(form, "本月没有补充物流费或其他扣款")

    def click():
        bk.real_click(d, _button(d, "Descargar"))
        time.sleep(3)
        _xlsx(ctx)
    name = download(ctx, click, R.BALANCE % periods.stamp(ctx.month))
    if not name:
        return ctx.give_up(form, "download-not-landed", step="descargar")
    ctx.got(form, [name])


# route name (periods.ROUTES) -> function(ctx)
ROUTES = {"ledger": ledger, "sales": sales, "payments": payments, "invoices": invoices,
          "orders": orders, "logistics": logistics, "balance": balance}


def collect_owed(ctx, entry):
    """Liverpool has no asynchronous reports: nothing is ever owed."""
    return []
