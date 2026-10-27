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
(NeedsLogin -> PAGE_BLOCKED). The marketplace's sign-in is never touched: its
session lives in the Ziniao store.
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
    ctx.got(form, names)


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
        for _ in range(20):                      # pages of 15: "16-30", "31-45", … then ">"
            nxt = [e for e in visible(d, "//a") if re.match(r"^\d+-\d+$", (e.text or "").strip())
                   and int(e.text.split("-")[0]) > len(hrefs)]
            if not nxt:
                break
            bk.real_click(d, nxt[0])
            time.sleep(4)
            work_area(ctx)
            grab()
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

def marketplace_open(ctx):
    """The marketplace seller centre, signed in (the session lives in the Ziniao store)."""
    d = ctx.d
    real_window(d)
    ctx.go("/orders", settle=12)
    url = d.current_url or ""
    if "login" in url or "entradaunica" in url:
        raise Blocked("Marketplace 被带到了登录页：%s" % url[:100])
    body = d.execute_script("return document.body.innerText || ''") or ""
    if "dispositivo móvil" in body:
        raise bk.PageError("Marketplace 以为是手机（窗口不对）")
    if "Pedidos" not in body:
        raise LanguageWrong("Marketplace 界面语言不是西班牙语（找不到 Pedidos）")


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
    """Mi tienda -> Balance de cuenta: Fecha de registro = the month -> Descargar -> XLSX."""
    form = "balance"
    d = ctx.d
    a, b = periods.month_bounds(ctx.month)
    ctx.go("/stores/balance", settle=12)

    def fields(d):
        fr = [e for e in d.find_elements(By.CSS_SELECTOR, "input[name='from']") if e.is_displayed()]
        to = [e for e in d.find_elements(By.CSS_SELECTOR, "input[name='to']") if e.is_displayed()]
        if not fr or not to:
            raise bk.PageError("找不到登记日期的输入框")
        return [(fr[0], periods.web_date(a)), (to[0], periods.web_date(b))]
    _apply_filters(ctx, fields)
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
