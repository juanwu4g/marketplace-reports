# -*- coding: utf-8 -*-
"""Browser helpers for the Falabella Seller Center.

Why this module exists (FORMS.md "How the pages are built"):

* The new pages are micro-frontends inside **open shadow roots**
  (div#orders, div#accountstatements-app-container-wrapper,
  div#settlement-liquidation-invoice-container). document.querySelector sees
  none of their content, so every lookup here walks the shadow roots.
* Ant Design calendar cells ignore a JavaScript click; they need a real mouse
  click. Buttons and menu items take either, and a JavaScript click goes through
  the success notifications that sometimes cover them.
* Pop-ups appear late and block clicks: the feedback modal on the payments page,
  the guided tour on the orders page, the dashboard survey.

Everything here finds elements by data-testid / role / id first and by visible
Spanish text only where nothing else exists. The language gate (page_locale)
runs before any route, so the Spanish texts are safe.
"""
import os
import time

from selenium.webdriver.common.action_chains import ActionChains

# Walk the document and every open shadow root below it.
WALK = ("function W(r,f){r.querySelectorAll('*').forEach(function(e){f(e);"
        "if(e.shadowRoot)W(e.shadowRoot,f);});}")

_FIND = WALK + """
var sel=arguments[0], want=arguments[1], all=arguments[2], out=[];
W(document,function(e){
  if(!e.matches||!e.matches(sel))return;
  if(!all){var b=e.getBoundingClientRect(); if(!(b.width>0&&b.height>0))return;}
  if(want!==null){var t=(e.innerText||e.textContent||'').trim().replace(/\\s+/g,' ');
                  if(t!==want)return;}
  out.push(e);});
return out;"""


class PageError(RuntimeError):
    """A control the route needs is not there. The message says which one.

    RuntimeError with "not found" in the text: download_diag classifies it as
    UI_NOT_FOUND (a selector to fix), not as a script bug.
    """


def find(d, selector, text=None, hidden=False):
    """Visible elements matching selector (shadow roots included), optionally
    with this exact trimmed text."""
    try:
        return d.execute_script(_FIND, selector, text, bool(hidden)) or []
    except Exception:
        return []


def wait_for(fn, timeout=30, poll=1.0):
    """Call fn until it returns something truthy; return it, or None at timeout."""
    end = time.time() + timeout
    while True:
        try:
            v = fn()
        except Exception:
            v = None
        if v:
            return v
        if time.time() >= end:
            return None
        time.sleep(poll)


def need(d, selector, text=None, timeout=30, what=None):
    """The first matching element, waiting for it; PageError if it never shows."""
    got = wait_for(lambda: find(d, selector, text), timeout=timeout)
    if not got:
        raise PageError("%s not found (%s%s)" % (
            what or "control", selector, " = %r" % text if text else ""))
    return got[0]


def js_click(d, el):
    d.execute_script("arguments[0].click()", el)


def real_click(d, el):
    """Scroll into view and click with the mouse (what calendars need)."""
    try:
        d.execute_script("arguments[0].scrollIntoView({block:'center'})", el)
    except Exception:
        pass
    time.sleep(0.3)
    ActionChains(d).move_to_element(el).click().perform()


def shadow(d, host_selector, inner_selector, all_=False):
    """querySelector(All) inside one known shadow root (the orders page)."""
    script = ("var h=document.querySelector(arguments[0]);"
              "if(!h||!h.shadowRoot)return null;"
              "return arguments[2]?h.shadowRoot.querySelectorAll(arguments[1])"
              ":h.shadowRoot.querySelector(arguments[1]);")
    try:
        return d.execute_script(script, host_selector, inner_selector, bool(all_))
    except Exception:
        return None


# ------------------------------------------------------------------ pop-ups

FEEDBACK_TEXT = "Qué te pareció"


def close_popups(d):
    """Close whatever known pop-up is in the way. Cheap; call before clicks."""
    closed = 0
    for m in find(d, ".ant-modal-wrap .ant-modal-content"):
        try:
            if FEEDBACK_TEXT in (m.text or ""):
                for b in find(d, ".ant-modal-wrap button", "Cancelar"):
                    js_click(d, b)
                    closed += 1
        except Exception:
            pass
    for b in find(d, "button", "Omitir"):          # guided tour on the orders page
        try:
            js_click(d, b)
            closed += 1
        except Exception:
            pass
    for b in find(d, "#dismiss"):                  # dashboard survey
        try:
            js_click(d, b)
            closed += 1
        except Exception:
            pass
    if closed:
        time.sleep(1)
    return closed


# ------------------------------------------------------------------ facts

_LOCALE = WALK + """var o=null;W(document,function(e){if(!o&&e.getAttribute&&
 e.getAttribute('data-localefull'))o={locale:e.getAttribute('data-localefull'),
 seller:e.getAttribute('data-seller'),name:e.getAttribute('data-sellername')};});return o;"""


def page_locale(d):
    """{'locale': 'es-CL', 'seller': 'SC2C911', 'name': 'Hogaria'} or None.

    Read from the micro-frontend root, not from <html lang> — that says "en" on
    these pages whatever the screen language is.
    """
    try:
        return d.execute_script(_LOCALE)
    except Exception:
        return None


_ROWS = WALK + """var re=new RegExp(arguments[0]),o=[];W(document,function(e){
 if(e.tagName==='TR'){var t=(e.innerText||'').replace(/\\s+/g,' ').trim();
 if(re.test(t))o.push(t);}});return o;"""


def table_rows(d, pattern=r"\S"):
    """Text of every table row (shadow roots included) matching a regex."""
    try:
        return d.execute_script(_ROWS, pattern) or []
    except Exception:
        return []


_ROW = WALK + """var want=arguments[0],hit=null;W(document,function(e){
 if(!hit&&e.tagName==='TR'&&(e.innerText||'').replace(/\\s+/g,' ').trim()
 .indexOf(want)===0)hit=e;});return hit;"""


def row_starting(d, text):
    """The first table row whose text starts with `text`."""
    try:
        return d.execute_script(_ROW, text)
    except Exception:
        return None


_ROW_HAS = WALK + """var want=arguments[0],hit=null;W(document,function(e){
 if(!hit&&e.tagName==='TR'&&(e.innerText||'').indexOf(want)>=0)hit=e;});return hit;"""


def row_containing(d, text):
    try:
        return d.execute_script(_ROW_HAS, text)
    except Exception:
        return None


def inside(d, el, selector, all_=False):
    try:
        return d.execute_script(
            "return arguments[2]?arguments[0].querySelectorAll(arguments[1])"
            ":arguments[0].querySelector(arguments[1])", el, selector, bool(all_))
    except Exception:
        return None


# ------------------------------------------------------------------ downloads table

DOWNLOAD_ROW = r"\.(xlsx|csv)|Procesando|Listo|Error"


def download_rows(d):
    """Rows of a "Descargas / Documents descargados" table, newest first."""
    return table_rows(d, DOWNLOAD_ROW)


def refresh_downloads(d):
    for b in find(d, "button", "Actualizar"):
        js_click(d, b)
        return True
    return False


def wait_new_ready_row(d, before, timeout=240, poll=6, prefix=None):
    """Wait until the newest downloads row is a new one and 'Listo'.

    before   the newest row text before the request (or None)
    prefix   only rows whose file name starts with this count
    Returns (row_text, state) — state 'ready', 'generating', 'absent', 'error'.
    """
    end = time.time() + timeout
    last = None
    while True:
        rows = [r for r in download_rows(d) if not prefix or r.startswith(prefix)]
        top = rows[0] if rows else None
        if top and top != before:
            last = top
            if "Listo" in top:
                return top, "ready"
            if "Error" in top:
                return top, "error"
        if time.time() >= end:
            return last, ("generating" if last else "absent")
        refresh_downloads(d)
        time.sleep(poll)


def file_name_of_row(row_text):
    return (row_text or "").split(" ")[0]


# ------------------------------------------------------------------ files

PARTIAL = (".crdownload", ".tmp", ".part")


def set_download_dir(d, path):
    """Route this browser's downloads into `path` (absolute; see run_downloads)."""
    os.makedirs(path, exist_ok=True)
    d.execute_cdp_cmd("Browser.setDownloadBehavior",
                      {"behavior": "allow", "downloadPath": os.path.abspath(path)})


def snapshot(path):
    try:
        return set(os.listdir(path))
    except OSError:
        return set()


def wait_new_files(path, before, timeout=90, settle=3, start_timeout=None):
    """New complete files in `path` since `before`, once the folder is quiet.

    start_timeout: give up early when no download (not even a partial file) has
    started after that many seconds.
    """
    began = time.time()
    end = began + timeout
    seen, changed = set(), time.time()
    while time.time() < end:
        time.sleep(1)
        new = snapshot(path) - set(before)
        if start_timeout and not new and time.time() - began >= start_timeout:
            return []
        done = {f for f in new if not f.endswith(PARTIAL)}
        partial = [f for f in new if f.endswith(PARTIAL)]
        if done != seen:
            seen, changed = done, time.time()
        if seen and not partial and time.time() - changed >= settle:
            return sorted(seen)
    return sorted(seen)
