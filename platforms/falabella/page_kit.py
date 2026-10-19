# -*- coding: utf-8 -*-
"""Browser helpers for the Falabella Seller Center.

The platform-independent helpers (shadow-root walk, find / need / wait_for,
js_click / real_click, table rows, download folder, PageError) live in
core/browser_kit.py and are re-exported here, so routes write pk.find,
pk.PageError ... whichever layer a helper comes from. This module keeps only
what is Falabella's own (FORMS.md "How the pages are built"):

* The new pages are micro-frontends inside **open shadow roots**
  (div#orders, div#accountstatements-app-container-wrapper,
  div#settlement-liquidation-invoice-container) — hence browser_kit's walk.
* Ant Design calendar cells ignore a JavaScript click; they need real_click.
* Pop-ups appear late and block clicks: the feedback modal on the payments page,
  the guided tour on the orders page, the dashboard survey (close_popups).
* The page language is read from data-localefull (page_locale) and switched in
  the footer box (switch_language); the "Descargas" table says Procesando /
  Listo / Error (download_rows, wait_new_ready_row).

Everything here finds elements by data-testid / role / id first and by visible
Spanish text only where nothing else exists. The language gate (page_locale)
runs before any route, so the Spanish texts are safe.
"""
import time

# The generic part (shadow-root walk, find/need/wait_for, clicks, table rows,
# download folder) is shared by every platform; it is re-exported here so the
# routes keep writing pk.find, pk.PageError, pk.wait_new_files ...
from core.browser_kit import (  # noqa: F401
    WALK, PageError, find, wait_for, need, js_click, real_click, shadow, inside,
    table_rows, row_starting, row_containing,
    PARTIAL, set_download_dir, snapshot, wait_new_files)


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


# ------------------------------------------------------------------ language

# The footer of every seller-centre page has a language box: label "Idioma" (or
# "Language" when the account is in English) and a box showing the current
# language. A JavaScript click on the box's inner element opens a small menu above
# it with the options (English / Español); a mouse click toggles it shut again, so
# use JavaScript here. Choosing an option changes the account's language (it
# persists for everyone using the store) and re-renders the page.
_LANG_BOX = """var labels=arguments[0],hit=null;document.querySelectorAll('div,span').forEach(function(e){
 if(!hit&&labels.indexOf((e.innerText||'').trim())>=0&&e.nextElementSibling)hit=e;});
 return hit?(hit.nextElementSibling.firstElementChild||hit.nextElementSibling):null;"""
_LANG_OPTION = """var want=arguments[0],hit=null;document.querySelectorAll('*').forEach(function(e){
 if(!hit&&e.children.length===0&&(e.innerText||'').trim()===want){var b=e.getBoundingClientRect();
 if(b.width>0&&b.height>0)hit=e;}});return hit;"""
LANG_LABELS = ["Idioma", "Language"]


def switch_language(d, option="Español", timeout=40):
    """Choose `option` in the footer language box. Returns the page locale after
    the switch (page_locale()), or None if the box or the option was not found."""
    box = d.execute_script(_LANG_BOX, LANG_LABELS)
    if box is None:
        return None
    try:
        d.execute_script("arguments[0].scrollIntoView({block:'center'})", box)
    except Exception:
        pass
    d.execute_script("arguments[0].click()", box)
    opt = wait_for(lambda: d.execute_script(_LANG_OPTION, option), timeout=5, poll=0.5)
    if opt is None:
        return None
    d.execute_script("arguments[0].click()", opt)
    time.sleep(4)
    return wait_for(lambda: (page_locale(d) or {}).get("locale", "").lower().startswith(
        option[:2].lower()) and page_locale(d), timeout=timeout, poll=2)
