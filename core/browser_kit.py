# -*- coding: utf-8 -*-
"""Browser helpers every platform's routes share (Selenium on a Ziniao browser).

Why these exist (lessons from the first two platforms, see
docs/PLATFORM_PLAYBOOK.zh.md):

* Modern seller centres are built from micro-frontends inside **open shadow
  roots** (Falabella's orders and payments pages are an example).
  document.querySelector sees none of their content, so every lookup here walks
  the document and every open shadow root below it.
* Some widgets (calendar cells of Ant Design, for instance) ignore a JavaScript
  click and need a real mouse click; buttons and menu items take either, and a
  JavaScript click goes through notifications that sometimes cover them. Hence
  both js_click and real_click.
* Downloads land in a folder the run chooses (set_download_dir); a file is only
  complete when the browser has renamed it from its partial name (PARTIAL), and
  the folder has been quiet for a moment (wait_new_files).

Nothing here knows a platform: no URLs, no labels in a site's language. A
platform's own layer (e.g. platforms/falabella/page_kit.py) adds pop-ups,
language gates and download-list wording on top, and re-exports these.

Every function that talks to the page swallows WebDriver errors and returns an
empty answer: the routes decide what a missing element means (usually
PageError, which download_diag classifies as UI_NOT_FOUND).
"""
import os
import time

# The document and every open shadow root below it, depth first.
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


# ------------------------------------------------------------------ finding

def find(d, selector, text=None, hidden=False):
    """Visible elements matching selector (shadow roots included), optionally
    with this exact trimmed, whitespace-collapsed text. hidden=True also returns
    elements that have no size on screen."""
    try:
        return d.execute_script(_FIND, selector, text, bool(hidden)) or []
    except Exception:
        return []


def wait_for(fn, timeout=30, poll=1.0):
    """Call fn until it returns something truthy; return it, or None at timeout.
    An exception inside fn counts as "not yet"."""
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


def shadow(d, host_selector, inner_selector, all_=False):
    """querySelector(All) inside one known shadow root (host found in the document)."""
    script = ("var h=document.querySelector(arguments[0]);"
              "if(!h||!h.shadowRoot)return null;"
              "return arguments[2]?h.shadowRoot.querySelectorAll(arguments[1])"
              ":h.shadowRoot.querySelector(arguments[1]);")
    try:
        return d.execute_script(script, host_selector, inner_selector, bool(all_))
    except Exception:
        return None


def inside(d, el, selector, all_=False):
    """querySelector(All) below one element already found."""
    try:
        return d.execute_script(
            "return arguments[2]?arguments[0].querySelectorAll(arguments[1])"
            ":arguments[0].querySelector(arguments[1])", el, selector, bool(all_))
    except Exception:
        return None


# ------------------------------------------------------------------ clicking

def js_click(d, el):
    d.execute_script("arguments[0].click()", el)


def real_click(d, el):
    """Scroll into view and click with the mouse (what calendars need)."""
    from selenium.webdriver.common.action_chains import ActionChains
    try:
        d.execute_script("arguments[0].scrollIntoView({block:'center'})", el)
    except Exception:
        pass
    time.sleep(0.3)
    ActionChains(d).move_to_element(el).click().perform()


# ------------------------------------------------------------------ tables

_ROWS = WALK + """var re=new RegExp(arguments[0]),o=[];W(document,function(e){
 if(e.tagName==='TR'){var t=(e.innerText||'').replace(/\\s+/g,' ').trim();
 if(re.test(t))o.push(t);}});return o;"""


def table_rows(d, pattern=r"\S"):
    """Text of every table row (shadow roots included) matching a regex, in page order."""
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
    """The first table row whose text contains `text`."""
    try:
        return d.execute_script(_ROW_HAS, text)
    except Exception:
        return None


# ------------------------------------------------------------------ files

# Names browsers give a download while it is still being written.
PARTIAL = (".crdownload", ".tmp", ".part")


def set_download_dir(d, path):
    """Route this browser's downloads into `path`.

    Always an absolute path: Chrome resolves a relative one against its own
    working directory, which is Ziniao's, not ours.
    """
    os.makedirs(path, exist_ok=True)
    d.execute_cdp_cmd("Browser.setDownloadBehavior",
                      {"behavior": "allow", "downloadPath": os.path.abspath(path)})


def snapshot(path):
    """The file names in `path` now (empty set if the folder is not there)."""
    try:
        return set(os.listdir(path))
    except OSError:
        return set()


def wait_new_files(path, before, timeout=90, settle=3, start_timeout=None):
    """New complete files in `path` since `before`, once the folder is quiet.

    before         a snapshot() taken just before the click
    settle         seconds with no change and no partial file before answering
    start_timeout  give up early when no download (not even a partial file) has
                   started after that many seconds
    Returns the sorted new names; at timeout, whatever was complete by then.
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


# ------------------------------------------------------------------ retry

def retried(ctx, what, fn, *a, attempts=2, retry_on=(PageError,), never_retry=(),
            pause=3):
    """Run fn(ctx, *a) again when it fails, up to `attempts` times; the last
    failure is raised.

    Meant for one page interaction that starts by loading its page, so a second
    try starts from a freshly rendered page. Calendars and menus sometimes
    swallow a click (Falabella's order page on 2026-10-19 kept the previous
    run's range); a second try is cheap and nearly always enough.

    retry_on     exception types worth another try (default: a control was missing)
    never_retry  exception types raised at once even if they also match retry_on
                 — store-level problems such as a wrong UI language, where trying
                 again only wastes time
    ctx.log      (if present) receives one Chinese line per failed try
    """
    log = getattr(ctx, "log", None) or print
    for i in range(attempts):
        try:
            return fn(ctx, *a)
        except tuple(never_retry):
            raise
        except tuple(retry_on) as e:
            if i == attempts - 1:
                raise
            log("  %s 第 %d 次没成功（%s），重新打开页面再试" % (what, i + 1, str(e)[:120]))
            time.sleep(pause)
