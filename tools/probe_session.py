# -*- coding: utf-8 -*-
"""Discovery session: open ONE store once, then run small scripts against it.

    python -m tools.probe_session --store NAME --cmd-dir DIR --out DIR [--max-minutes 120]

Discovery of a new platform is a conversation with the page: look, click one
thing, look again. Opening the store for every step costs a minute and loses the
page's state, so this keeps one store open and runs each script dropped into
the command folder:

  1. opens the store through core.ziniao_client (which refuses any store not on
     the config.json allowlist — checked here too, before Ziniao is touched)
  2. routes downloads into <out>/files (an absolute path)
  3. writes READY into the command folder
  4. for every new <name>.py that appears there (in name order), runs it and
     writes <name>.out: everything it printed, the traceback if it failed, and a
     last line "<<DONE>>" — wait for that line before reading the file
  5. stops when a file named STOP appears or after --max-minutes, and closes
     the store through the Ziniao client (never driver.quit())

A script runs with these names already defined (one namespace for the whole
session, so a variable set by one script is there for the next):

    d, driver    the Selenium driver of the open store
    out, files   the output folder and the download folder (absolute)
    dump(d, tag) screenshot + JSON (url, title, visible text including text in
                 shadow roots, the controls with their attributes) into <out>
    js(d, src, *args)  execute_script that returns "JSERR ..." instead of raising
    every helper of core/browser_kit (find, need, wait_for, js_click,
    real_click, shadow, table_rows, snapshot, wait_new_files, WALK …)
    By, Keys, ActionChains, os, re, json, time

Example command file 001_look.py:

    d.get("https://seller.example.com/reports")
    time.sleep(10)
    dump(d, "reports")
    print([b.text for b in find(d, "button")])

Scripts that hang hang the session (there is no time limit per script); keep
each one small. Files already in the command folder when the session starts are
treated as done, never re-run. This is the "probe, don't guess" step of
docs/NEW_PLATFORM_SOP.md §2: run it on the test machine only.
"""
import argparse
import contextlib
import io
import json
import os
import re
import sys
import time
import traceback

from core import browser_kit
from core import console  # noqa: F401  中文输出编码保护
from core import store_config

READY = "READY"
STOP = "STOP"
DONE_MARK = "<<DONE>>"

# Everything a person needs to write the next step: visible text of the document
# and of every shadow root (style / script / link contents left out), and every
# control with the attributes selectors are made of.
DUMP_JS = """
var out={hosts:[],controls:[],text:[]};
function vis(e){var r=e.getBoundingClientRect();return r.width>0&&r.height>0;}
function attrs(e){var a={};for(var i=0;i<e.attributes.length;i++){var n=e.attributes[i].name;
 if(/^(id|role|name|type|href|aria-|data-|placeholder|value|for|title)/.test(n))
 a[n]=e.attributes[i].value.slice(0,100);}return a;}
function own(e){return (e.innerText||e.textContent||'').trim().replace(/\\s+/g,' ').slice(0,100);}
var SKIP={STYLE:1,SCRIPT:1,LINK:1,NOSCRIPT:1,TEMPLATE:1};
function text(root){return [].map.call(root.children,function(c){
 return SKIP[c.tagName]?'':(c.innerText||'');}).join('\\n').trim();}
function walk(root,depth){
  var all=root.querySelectorAll('*');
  for(var i=0;i<all.length;i++){var e=all[i];
    if(e.shadowRoot){out.hosts.push(e.tagName.toLowerCase()+(e.id?'#'+e.id:'')+' d'+depth);
                     walk(e.shadowRoot,depth+1);}
    if(e.matches&&e.matches('a,button,input,select,textarea,[role=button],[role=tab],'+
       '[role=option],[role=menuitem],[data-testid]')&&out.controls.length<400)
      out.controls.push({d:depth,tag:e.tagName.toLowerCase(),visible:vis(e),text:own(e),
        cls:(e.className||'').toString().slice(0,60),attrs:attrs(e)});}
  var t=depth>0?text(root):(document.body?text(document.body):'');
  if(t)out.text.push('[d'+depth+'] '+t.slice(0,8000));
}
walk(document,0);
return out;"""


def js(d, src, *args):
    """execute_script that never raises: a broken probe script is data too."""
    try:
        return d.execute_script(src, *args)
    except Exception as e:
        return "JSERR " + repr(e)[:300]


def dump(d, tag, out_dir, extra=None, seq=None):
    """Save <out_dir>/NN_<tag>.png and .json (url, title, deep page text,
    controls). Returns the JSON path. seq: a one-item list used as counter."""
    os.makedirs(out_dir, exist_ok=True)
    seq = seq if seq is not None else [len([f for f in os.listdir(out_dir)
                                            if f.endswith(".json")])]
    n = seq[0]
    seq[0] += 1
    safe = re.sub(r"[^\w.-]+", "_", str(tag))[:60] or "page"
    base = os.path.join(out_dir, "%02d_%s" % (n, safe))
    info = {"tag": tag}
    for key, get in (("url", lambda: d.current_url), ("title", lambda: d.title)):
        try:
            info[key] = get()
        except Exception as e:
            info[key] = "ERR " + repr(e)[:200]
    info["page"] = js(d, DUMP_JS)
    if extra:
        info.update(extra)
    try:
        d.save_screenshot(base + ".png")
    except Exception as e:
        info["screenshot_error"] = repr(e)[:200]
    with io.open(base + ".json", "w", encoding="utf-8") as f:
        f.write(json.dumps(info, ensure_ascii=False, indent=1))
    print(n, tag, info.get("url"), flush=True)
    return base + ".json"


def namespace(driver, out_dir, files_dir):
    """The names a command script sees (see the module docstring)."""
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    ns = {k: getattr(browser_kit, k) for k in dir(browser_kit) if not k.startswith("_")
          and k not in ("os", "time")}
    seq = [0]
    ns.update(d=driver, driver=driver, out=out_dir, files=files_dir, js=js,
              dump=lambda d, tag, extra=None: dump(d, tag, out_dir, extra, seq),
              By=By, Keys=Keys, ActionChains=ActionChains,
              os=os, re=re, json=json, time=time, __name__="probe_command")
    return ns


def run_command(path, ns):
    """Run one command file in `ns`; write <name>.out ending with DONE_MARK."""
    buf = io.StringIO()
    try:
        with io.open(path, encoding="utf-8") as f:
            src = f.read()
        code = compile(src, path, "exec")
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            exec(code, ns)
    except BaseException as e:                 # keep the session alive whatever the script did
        if isinstance(e, KeyboardInterrupt):
            raise
        traceback.print_exc(file=buf)
    out = path[:-3] + ".out"
    with io.open(out, "w", encoding="utf-8") as f:
        f.write(buf.getvalue().rstrip("\n") + "\n" + DONE_MARK + "\n")
    return out


def pending(cmd_dir, done):
    return [f for f in sorted(os.listdir(cmd_dir)) if f.endswith(".py") and f not in done]


def serve(cmd_dir, ns, max_seconds, poll=1.0, log=print):
    """Watch cmd_dir and run each new .py until STOP or max_seconds. Returns the
    list of command files run, in order."""
    done = set(pending(cmd_dir, set()))       # left over from an earlier session
    for f in (READY, STOP):
        try:
            os.remove(os.path.join(cmd_dir, f))
        except OSError:
            pass
    with io.open(os.path.join(cmd_dir, READY), "w", encoding="utf-8") as f:
        f.write("ok %s\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
    ran = []
    end = time.time() + max_seconds
    while time.time() < end and not os.path.exists(os.path.join(cmd_dir, STOP)):
        for name in pending(cmd_dir, done):
            done.add(name)
            log("运行 %s" % name)
            run_command(os.path.join(cmd_dir, name), ns)
            ran.append(name)
        time.sleep(poll)
    log("探查结束（%s）" % ("收到 STOP" if os.path.exists(os.path.join(cmd_dir, STOP))
                          else "到时间了"))
    return ran


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--store", required=True, help="exact Ziniao store name (on the allowlist)")
    ap.add_argument("--cmd-dir", required=True, help="folder to drop <name>.py commands into")
    ap.add_argument("--out", required=True, help="folder for dumps; downloads go to <out>/files")
    ap.add_argument("--max-minutes", type=float, default=120)
    a = ap.parse_args(argv)

    # Refuse before Ziniao is even started (ZiniaoClient.find_store checks again).
    store_config.require_allowed(a.store)
    cmd_dir = os.path.abspath(a.cmd_dir)
    out_dir = os.path.abspath(a.out)
    files_dir = os.path.join(out_dir, "files")
    os.makedirs(cmd_dir, exist_ok=True)
    os.makedirs(files_dir, exist_ok=True)

    from core.ziniao_client import ZiniaoClient
    client = ZiniaoClient()
    if not client.is_up():
        client.start(restart=True)
    store = client.find_store(a.store)
    with client.open_store(store) as s:
        d = s.driver
        browser_kit.set_download_dir(d, files_dir)
        print("店铺已打开：%s；命令目录 %s" % (a.store, cmd_dir), flush=True)
        serve(cmd_dir, namespace(d, out_dir, files_dir), a.max_minutes * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
