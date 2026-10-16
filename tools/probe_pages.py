# -*- coding: utf-8 -*-
"""Discovery probe: open one store, visit pages, save what each page shows.

    python -m tools.probe_pages --store IDG_PE --out probe_out URL [URL ...]
    python -m tools.probe_pages --store IDG_PE --out probe_out --js grab.js URL ...

For each URL it saves, in --out:
    pNN.png    screenshot
    pNN.json   asked URL, final URL, title, the first 6000 characters of visible
               text, and (with --js) whatever the script returns

This is the "probe, don't guess" step of docs/NEW_PLATFORM_SOP.md §2. It writes
nothing to the store and downloads nothing. The store must be on the
config.json allowlist (ZiniaoClient refuses others) — use a test machine whose
allowlist holds only the stores you are exploring.
"""
import argparse
import io
import json
import os
import sys
import time

from core import console  # noqa: F401  中文输出编码保护
from core.ziniao_client import ZiniaoClient


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", required=True, help="exact Ziniao store name")
    ap.add_argument("--out", required=True, help="folder for the dumps")
    ap.add_argument("--js", default=None,
                    help="file with JavaScript to run on every page; its return value is saved")
    ap.add_argument("--wait", type=float, default=8, help="seconds to let each page settle")
    ap.add_argument("urls", nargs="+")
    a = ap.parse_args(argv)

    os.makedirs(a.out, exist_ok=True)
    js = io.open(a.js, encoding="utf-8").read() if a.js else None
    client = ZiniaoClient()
    if not client.is_up():
        client.start(restart=True)
    store = client.find_store(a.store)          # raises if not on the allowlist
    with client.open_store(store) as s:
        d = s.driver
        for i, url in enumerate(a.urls):
            info = {"asked": url}
            try:
                d.get(url)
                time.sleep(a.wait)
                info.update(url=d.current_url, title=d.title, text=d.execute_script(
                    "return document.body ? document.body.innerText.slice(0, 6000) : ''"))
                if js:
                    info["js"] = d.execute_script(js)
                d.save_screenshot(os.path.join(a.out, "p%02d.png" % i))
            except Exception as e:                     # keep going: one bad page is data too
                info["error"] = repr(e)
            with io.open(os.path.join(a.out, "p%02d.json" % i), "w", encoding="utf-8") as f:
                f.write(json.dumps(info, ensure_ascii=False, indent=1))
            print(i, info.get("url"), info.get("title") or info.get("error"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
