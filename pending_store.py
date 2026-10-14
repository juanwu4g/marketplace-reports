# -*- coding: utf-8 -*-
"""
The reports we have asked MercadoPago to build but have not collected yet.

A request is a side effect on someone else's server. Once `Generar` is clicked
the report is being built whether or not this process survives, so the only
thing that can lose it is us forgetting we asked. That happened on 2026-09-07:
TANKE_EE timed out with two reports still generating and nothing recorded them,
so the next run re-requested from scratch.

This is the record. It is written the moment a store's requests go in, and an
entry is removed only when its file is on disk. Anything left over is collected
by a later run.

Why this can be a flat file rather than anything cleverer: an entry is not a
handle into a browser session. It is the text of a row - report type, url and
the exact period string - which is how `mercadopago.probe()` finds it anyway.
That is why a report survives the store being closed and reopened, or the
machine being rebooted: nothing about the identity is in-memory.
"""
import io
import console  # noqa: F401  中文店名输出保护
import json
import os
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(ROOT, "logs")
PATH = os.path.join(LOG_DIR, "pending.json")
EXPIRED_NAME = "pending_expired.json"

# Reports do not stay downloadable forever, and an entry nobody has collected
# after this long is far more likely to be a stale record than a live report.
MAX_AGE_DAYS = 7


def _key(e):
    """What makes two entries the same outstanding report."""
    return (e.get("store"), e.get("report"), e.get("range"))


def load(path=PATH):
    """Every outstanding entry. Never raises.

    A corrupt file must not stop a batch: the downloads are the point, and the
    worst case of ignoring it is re-requesting a report. It is moved aside
    rather than deleted so it can still be looked at.
    """
    if not os.path.exists(path):
        return []
    try:
        with io.open(path, encoding="utf-8") as f:
            doc = json.load(f)
        entries = doc.get("entries")
        return list(entries) if isinstance(entries, list) else []
    except ValueError as e:
        broken = path + ".broken"
        print("  [警告] %s 不是合法 JSON（%s），已移动到 %s，"
              "按空列表继续" % (path, e, broken))
        try:
            os.replace(path, broken)
        except OSError:
            pass
        return []
    except (IOError, OSError) as e:
        print("  [警告] 无法读取 %s（%s），忽略后继续" % (path, e))
        return []


def save(entries, path=PATH):
    """Replace the file. Never raises - see load()."""
    try:
        if not os.path.isdir(os.path.dirname(path)):
            os.makedirs(os.path.dirname(path))
        with io.open(path, "w", encoding="utf-8") as f:
            # utf-8 + ensure_ascii=False: store names are Chinese on this
            # account (GUIDE.md 6.4).
            f.write(json.dumps({"entries": list(entries)},
                               ensure_ascii=False, indent=2))
        return True
    except (IOError, OSError) as e:
        print("  [警告] 无法写入 %s（%s），待取报表记录将不会保留到"
              "下次运行" % (path, e))
        return False


def add(new_entries, path=PATH):
    """Record newly requested reports, replacing any earlier entry for the same
    (store, report, range).

    Replacing rather than duplicating matters for a same-day re-run: the period
    string is the only identity a row has, so two requests for the same range
    are indistinguishable on the page. Keeping the newest means `requested_at`
    reflects the request actually outstanding.
    """
    entries = [e for e in load(path)
               if _key(e) not in {_key(n) for n in new_entries}]
    entries.extend(new_entries)
    save(entries, path)
    return entries


def remove(done, path=PATH):
    """Drop entries that have been collected (or given up on)."""
    gone = {_key(d) for d in done}
    entries = [e for e in load(path) if _key(e) not in gone]
    save(entries, path)
    return entries


def prune(max_age_days=MAX_AGE_DAYS, path=PATH, now=None):
    """Drop entries too old to still be worth chasing. Returns what it dropped.

    丢掉的条目不是凭空消失：每一条都标上 cause=EXPIRED 追加进同目录的
    pending_expired.json。以前只打印一个数字（"已清理 2 张"），是哪家店的
    哪张报表、最后看到它是什么状态，全都跟着没了 —— 而那恰恰是一张**确定
    没拿到**的报表。
    """
    now = time.time() if now is None else now
    cutoff = now - max_age_days * 86400
    entries, stale = [], []
    for e in load(path):
        (stale if e.get("requested_at", 0) < cutoff else entries).append(e)
    if stale:
        for e in stale:
            e["cause"] = "EXPIRED"
            e["expired_at"] = now
        save(entries, path)
        _archive_expired(stale, path)
    return stale


def expired_path(path=PATH):
    return os.path.join(os.path.dirname(path), EXPIRED_NAME)


def _archive_expired(stale, path=PATH):
    """追加到过期记录里。写不出只警告 —— 和 save() 一个态度。"""
    out = expired_path(path)
    try:
        old = []
        if os.path.exists(out):
            with io.open(out, encoding="utf-8") as f:
                old = (json.load(f) or {}).get("entries") or []
        with io.open(out, "w", encoding="utf-8") as f:
            f.write(json.dumps({"entries": list(old) + list(stale)},
                               ensure_ascii=False, indent=2))
        return True
    except (IOError, OSError, ValueError) as e:
        print("  [警告] 无法写入 %s（%s），过期的待取记录没有留档" % (out, e))
        return False


def describe_entry(e):
    """一条待取记录的一行说明：哪家店、哪张报表、哪个区间。"""
    form = ("ventas" if e.get("kind") == "ventas"
            else "mercadopago/%s" % e.get("report", "?"))
    return "%s %s%s" % (e.get("store"), form,
                        "（%s）" % e["range"] if e.get("range") else "")


def for_store(name, path=PATH):
    return [e for e in load(path) if e.get("store") == name]


def describe(entries=None, path=PATH):
    entries = load(path) if entries is None else entries
    if not entries:
        return "无待取报表"
    by_store = {}
    for e in entries:
        by_store.setdefault(e.get("store"), []).append(e)
    return "; ".join("%s: %d" % (s, len(v)) for s, v in sorted(by_store.items()))
