# -*- coding: utf-8 -*-
"""
Write a human-readable record of one run into its own download folder.

Every run drops `REPORT.md` next to the files it downloaded, so the folder
explains itself months later: what was pulled, from which route, what was
skipped and why, and what is still owed.

Kept apart from run_downloads.py because it is presentation, not pipeline:
nothing here touches the browser, and a failure to write the report must never
cost the downloads it is describing.
"""
import datetime
import io
import os

REPORT_NAME = "REPORT.md"

# "未下载 —— 原因"这一段前后的记号。补收之后结论会变（欠着的取回了，或者
# 终于分得清原因了），批次收尾要把这一段换掉；有记号就能只换这一段，不必
# 重新生成整份报告 —— --collect-only 时手里也没有当初那一趟的完整结果。
MISSING_BEGIN = "<!-- missing:begin -->"
MISSING_END = "<!-- missing:end -->"

# MercadoPago names its downloads after the report, so the group column can be
# filled from the filename. Longest prefix first: "settlement_v2-" would
# otherwise be missed by a shorter match, and account_statement_generic- shares
# no prefix with the rest only by luck.
MP_PREFIXES = [
    ("account_statement_generic-", "Estados de saldos"),
    ("after_collection-", "Poscobro"),
    ("reserve-release-", "Liberaciones"),
    ("settlement_v2-", "Todas las transacciones"),
    ("collection-", "Cobros"),
    ("withdraw-", "Retiros"),
]


def _mp_group(filename):
    low = filename.lower()
    for prefix, label in MP_PREFIXES:
        if low.startswith(prefix):
            return label
    return ""


def _size(path):
    try:
        n = os.path.getsize(path)
    except OSError:
        return "missing"
    if n < 1024:
        return "%d B" % n
    if n < 1024 * 1024:
        return "%.1f KB" % (n / 1024.0)
    return "%.1f MB" % (n / (1024.0 * 1024.0))


def _rows(result, out_dir):
    """(route, group, filename, size) for everything this run downloaded."""
    rows = []
    for period, files in sorted(result.get("billing", {}).items()):
        for f in files:
            f = os.path.basename(f)
            rows.append(("Facturación", period, f, _size(os.path.join(out_dir, f))))
    for f in result.get("sales", []):
        f = os.path.basename(f)
        rows.append(("Ventas", "Excel de ventas", f, _size(os.path.join(out_dir, f))))
    for key, files in sorted(result.get("stock", {}).items()):
        for f in files:
            f = os.path.basename(f)
            rows.append(("Stock", key, f, _size(os.path.join(out_dir, f))))
    for f in result.get("mp", []):
        f = os.path.basename(f)
        rows.append(("MercadoPago", _mp_group(f), f,
                     _size(os.path.join(out_dir, f))))
    return rows


def _cell(text):
    """表格单元格：竖线和换行会把 markdown 表格撑破。"""
    return str(text if text is not None else "").replace("|", "/").replace(
        "\n", " ").strip()


def missing_section(missing, cautions=None, flags=None, late_files=None):
    """"未下载 —— 原因"这一段，带前后记号。返回行列表。

    missing    download_diag 的诊断记录列表
    cautions   已下载但要留意的（按列表里已有的行匹配）
    flags      ip_check_failed / range_clamped / locale
    late_files 补收阶段取回的文件名
    """
    L = [MISSING_BEGIN, "## 未下载 —— 原因", ""]
    if missing:
        L.append("| 报表 | 期间 | 原因 | 重试有用? | 下一步 | 证据 |")
        L.append("|---|---|---|---|---|---|")
        for r in missing:
            ev = r.get("evidence") or {}
            ref = ev.get("file") or ev.get("screenshot") or ""
            L.append("| `%s` | %s | **%s** %s | %s | %s | %s |" % (
                _cell(r.get("form")), _cell(r.get("period")) or "-",
                _cell(r.get("cause")), _cell(r.get("explain")),
                "是" if r.get("retry") else "否",
                _cell(r.get("action")),
                "`%s`" % _cell(ref) if ref else "-"))
        L.append("")
        L.append("责任方：" + "；".join(
            "`%s` %s" % (r.get("form"), r.get("owner")) for r in missing))
    else:
        L.append("没有。按参数应下的报表全部到齐。")
    L.append("")

    if late_files:
        L.append("### 补收阶段取回")
        L.append("")
        for f in late_files:
            L.append("- `%s`" % _cell(f))
        L.append("")

    if cautions:
        L.append("### 留意：按列表里已有的行匹配")
        L.append("")
        L.append("点了生成之后列表里没有出现对应的新行，脚本改按已有的那一行")
        L.append("去取。通常那就是当天早些时候申请的同一区间；但若「申请」和")
        L.append("「采用」对不上，取回的可能是一张**更早**的报表。")
        L.append("")
        L.append("| 报表 | 申请 | 采用 | 说明 |")
        L.append("|---|---|---|---|")
        for c in cautions:
            L.append("| `%s` | %s | %s | %s |" % (
                _cell(c.get("form")), _cell(c.get("requested")),
                _cell(c.get("adopted")), _cell(c.get("reason"))))
        L.append("")

    flags = flags or {}
    notes = []
    if flags.get("ip_check_failed"):
        notes.append("IP 检测没有通过（`ip_check_failed`）")
    for form, v in sorted((flags.get("range_clamped") or {}).items()):
        notes.append("`%s` 的时间范围被收窄（`range_clamped`）：%s" % (
            form, "，".join("%s=%s" % kv for kv in sorted(v.items()))
            if isinstance(v, dict) else v))
    if "locale" in flags and flags.get("locale") not in (None, "es_MX"):
        notes.append("进店时界面语言是 `%s`（`locale`），已改回 es_MX"
                     % flags.get("locale"))
    if notes:
        L.append("### 运行标记")
        L.append("")
        for x in notes:
            L.append("- %s" % x)
        L.append("")
    L.append(MISSING_END)
    L.append("")
    return L


def render(result, args=None, started=None):
    """The report body, as markdown."""
    out_dir = result["out_dir"]
    rows = _rows(result, out_dir)
    known = set(r[2] for r in rows)
    # Anything in the folder we did not attribute to a route. Usually nothing;
    # if it appears, it is a leftover from an interrupted run and the reader
    # should know rather than quietly counting it as this run's output.
    extra = sorted(f for f in os.listdir(out_dir)
                   if f.lower().endswith((".xlsx", ".xls", ".csv"))
                   and f not in known)

    started = started or datetime.datetime.now()
    secs = result.get("seconds", 0)

    L = []
    L.append("# Descargas — %s" % result["store"])
    L.append("")
    L.append("| | |")
    L.append("|---|---|")
    L.append("| Store | `%s` |" % result["store"])
    L.append("| Run started | %s |" % started.strftime("%Y-%m-%d %H:%M:%S"))
    L.append("| Duration | %d min %02d s |" % (secs // 60, secs % 60))
    L.append("| Files downloaded | **%d** |" % len(rows))
    L.append("| Folder | `%s` |" % out_dir)
    if args is not None:
        # Only the parameters that actually shaped this run. Printing the
        # billing period count on an --only ventas run invites the reader to
        # wonder where the billing files went.
        only = getattr(args, "only", "all")
        L.append("| Routes | %s |" % only)
        if only in ("billing", "all"):
            L.append("| Billing periods | %s |" % getattr(args, "months", "-"))
        if only in ("ventas", "all"):
            L.append("| Ventas range | últimos %s meses |"
                     % getattr(args, "period_months", "-"))
        if only in ("mercadopago", "all"):
            L.append("| MercadoPago window | %s days |"
                     % getattr(args, "mp_days", "-"))
    L.append("")

    if rows:
        L.append("## Files")
        L.append("")
        L.append("| Route | Group | File | Size |")
        L.append("|---|---|---|---|")
        for route, group, f, size in rows:
            L.append("| %s | %s | `%s` | %s |" % (route, group, f, size))
        L.append("")
    else:
        L.append("## Files")
        L.append("")
        L.append("Nothing was downloaded. See the sections below for why.")
        L.append("")

    # 只有做过诊断的结果才有这个键（download_history.py 的结果没有）
    if "missing" in result:
        L.extend(missing_section(result.get("missing"),
                                 cautions=result.get("cautions"),
                                 flags=result.get("flags"),
                                 late_files=result.get("late_files")))

    empty = result.get("empty") or []
    if empty:
        L.append("## No data — not an error")
        L.append("")
        L.append("MercadoPago had no movements in the requested range, so it")
        L.append("would not build these. Re-requesting the same range will not")
        L.append("help; a different period might.")
        L.append("")
        for x in empty:
            L.append("- %s" % x)
        L.append("")

    pending = result.get("pending") or []
    if pending:
        L.append("## Still owed")
        L.append("")
        L.append("Requested but not collected in time. The request stands on the")
        L.append("provider's side — **a later run picks these up**; nothing was")
        L.append("lost. Entries marked `datos en proceso` need a few hours.")
        L.append("")
        for x in pending:
            L.append("- %s" % x)
        L.append("")

    errors = result.get("errors") or []
    if errors:
        L.append("## Errors")
        L.append("")
        L.append("These need someone to look at them.")
        L.append("")
        for x in errors:
            L.append("- %s" % x)
        L.append("")

    if extra:
        L.append("## Unattributed files in this folder")
        L.append("")
        L.append("Present on disk but not produced by this run — most likely a")
        L.append("leftover from an interrupted run.")
        L.append("")
        for f in extra:
            L.append("- `%s` (%s)" % (f, _size(os.path.join(out_dir, f))))
        L.append("")

    L.append("---")
    L.append("")
    L.append("Generated by `run_downloads.py`. Form-by-form detail: `platforms/mercadolibre/FORMS.md`.")
    L.append("")
    return "\n".join(L)


def write(result, args=None, started=None):
    """Write REPORT.md into the run folder. Returns its path, or None.

    Never raises: a run that downloaded 18 files must not be reported as failed
    because a summary file could not be written.
    """
    try:
        path = os.path.join(result["out_dir"], REPORT_NAME)
        # utf-8: the content is Spanish, and Windows' default codepage cannot
        # encode "Facturación" (GUIDE.md 6.4).
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(render(result, args=args, started=started))
        return path
    except Exception as e:
        print("  [warn] could not write %s: %s" % (REPORT_NAME, e))
        return None


def update_missing(out_dir, missing, cautions=None, flags=None, late_files=None):
    """只换掉已有 REPORT.md 里"未下载 —— 原因"那一段。返回路径，或 None。

    给批次收尾用：补收之后结论变了，而那个目录的报告是早先写的。没有记号的
    老报告就把这一段插在页脚之前。和 write() 一样绝不抛异常。
    """
    try:
        path = os.path.join(out_dir, REPORT_NAME)
        if not os.path.isfile(path):
            return None
        with io.open(path, encoding="utf-8") as f:
            text = f.read()
        block = "\n".join(missing_section(missing, cautions=cautions,
                                          flags=flags, late_files=late_files))
        a, b = text.find(MISSING_BEGIN), text.find(MISSING_END)
        if 0 <= a < b:
            text = text[:a] + block.rstrip("\n") + text[b + len(MISSING_END):]
        else:
            foot = text.rfind("\n---\n")
            if foot < 0:
                text = text.rstrip("\n") + "\n\n" + block
            else:
                text = text[:foot + 1] + block + "\n" + text[foot + 1:]
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path
    except Exception as e:
        print("  [warn] could not update %s: %s" % (REPORT_NAME, e))
        return None
