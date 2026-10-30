# -*- coding: utf-8 -*-
"""The window's 期间报表 tab without tkinter: the choices it offers, the grid
it previews, the command it runs, and what an exit code means.

app.py only lays out widgets and calls these (tests/test_period_report.py pins
them; tests/test_app_smoke.py builds the tab with the fake tkinter).
"""
import os

from core import layout
from core import markets
from core import results as R
from core.period_report import output as G
from core.period_report import model as MD

KINDS = (("quarterly", "季度"), ("annual", "全年"), ("custom", "自定义"))
BASES = (("bill", "账单基准（会计入账，默认）"), ("order", "订单基准（经营分析参考）"))

EXIT_TEXT = {0: ("✓ 期间报表都生成了，每个店铺每个月都有完整版本。", "good"),
             1: ("⚠ 期间报表生成了，但有店铺月份没有完整版本，没计入合计（见报表的①和⑩）。", "warn"),
             2: ("✗ 没生成：所选范围在这个期间没有月度结果，或者出错了（见下面的输出）。", "bad")}


def exit_text(code):
    return EXIT_TEXT.get(code, ("✗ 没跑完（退出码 %s），看下面的输出。" % code, "bad"))


def quarters(latest_month, n=8):
    """The quarters to offer, newest first: the quarter of the latest full month
    first (it may still be open), then the n-1 before it."""
    y, q = int(latest_month[:4]), (int(latest_month[5:7]) - 1) // 3 + 1
    out = []
    for _ in range(n):
        out.append("%d-Q%d" % (y, q))
        y, q = (y - 1, 4) if q == 1 else (y, q - 1)
    return out


def years(latest_month, n=4):
    y = int(latest_month[:4])
    return [str(y - i) for i in range(n)]


def period_of(kind, quarter=None, year=None, first=None, last=None):
    """The tab's choice -> model.Period (ValueError with a Chinese message)."""
    if kind == "quarterly":
        return MD.quarter(quarter)
    if kind == "annual":
        return MD.year(year)
    if not first or not last:
        raise ValueError("自定义期间要选起止月份")
    return MD.custom(first, last)


def period_argv(period):
    if period.kind == "quarterly":
        return ["--quarter", period.label]
    if period.kind == "annual":
        return ["--year", period.label]
    return ["--from", period.months[0], "--to", period.months[-1]]


def command(py, root, period, region, stores=None, basis=None, out_dir=None):
    """run_period.py's command line for the tab's choice."""
    cmd = [py, "-u", os.path.join(root, "run_period.py")] + period_argv(period) + ["--region", region]
    if stores:
        cmd += ["--stores"] + list(stores)
    if basis and basis != "bill":
        cmd += ["--basis", basis]
    if out_dir:
        cmd += ["--out-dir", out_dir]
    return cmd


def grid(period, region, stores=None, basis=None, config_path=None):
    """What the tab previews (reads results only, writes nothing):

    [{"market", "title", "basis", "rows": [(store, [(month, ok, text, reason)])],
      "missing": n, "error": message or None}] — one entry per market of the
    region that has any result in the period. A month before the store's first
    data month shows '○ 未上线' with ok True (it is not a gap) and does not
    count in "missing"."""
    out = []
    for m in G.markets_of(region):
        pf, _ = markets.split(m)
        b = MD.basis_of(pf, basis)
        title = markets.label(m) + ("（%s）" % MD.basis_cn(b) if MD.basis_cn(b) else "")
        try:
            D = MD.gather(m, period, b, stores, config_path)
        except Exception as e:                           # noqa: BLE001
            out.append({"market": m, "title": title, "basis": b, "rows": [], "missing": 0,
                        "error": "%s: %s" % (e.__class__.__name__, e)})
            continue
        if not D.stores or not any(c.final or c.newest for c in D.cells.values()):
            continue
        rows = []
        for s in D.stores:
            cells = []
            for month in D.months:
                c = D.cell(s, month)
                if c.final:
                    cells.append((month, True, "✓ " + R.GRADE_CN[c.final["grade"]["level"]], ""))
                elif c.not_started:
                    cells.append((month, True, "○ %s（从 %s 起有数据）" % (MD.NOT_STARTED, c.first_month), ""))
                else:
                    cells.append((month, False, "✗ 无完整版本", c.reason()))
            rows.append((s, cells))
        out.append({"market": m, "title": title, "basis": b, "rows": rows,
                    "missing": len(D.missing()), "error": None})
    return out


def folder(period, region, out_dir=None):
    """The folder 打开期间报表目录 opens: the market's folder when the scope is one
    market, the period's folder otherwise — else the nearest of the period's
    kind folder and the period reports' root that exists; None when none does."""
    ms = G.markets_of(region)
    d = layout.period_dir(period.kind, period.label, ms[0], override=out_dir)
    period_folder = os.path.dirname(d)
    kind_folder = os.path.dirname(period_folder)
    for x in ([d] if len(ms) == 1 else []) + [period_folder, kind_folder, os.path.dirname(kind_folder)]:
        if os.path.isdir(x):
            return x
    return None


def child_env():
    """The environment run_period.py runs in (UTF-8 output, like the preview)."""
    return dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")


def creation_flags():
    """No console window flashing up from the windowed (pythonw) app."""
    if os.name == "nt":
        import subprocess
        return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)}
    return {}

