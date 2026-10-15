# -*- coding: utf-8 -*-
"""「ML 报表」统一窗口背后不碰 tkinter 的那一半。

窗口（ml_app.py）一 import 就要 tkinter，跑测试的机器上没有；所以凡是能写成
"给参数、得结果"的东西都放在这里：界面偏好、用哪个解释器、要跑多久、确认框
里写什么、按钮该不该亮、起一轮的命令行、「需要人工处理」那一页的文字。

国家和计划任务的规则在 region_choice.py，运行记录和日志在 runlog.py。
"""
import datetime
import io
import json
import os
import sys

from core import region_choice
from core import runlog
from core import site_time

from core.paths import ROOT
DEFAULT_REPORT_DIR = os.path.join(ROOT, "downloads", "data", "reports", "financial")
SETTINGS = os.path.join(ROOT, "gui_settings.json")
RUNNER = os.path.join(ROOT, "run_monthly.py")

# 窗口打开时停在哪一页。ML报表.cmd = 立即运行；--tab 可以直接打开别的页。
TABS = ("run", "human", "schedule")
TAB_ALIASES = {"run": "run", "launch": "run", "启动": "run",
               "human": "human", "人工": "human",
               "schedule": "schedule", "sched": "schedule", "定时": "schedule"}

# 每家店下载大约要多久（分钟），给确认框估时用。实测 8–15 分钟，取偏长的。
MIN_PER_STORE = 15


# ────────────────────────────────────────────────────────── 偏好

def load_settings(path=None):
    """界面自己的偏好（输出目录、选的国家、定时的几号几点）。

    刻意不写进 config.json —— 那个文件装凭据和店铺白名单，是照着
    config.example.json 填出来的；把界面偏好混进去，新人对着范本填的时候
    会多出不认识的键。这份坏了删掉即可，不影响任何业务。
    """
    try:
        with io.open(path or SETTINGS, encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def save_settings(d, path=None):
    try:
        with io.open(path or SETTINGS, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass                      # 存不下偏好不该妨碍干活


def update_settings(path=None, **kw):
    save_settings(dict(load_settings(path), **kw), path)


def interpreter():
    """永远优先用项目自己的 .venv。

    不是洁癖：requirements.txt 锁的是 pandas 3，而机器上常见的 conda 环境
    是 pandas 2，两者对 NaN 转字符串的处理不同，报表会崩在 ⑨ 库存页。
    """
    p = os.path.join(ROOT, ".venv", "Scripts",
                     "python.exe" if os.name == "nt" else "python")
    return p if os.path.exists(p) else sys.executable


def pick_tab(argv):
    """命令行里的 --tab xxx（或直接一个词）-> 打开哪一页。认不出来就是立即运行。"""
    argv = list(argv or [])
    for i, a in enumerate(argv):
        if a == "--tab" and i + 1 < len(argv):
            return TAB_ALIASES.get(argv[i + 1].lower(), "run")
        if a.startswith("--tab="):
            return TAB_ALIASES.get(a.split("=", 1)[1].lower(), "run")
    for a in argv:
        if a.lower() in TAB_ALIASES:
            return TAB_ALIASES[a.lower()]
    return "run"


# ────────────────────────────────────────────────────────── 月份、时长

def recent_months(n=15, today=None):
    """最近 n 个完整月份，最近的在前。和 run_reports.prev_month 同一个"今天"。"""
    d = (today or site_time.site_today()).replace(day=1)
    out = []
    for _ in range(n):
        d -= datetime.timedelta(days=1)
        out.append("%04d-%02d" % (d.year, d.month))
        d = d.replace(day=1)
    return out


def run_hours(n):
    """n 家店大概要跑多久，给确认框用。每家约 15 分钟。"""
    m = n * MIN_PER_STORE
    return "%d 分钟" % m if m < 60 else "%.1f 小时" % (m / 60.0)


def estimate(step, n_stores, n_regions=1):
    """这一步大概要多久，一句话。"""
    if step in ("both", "download"):
        return "约 %s" % run_hours(n_stores)
    if step == "report":
        return "约 %d 分钟" % max(1, n_regions)
    return "几分钟到半小时（看欠了多少张）"


# ────────────────────────────────────────────────────────── 起一轮

STEP_TITLES = {"both": "下载并出报表", "download": "只下载数据",
               "report": "只出报表", "collect": "补取欠的报表"}


def runner_cmd(step, month, choice, stores=None, out_dir=None, reports_for=None,
               py=None):
    """「立即运行」起的那条命令：run_monthly.py --entrance manual …。

    choice 是区域代码或 ALL。和计划任务的区别只有 --entrance（见 runlog.runner_argv）。
    """
    return [py or interpreter(), "-u", RUNNER] + runlog.runner_argv(
        "manual", step, month, stores, out_dir, reports_for=reports_for,
        region=choice)


def check_stores(choice, all_checked, picked):
    """「店铺」那一栏能不能拿去跑。返回 (店铺子集或 None, 问题或 None)。

    None = 这个范围里的全部店铺。全部国家时没有子集可选（run_monthly 也不收）。
    取消了"全部店铺"又一家都没选：以前会悄悄按全部跑、同时又不出汇总表，
    两头不靠 —— 现在直接说。
    """
    if choice == region_choice.ALL or all_checked:
        return None, None
    if not picked:
        return None, ("取消了“全部店铺”，但列表里一家都没选。\n"
                      "按住 Ctrl 点选要跑的店，或者重新勾上“全部店铺”。")
    return list(picked), None


def confirm_text(step, choice, month, subset=None, latest=None, groups=None):
    """起一轮之前的确认框：(标题, 正文)。

    写清楚范围（几个国家、几家店）、大概多久、期间别碰紫鸟、窗口可以关。
    groups 不给就按 choice 去读 config（region_choice.lineup）。
    """
    groups = groups if groups is not None else region_choice.lineup(choice)
    if subset:
        region = groups[0][0] if groups else choice
        groups = [(region, list(subset))]
    n = sum(len(ss) for _, ss in groups)
    scope = region_choice.scope_text(groups)
    if subset:
        scope = "%s的 %d 家：%s" % (region_choice.label(groups[0][0]), n,
                                  "、".join(subset))
    title = STEP_TITLES.get(step, step)
    lines = ["范围：%s" % scope, "月份：%s" % month,
             "预计：%s" % estimate(step, n, len(groups))]
    if choice == region_choice.ALL and len(groups) > 1:
        lines.append("顺序：%s，一个国家跑完再跑下一个"
                     % " → ".join(region_choice.label(r) for r, _ in groups))
    if subset and step in ("both", "report"):
        lines.append("只选了部分店铺，不出全店汇总表（避免覆盖真正的那份）。")
    if latest and month != latest:
        lines.append("")
        lines.append("注意：%s 不是最近一个月（%s）。脚本会去那个月的账单明细页"
                     "取数，下载目录带 _m%s 后缀。" % (month, latest,
                                                 month.replace("-", "")))
    if step in ("both", "download", "collect"):
        lines += ["", "期间这台电脑会反复开关紫鸟浏览器 —— 请不要手动去点紫鸟，",
                  "那会打断下载。"]
    lines += ["", "这个窗口可以关掉，不影响后台继续跑。", "", "开始？"]
    return title, "\n".join(lines)


# ────────────────────────────────────────────────────────── 按钮

def button_states(busy, has_stores, has_human, picked_human=False):
    """哪些按钮该亮。返回 {名字: True/False}。

    busy：有一轮在跑（不管谁起的）—— 起新一轮的按钮全灭，中止亮。下载一启动就
    会重启紫鸟客户端，再起一轮会把正在跑的那轮连浏览器一起杀掉。
    has_stores：所选范围里有没有店；没有就没什么可跑的。
    """
    can = not busy and has_stores
    return {"run": can, "download": can, "report": can, "collect": can,
            "preview": has_stores, "stop": busy,
            "open_stores": not busy and has_human and picked_human,
            "rerun": not busy and has_human and picked_human}


# ────────────────────────────────────────────────────────── 需要人工处理

def human_title(human):
    """那一页标签上的字：「需要人工处理 (2)」。"""
    n = len(set(a["store"] for a in human.get("actions") or []))
    return "需要人工处理（%d）" % n if n else "需要人工处理"


def human_headline(human):
    """那一页最上面一行：(文字, 颜色类别)。"""
    acts = human.get("actions") or []
    if not acts:
        return "最近一轮下载里，没有需要人去店铺里处理的事。", "good"
    stores = sorted(set(a["store"] for a in acts))
    return ("%d 家店铺需要人去处理（%s）。脚本自己解决不了，处理完点「重跑」。"
            % (len(stores), human.get("month") or ""), "bad")


def human_detail(picked):
    """选中的几条：每条怎么处理。"""
    if not picked:
        return "在上面选中要处理的店铺（按住 Ctrl 可多选）。"
    return "\n".join("%s：%s%s" % (a["store"], a["how"],
                                  "" if a.get("rerun") else "（处理之前重跑没用）")
                     for a in picked)


def rerun_confirm(plan, month):
    """「处理好了：重跑」的确认文字，或者不能重跑时的提示。

    返回 (能不能跑, 标题, 正文)。规则在 region_choice.rerun_plan。"""
    stores, useless = plan["stores"], plan["useless"]
    if stores and not plan["region"]:
        return False, "重跑", (
            "所选的店铺分属不同的国家，一轮只能补跑一个国家：\n\n%s\n\n"
            "请只选同一个国家的店铺，分两次点。" % "\n".join(
                "%s：%s" % (region_choice.label(r), "、".join(ss))
                for r, ss in plan["groups"]))
    if not stores:
        return False, "重跑", ("所选的店铺都是重跑解决不了的问题：\n%s"
                              % "\n".join(useless) if useless
                              else "先在列表里选中要重跑的店铺。")
    note = ("\n\n这几项重跑也解决不了，先跳过：\n%s" % "\n".join(useless)
            if useless else "")
    return True, "重跑并重出报表", (
        "会重新下载 %s 的 %s 数据（每家约 %d 分钟，共约 %s），\n"
        "下完给%s全部店铺重出报表，并重出全店汇总表。%s\n\n开始？"
        % ("、".join(stores), month, MIN_PER_STORE, run_hours(len(stores)),
           region_choice.label(plan["region"]), note))


# ────────────────────────────────────────────────────────── 状态栏

def status_line(st, records=(), now=None):
    """窗口最上面那条横幅：(文字, 颜色类别)。

    在跑：国家、第几个国家、月份、步骤、店铺、已运行多久（runlog.live_status）。
    卡住：说清楚、指向「中止」。结束：最近一轮的结果。
    """
    if not st:
        return "还没有运行记录。选好范围和月份，点「下载并出报表」。", None
    who = runlog.ENTRANCE_ZH.get(st.get("entrance"), "?")
    if st.get("state") == "running" and st.get("stalled"):
        return "▲ 正在跑的那一轮（%s）%s" % (who, st.get("text", "")), "bad"
    if st.get("state") == "running":
        live = runlog.live_status(st, records, now) or ""
        return "● 正在运行（%s）：%s" % (who, live), "run"
    mark = {"good": "✓", "warn": "⚠", "bad": "✗"}.get(st.get("kind"), "■")
    return ("%s 最近一轮（%s，%s%s）：%s"
            % (mark, who, runlog.region_tag(st), st.get("month") or "",
               st.get("text", "")), st.get("kind"))


# ────────────────────────────────────────────────────────── 双击之前的自检

def self_check(out=None):
    """ML报表.cmd 打开窗口之前先跑一遍（python -m core.ml_app --check）。

    pythonw 出错时什么都不显示，用户只会觉得"双击没反应"。这里用带控制台的
    python 把会让窗口打不开的几种情况先查一遍，查到就说一句人话（不甩整屏
    traceback），返回 1；cmd 看到非 0 就把这段话摆出来并停住。全过返回 0。
    """
    out = out or sys.stdout
    problems = []

    def check(what, fn, fix):
        try:
            fn()
        except Exception as e:
            problems.append("%s：%s: %s\n    怎么办：%s"
                            % (what, e.__class__.__name__, str(e)[:200], fix))

    def tk_ok():
        import tkinter
        r = tkinter.Tk()
        r.withdraw()
        r.destroy()

    check("界面库 tkinter", tk_ok,
          "重新安装 Python（勾选 tcl/tk and IDLE），再重建 .venv。")
    if not problems:                    # 没有 tkinter 时这一项必然也失败，不重复说
        check("窗口程序文件", lambda: __import__("core.ml_app"),
              "程序文件不全或被改坏了，从仓库重新拉一份。")
    check("报表依赖（pandas 等）", lambda: __import__("core.run_monthly"),
          "在本目录执行 .venv\\Scripts\\pip install -r requirements.txt")
    if not problems:
        return 0
    for p in problems:
        try:
            out.write("  " + p + "\n\n")
        except Exception:
            pass
    return 1
