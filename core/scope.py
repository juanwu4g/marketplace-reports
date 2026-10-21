# -*- coding: utf-8 -*-
"""「月度报表」窗口里的范围：平台 → 国家 → 店铺 这棵树，和勾选之后跑什么。

两页（立即运行、每月自动运行）用的是同一棵树。树完全从 core/markets.py（有哪些
平台、哪些国家、叫什么）和 config.json（每个国家有哪些店）算出来：加平台、加国家
只改那两处，窗口里会自己多出一行，这里和 app.py 都不用动。config 里一家店都没有
的平台 / 国家不出现。

勾选（selection）是 {区域: None 或 [店]}：None = 这个区域的全部店铺（以后
config 里加的店也算进来），列表 = 只跑这几家。区域就是 run_monthly 认的
--region（美客多是国家代码 MX，别的平台是市场代码 FA-CL）。

和 region_choice 一样不依赖 tkinter：规则都在这里，有测试钉着。
"""
import collections

from core import markets
from core import region_choice as RC
from core import runlog

Unit = collections.namedtuple("Unit", "unit platform country name stores")
Platform = collections.namedtuple("Platform", "code name units")

ALL = RC.ALL


def platform_name(code):
    """"美客多 MercadoLibre"、"Falabella"、"亚马逊 Amazon"：中文名，英文名不同就跟在后面。"""
    p = markets.platform(code)
    return p.name_zh if p.name_zh == p.name_en else "%s %s" % (p.name_zh, p.name_en)


def tree(groups=None):
    """[Platform(code, name, [Unit])]。平台按 markets.PLATFORMS 的顺序，每个平台里的
    国家按 markets.COUNTRIES 的顺序（新加的国家排在已知的后面）。

    groups：region_choice.lineup(ALL) 的结果（测试可以直接给）。"""
    groups = RC.lineup(ALL) if groups is None else groups
    countries = list(markets.COUNTRIES)
    by = collections.OrderedDict((code, []) for code in markets.PLATFORMS)
    for unit, stores in groups:
        if not stores:
            continue
        p, c = markets.split(markets.unit_market(unit))
        by.setdefault(p, []).append(Unit(unit, p, c, markets.country_name(c), list(stores)))
    out = []
    for code, units in by.items():
        if units:
            units.sort(key=lambda u: countries.index(u.country) if u.country in countries
                       else len(countries))
            out.append(Platform(code, platform_name(code), units))
    return out


def units(t):
    return [u for p in t for u in p.units]


def unit_label(u):
    """"美客多 · 墨西哥"：平台和国家都写出来（以前"墨西哥"其实是美客多墨西哥）。"""
    return "%s · %s" % (markets.platform(u.platform).name_zh, u.name)


def find(t, unit):
    for u in units(t):
        if u.unit == unit:
            return u
    return None


# ────────────────────────────────────────────────────────── 勾选

def everything(t):
    return collections.OrderedDict((u.unit, None) for u in units(t))


def normalize(sel, t):
    """去掉树里没有的区域和店；一个区域的店全勾了就是 None（全部）；一家都没勾的
    区域去掉。顺序按树。"""
    out = collections.OrderedDict()
    for u in units(t):
        if u.unit not in (sel or {}):
            continue
        picked = sel[u.unit]
        if picked is None:
            out[u.unit] = None
            continue
        keep = [s for s in u.stores if s in picked]
        if keep:
            out[u.unit] = None if len(keep) == len(u.stores) else keep
    return out


def is_everything(sel, t):
    sel = normalize(sel, t)
    return len(sel) == len(units(t)) and all(v is None for v in sel.values())


def region_arg(sel, t):
    """run_monthly --region：全都勾了 = ALL（以后新加的平台/国家也自动包含，就是
    生产机器上已经登记的那个任务）；只勾一个 = 它的代码；几个 = "MX,FA-CL"。
    一个都没勾 = None。"""
    sel = normalize(sel, t)
    if not sel:
        return None
    if is_everything(sel, t):
        return ALL
    return ",".join(sel)


def subset(sel, t):
    """--stores：只勾了部分店的那些区域的店（全勾的区域不列）。没有就是 None。"""
    out = [s for v in normalize(sel, t).values() if v for s in v]
    return out or None


def groups(sel, t):
    """[(Unit, 要跑的店, 是不是只跑部分)]，按树的顺序。"""
    out = []
    for unit, picked in normalize(sel, t).items():
        u = find(t, unit)
        out.append((u, list(u.stores) if picked is None else picked, picked is not None))
    return out


def store_count(sel, t):
    return sum(len(ss) for _, ss, _ in groups(sel, t))


def scope_text(sel, t):
    """"美客多：墨西哥 12 家、智利 1/3 家；Falabella：智利 3 家 —— 共 16 家"。"""
    gs = groups(sel, t)
    if not gs:
        return "一个都没勾"
    parts = collections.OrderedDict()
    for u, ss, part in gs:
        parts.setdefault(u.platform, []).append(
            "%s %d/%d 家" % (u.name, len(ss), len(u.stores)) if part
            else "%s %d 家" % (u.name, len(ss)))
    body = "；".join("%s：%s" % (markets.platform(p).name_zh, "、".join(cs))
                    for p, cs in parts.items())
    return "%s —— 共 %d 家" % (body, store_count(sel, t))


def from_region_arg(arg, t, stores=None):
    """反过来：已登记任务的 --region（和 --stores）-> 勾选。认不出来的区域丢掉。"""
    if not arg or arg == ALL:
        return everything(t)
    sel = collections.OrderedDict((u, None) for u in RC.split_units(arg))
    if stores:
        for u in units(t):
            mine = [s for s in stores if s in u.stores]
            if mine and u.unit in sel:
                sel[u.unit] = mine
    return normalize(sel, t)


def to_settings(sel, t):
    """记进 gui_settings.json 的样子：{"units": {区域: null 或 [店]}}。"""
    return {"units": dict(normalize(sel, t))}


def from_settings(d, t, default=None):
    """gui_settings.json 里记的勾选。没记过、记的东西全都对不上：default（不给就是全部）。"""
    got = normalize((d or {}).get("units") or {}, t) if isinstance(d, dict) else {}
    if got:
        return got
    return default if default is not None else everything(t)


# ────────────────────────────────────────────────────────── 上次运行

def last_results(log_dir=None, limit=60):
    """每个区域最近一次运行的结果：{区域: {"day", "exit", "state", "partial"}}。

    一轮跑好几个区域（ALL 或勾选的几个）时每个区域各记各的。读不到日志就是 {}。"""
    out = {}
    try:
        runs = runlog.list_runs(log_dir, limit=limit)
    except Exception:
        return out
    for run in runs:
        doc = runlog.read_json(runlog.run_paths(run["run_id"], log_dir)["json"])
        if not isinstance(doc, dict):
            continue
        day = run["when"].strftime("%m-%d") if hasattr(run["when"], "strftime") else ""
        if runlog.is_multi(doc.get("region")):
            recs = [(r.get("region"), r.get("exit"), r.get("status"),
                     not r.get("all_stores", True))
                    for r in doc.get("regions") or [] if isinstance(r, dict)]
        else:
            recs = [(doc.get("region") or "MX", doc.get("exit"), doc.get("status"),
                     doc.get("stores") not in ("all", None))]
        for unit, code, state, partial in recs:
            if unit and unit not in out:
                out[unit] = {"day": day, "exit": code, "state": state, "partial": partial}
    return out


def status_text(res):
    """树上"上次运行"那一列：'10-15 ✔ 正常'。没跑过是空串。"""
    if not res:
        return ""
    if res.get("state") == "running":
        return "%s ● 正在跑" % res["day"]
    code = res.get("exit")
    word = {0: "✔ 正常", 1: "⚠ 需关注", 5: "· 被跳过", 6: "· 中止", 7: "⚠ 超时"}.get(
        code, "✗ 出错" if code is not None else "?")
    return "%s %s%s" % (res["day"], word, "（部分店）" if res.get("partial") else "")


def unit_status(results, unit):
    return status_text((results or {}).get(unit))
