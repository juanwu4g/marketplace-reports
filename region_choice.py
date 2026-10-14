# -*- coding: utf-8 -*-
"""两个窗口（启动 / 定时设置）里"选哪个国家"的那部分逻辑。

放在这里而不是 gui.py / schedule_gui.py 里，是因为那两个文件一 import 就要
tkinter，而跑测试的机器上没有 tkinter。窗口里只剩摆控件、调这里的函数；
规则（店铺清单、任务名、任务参数、两个任务会不会撞车、补跑传哪个区域）都
在这里，有测试钉着。

和 store_config 一样不依赖 selenium / tkinter。
"""
import re

import runlog
import store_config

# 窗口里能选的国家，顺序就是下拉框里的顺序。墨西哥排第一、是默认值。
REGIONS = ("MX", "CL")
DEFAULT = store_config.DEFAULT_REGION

# 计划任务名。墨西哥那个名字生产机器上已经登记过，**不能改**：改了等于
# 新建一个任务，旧的那个还在，每月跑两次。
TASK_NAMES = {"MX": "MX月度报表", "CL": "CL月度报表"}

# 计划任务的超时（schedule_gui.TIME_LIMIT 是 6 小时）。两个国家的任务开始
# 时间离得比这个近，后起的那个就可能撞上前一个的锁、直接被跳过（退出码 5）。
OVERLAP_HOURS = 6


def label(region):
    """下拉框里的字：墨西哥 / 智利。"""
    return store_config.region_name(region)


def labels():
    return [label(r) for r in REGIONS]


def from_label(text):
    """下拉框的字 -> 区域代码。认不出来就当墨西哥。"""
    for r in REGIONS:
        if text in (r, label(r)):
            return r
    return DEFAULT


def remembered(settings):
    """gui_settings.json 里记的区域。没记、记了个不认识的，都是墨西哥。"""
    r = str((settings or {}).get("region") or "").upper()
    return r if r in REGIONS else DEFAULT


def stores(region):
    """窗口里"全部店铺"的清单。读不到 config 返回 [] —— 窗口照常能开。"""
    try:
        return list(store_config.stores_for_region(region))
    except Exception:
        return []


def window_title(region):
    """启动窗口的标题。墨西哥的和以前一字不差。"""
    return "meli-mx-pipeline —— %s店铺报表" % label(region)


def preview_cmd(py, script, month, stores_, region):
    """「先看用哪批数据」的命令：run_reports --dry-run，带上区域。"""
    cmd = [py, script, "--month", month, "--dry-run"]
    if region != DEFAULT:
        cmd += ["--region", region]
    if stores_:
        cmd += ["--stores"] + list(stores_)
    return cmd


def action_region(action):
    """一条「需要人工处理」属于哪个国家：按店铺在 config 里的区域算；config
    读不了就退回那一轮记下的区域。"""
    try:
        return store_config.region_of(action["store"])
    except Exception:
        return action.get("region") or DEFAULT


def action_line(action):
    """「需要人工处理」列表里的一行。墨西哥的店和以前一样，别的国家标出来。"""
    r = action_region(action)
    tag = "" if r == DEFAULT else "（%s）" % label(r)
    return "%s%s —— %s" % (action["store"], tag, action["title"])


def rerun_plan(picked):
    """「处理好了：重跑」要起的那一轮。返回 dict：

    region   这些店的国家（一个）；分属几个国家时为 None
    stores   要重跑的店（去重、保持顺序）
    useless  重跑也没用的那几条（"店名（原因）"）
    groups   [(区域, [店])]，给"分属几个国家"时的提示用

    一轮只能跑一个国家（锁是全局的，run_monthly 也会拒收别国的店），所以选中
    的店分属两国时不起，让人分两次点。
    """
    names, useless, groups = [], [], []
    for a in picked:
        if not a.get("rerun"):
            useless.append("%s（%s）" % (a["store"], a["title"]))
        elif a["store"] not in names:
            names.append(a["store"])
            r = action_region(a)
            for reg, ss in groups:
                if reg == r:
                    ss.append(a["store"])
                    break
            else:
                groups.append((r, [a["store"]]))
    return {"region": groups[0][0] if len(groups) == 1 else None,
            "stores": names, "useless": useless, "groups": groups}


# ────────────────────────────────────────────────────────── 计划任务

def task_name(region):
    return TASK_NAMES.get(region) or "%s月度报表" % region


def rehearsal_name(region):
    return task_name(region) + "-试跑"


def task_args(region):
    """计划任务传给 run_monthly.cmd 的参数。月份、店铺、输出目录都不写。"""
    return " ".join(runlog.runner_argv("scheduled", region=region))


def rehearsal_args(region):
    return " ".join(runlog.runner_argv("scheduled", step="report", region=region))


def args_region(args):
    """从已登记任务的参数里读出区域。旧任务没有 --region，按墨西哥算。"""
    m = re.search(r"--region\s+(\w+)", args or "")
    return m.group(1).upper() if m else DEFAULT


def sched_key(key, region):
    """gui_settings.json 里定时设置的键。墨西哥沿用旧键（sched_day …），
    别的国家加后缀（sched_day_CL …），两边各记各的。"""
    return key if region == DEFAULT else "%s_%s" % (key, region)


def _when(st):
    """task_state() 的结果 -> (几号, 分钟数)。读不出来返回 None。"""
    try:
        h, m = (st or {}).get("TIME", "").split(":")
        return int(st.get("DAY")), int(h) * 60 + int(m)
    except (ValueError, TypeError, AttributeError):
        return None


def clash(a, b):
    """两个国家的任务会不会撞车。a、b 是 (几号, "HH:MM") 或 task_state() 的
    结果；有一个没登记就是 None。撞车时返回一句话，否则 None。

    同一天、开始时间相差不到 OVERLAP_HOURS 小时就算：两国共用一个紫鸟客户端，
    run_monthly 的锁是全局的，后起的那一轮看到锁会直接跳过（退出码 5）——
    同一个钟点当然撞，差两三个小时一样撞（一轮要跑 3 小时左右）。
    """
    def norm(x):
        if isinstance(x, dict):
            return _when(x)
        if not x:
            return None
        try:
            day, hm = x
            h, m = str(hm).split(":")
            return int(day), int(h) * 60 + int(m)
        except (ValueError, TypeError):
            return None
    wa, wb = norm(a), norm(b)
    if not wa or not wb or wa[0] != wb[0]:
        return None
    gap = abs(wa[1] - wb[1])
    if gap >= OVERLAP_HOURS * 60:
        return None
    return ("两个国家都定在每月 %d 号、开始时间只差 %d 小时 %02d 分。两国共用一个"
            "紫鸟，同一时间只能跑一轮：后到点的那个会因为前一个还没跑完而被跳过"
            "（退出码 5）。请把其中一个改到别的日子，或者错开 %d 小时以上。"
            % (wa[0], gap // 60, gap % 60, OVERLAP_HOURS))
