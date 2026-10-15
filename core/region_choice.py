# -*- coding: utf-8 -*-
"""「ML 报表」窗口里"选哪个国家"、"计划任务怎么登记"的那部分逻辑。

放在这里而不是 ml_app.py 里，是因为窗口一 import 就要 tkinter，而跑测试的
机器上没有 tkinter。窗口里只剩摆控件、调这里的函数；规则（店铺清单、全部
国家、任务名、任务参数、旧任务怎么处理、两个任务会不会撞车、补跑传哪个区域）
都在这里，有测试钉着。

和 store_config 一样不依赖 selenium / tkinter。
"""
import re

from core import runlog
from core import store_config

# 窗口里能选的国家，顺序就是下拉框里的顺序。墨西哥排第一、是默认值。
REGIONS = tuple(store_config.REGION_NAMES)
DEFAULT = store_config.DEFAULT_REGION

# 「全部国家」：run_monthly --region ALL，一轮里依次跑每个有店铺的国家。
ALL = "ALL"
ALL_LABEL = "全部国家"

# 推荐的那个计划任务：一个任务跑全部国家。按国家分开的任务（MX月度报表 …）
# 两个时间挨得近就撞锁、后到的被跳过；一个任务从根上没有这个问题。
ALL_TASK = "ML月度报表"

# 计划任务名。墨西哥那个名字生产机器上已经登记过，**不能改**：改了等于
# 新建一个任务，旧的那个还在，每月跑两次。
TASK_NAMES = {r: r + "月度报表" for r in REGIONS}

# 计划任务的超时（win_tasks.TIME_LIMIT 是 6 小时）。两个国家的任务开始
# 时间离得比这个近，后起的那个就可能撞上前一个的锁、直接被跳过（退出码 5）。
OVERLAP_HOURS = 6


def label(region):
    """下拉框里的字：墨西哥 / 智利 / 全部国家。"""
    if region == ALL:
        return ALL_LABEL
    return store_config.region_name(region)


def labels():
    return [label(r) for r in REGIONS]


def from_label(text):
    """下拉框的字 -> 区域代码。认不出来就当墨西哥。"""
    if text in (ALL, ALL_LABEL):
        return ALL
    for r in REGIONS:
        if text in (r, label(r)):
            return r
    return DEFAULT


def choice_labels():
    """「立即运行」里国家下拉框的字：全部国家排第一，后面是各个国家。"""
    return [ALL_LABEL] + labels()


def remembered_choice(settings):
    """「立即运行」上次选的范围。没记过就是全部国家 —— 这个窗口就是给"所有
    ML 店铺一起出报表"用的。"""
    r = str((settings or {}).get("region") or "").upper()
    return r if r in REGIONS or r == ALL else ALL


def lineup(choice):
    """选了 choice 之后要跑哪些国家、各几家店：[(区域, [店名])]。

    全部国家：和 run_monthly.region_lineup 同一个顺序（墨西哥第一），一家店都
    没有的国家不出现。读不到 config 返回 [] —— 窗口照常能开。"""
    if choice != ALL:
        return [(choice, stores(choice))]
    try:
        order = list(store_config.regions())
    except Exception:
        return []
    order += [r for r in REGIONS if r not in order]
    out = []
    for r in order:
        names = stores(r)
        if names:
            out.append((r, names))
    return out


def scope_text(groups):
    """"墨西哥 12 家、智利 3 家，共 15 家" —— 确认框和店铺区下面那行字用。"""
    n = sum(len(ss) for _, ss in groups)
    if not groups:
        return "没有店铺（config.json 读不到，或者一家都没配）"
    if len(groups) == 1:
        return "%s %d 家" % (label(groups[0][0]), n)
    return "%s，共 %d 家" % ("、".join("%s %d 家" % (label(r), len(ss))
                                      for r, ss in groups), n)


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
    """旧「启动」窗口的标题（现在的统一窗口叫 MercadoLibre 月度报表）。"""
    return "meli-mx-pipeline —— %s店铺报表" % label(region)


def preview_cmd(py, script, month, stores_, region):
    """「先看用哪批数据」的命令：run_reports --dry-run，带上区域。"""
    # script：脚本路径，或 ["-m", "包.模块"]
    head = list(script) if isinstance(script, (list, tuple)) else [script]
    cmd = [py] + head + ["--month", month, "--dry-run"]
    if region != DEFAULT:
        cmd += ["--region", region]
    if stores_:
        cmd += ["--stores"] + list(stores_)
    return cmd


def preview_cmds(py, script, month, stores_, choice):
    """同上，但「全部国家」时每个国家一条（run_reports 一次只看一个国家）。
    返回 [(区域, 命令)]。"""
    if choice != ALL:
        return [(choice, preview_cmd(py, script, month, stores_, choice))]
    return [(r, preview_cmd(py, script, month, None, r)) for r, _ in lineup(ALL)]


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
    if region == ALL:
        return ALL_TASK
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


# ────────────────────────────────────────────────────────── 计划任务总览

# 一个国家一轮最多跑多久：下载自己 5 小时掐断（run_monthly.MAX_HOURS），再加
# 出报表和写告警。Windows 那个"超过 N 小时就停"按国家数放宽，否则跑全部国家的
# 任务会在第二个国家半路上被整个杀掉。
HOURS_PER_REGION = 6


def time_limit_hours(n_regions):
    """全部国家那个任务的超时（小时）。至少 6 小时，和以前单个国家的一样。"""
    return HOURS_PER_REGION * max(1, int(n_regions or 1))


def time_limit_xml(hours):
    return "PT%dH" % int(hours)


def registered_limit(st):
    """已登记任务的超时（小时），读不出来返回 None。task_state() 里的 XML
    是压成一行的原文。"""
    m = re.search(r"<ExecutionTimeLimit>PT(\d+)H</ExecutionTimeLimit>",
                  (st or {}).get("XML", ""))
    return int(m.group(1)) if m else None


# 0–7 是 run_monthly.py 自己的退出码（见它的文件头），267xxx 是 Windows 任务
# 计划程序的状态码。都翻成人话，不要让人去查代码表。
_RESULTS = {"0": "上次跑完，一切正常",
            "1": "上次跑完了，但有几项要人看一眼（看右边的「运行日志」）",
            "2": "上次没跑成：浏览器没能启动 / 没有一家店下到数据 / 有报表没出来",
            "3": "上次没跑成：店铺名单对不上，要改配置",
            "4": "上次没跑成：配置文件读不了",
            "5": "上次没跑：另一轮（可能是手动点的）还没结束，这次被跳过了",
            "6": "上次被人点了「中止」",
            "7": "上次下载超过 5 小时被掐断，报表是按已下到的数据出的",
            "267009": "正在跑",
            "267011": "还没跑过",
            "267014": "上次被 Windows 中途停掉了（超过时限，或有人在任务计划程序里停的）"}


def task_verdict(code):
    code = str(code if code is not None else "")
    return _RESULTS.get(code, "上次没跑成（代码 %s），看右边的「运行日志」" % code)


def task_when(st):
    """"每月 15 号 02:00"；没登记返回 "未登记"。"""
    if not st:
        return "未登记"
    return "每月 %s 号 %s" % (st.get("DAY", "?"), st.get("TIME", "?"))


def schedule_overview(states, n_regions=1):
    """「每月自动运行」那一页要显示什么。纯函数。

    states：{任务名: task_state() 的结果或 None}，至少有 ALL_TASK 和每个国家的
    任务名（没登记的是 None）。n_regions：现在有店铺的国家数。

    返回 dict：
      main     全部国家那个任务的现状（None = 没登记）
      legacy   已登记的按国家的旧任务：[(区域, 任务名, 现状)]
      notes    要提醒的话：[(级别 good/warn/bad, 一句话)]
      headline 状态区那一大行
      kind     那一行的颜色类别

    规则：
    * 推荐一个「全部国家」任务（ALL_TASK，--region ALL）。
    * 旧的按国家任务**只显示、不自动删**：生产机器上早就登记了 MX月度报表，
      悄悄删掉它，换了人的那个月就没有报表了。由人点「替换」或「删除」。
    * 新旧都在：墨西哥一个月会被跑两次（后到的撞锁被跳过，或者白下一遍）—— 红字。
    * 只有旧任务：照样能跑，提示可以换成一个任务。
    * 几个旧任务之间挨得太近：沿用 clash() 那句话。
    * 全部国家那个任务的超时比现在需要的短（后来又加了国家）：提示重新登记。
    """
    main = states.get(ALL_TASK)
    legacy = [(r, task_name(r), states.get(task_name(r))) for r in REGIONS
              if states.get(task_name(r))]
    notes = []
    if main:
        if args_region(main.get("ARGS")) != ALL:
            notes.append(("bad", "「%s」登记的参数（%s）跑的不是全部国家：再点一次"
                                 "「设为每月自动运行」改过来。"
                          % (ALL_TASK, main.get("ARGS") or "无")))
        lim = registered_limit(main)
        need = time_limit_hours(n_regions)
        if lim is not None and lim < need:
            notes.append(("warn", "「%s」的时限是 %d 小时，现在有 %d 个国家要跑，"
                                  "建议 %d 小时：再点一次「设为每月自动运行」更新。"
                          % (ALL_TASK, lim, n_regions, need)))
        for r, name, st in legacy:
            notes.append(("bad", "按国家的旧任务「%s」（%s，%s）也还在：%s的店每月会"
                                 "被跑两次，后到点的那次会撞上锁被跳过。建议点它"
                                 "旁边的「删除」。"
                          % (name, label(r), task_when(st), label(r))))
        headline = "● 已设置：%s 自动跑全部国家" % task_when(main)
        kind = "bad" if any(k == "bad" for k, _ in notes) else "good"
    elif legacy:
        headline = "● 现在用的是按国家登记的旧任务：%s" % "、".join(
            "%s（%s）" % (name, task_when(st)) for _, name, st in legacy)
        kind = "warn"
        notes.append(("warn", "照样能跑。推荐换成一个任务跑全部国家（%s）："
                              "点「设为每月自动运行」，会先问你要不要替换旧任务。"
                      % ALL_TASK))
        for i, (ra, na, sa) in enumerate(legacy):
            for rb, nb, sb in legacy[i + 1:]:
                msg = clash(sa, sb)
                if msg:
                    notes.append(("bad", "「%s」和「%s」撞车：%s" % (na, nb, msg)))
    else:
        headline = "● 还没设置，现在不会自动跑"
        kind = "warn"
        notes.append(("warn", "选好几号几点，点「设为每月自动运行」。"))
    return {"main": main, "legacy": legacy, "notes": notes,
            "headline": headline, "kind": kind}


def replace_question(legacy):
    """登记全部国家任务前，要不要顺手删掉旧任务 —— 确认框里的话。没有旧任务返回 None。"""
    if not legacy:
        return None
    return ("这台电脑上还登记着按国家的旧任务：\n\n%s\n\n"
            "新任务会跑全部国家（含这些国家）。旧任务留着的话，这些国家每月会被跑"
            "两次，后到点的那次会撞上锁被跳过。\n\n"
            "「是」= 登记新任务，并删除上面这些旧任务（推荐）\n"
            "「否」= 只登记新任务，旧任务先留着\n"
            "「取消」= 什么都不做"
            % "\n".join("  · %s（%s，%s）" % (name, label(r), task_when(st))
                         for r, name, st in legacy))
