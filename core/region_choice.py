# -*- coding: utf-8 -*-
"""「月度报表」窗口里"选哪个国家"、"计划任务怎么登记"的那部分逻辑。

放在这里而不是 app.py 里，是因为窗口一 import 就要 tkinter，而跑测试的
机器上没有 tkinter。窗口里只剩摆控件、调这里的函数；规则（店铺清单、全部
国家、任务名、任务参数、旧任务怎么处理、两个任务会不会撞车、补跑传哪个区域）
都在这里，有测试钉着。

和 store_config 一样不依赖 selenium / tkinter。
"""
import re

from core import markets
from core import runlog
from core import store_config

# 窗口里能选的国家，顺序就是下拉框里的顺序。墨西哥排第一、是默认值。美客多的
# 国家照旧用国家代码；别的平台接在后面，用市场代码（"FA-CL" = Falabella 智利）。
REGIONS = tuple(markets.all_units())
DEFAULT = store_config.DEFAULT_REGION

# 「全部国家」：run_monthly --region ALL，一轮里依次跑每个有店铺的国家。
ALL = "ALL"
ALL_LABEL = "全部国家"

# 推荐的那个计划任务：一个任务跑全部国家。按国家分开的任务（MX月度报表 …）
# 两个时间挨得近就撞锁、后到的被跳过；一个任务从根上没有这个问题。
ALL_TASK = "月度报表"

# 计划任务名。墨西哥那个名字生产机器上已经登记过，**不能改**：改了等于
# 新建一个任务，旧的那个还在，每月跑两次。
TASK_NAMES = {r: r + "月度报表" for r in REGIONS}

# 计划任务的超时（win_tasks.TIME_LIMIT 是 6 小时）。两个国家的任务开始
# 时间离得比这个近，后起的那个就可能撞上前一个的锁、直接被跳过（退出码 5）。
OVERLAP_HOURS = 6


def is_list(region):
    """窗口里勾出来的一串区域："MX,FA-CL"（run_monthly --region 照收）。"""
    return "," in str(region or "")


def split_units(region):
    """"MX,FA-CL" -> ["MX", "FA-CL"]；单个区域 -> [它]。"""
    return [u for u in str(region or "").split(",") if u]


def config_problems():
    """config.json 里这一版程序跑不了的店：[(店名, 原因)]；config 读不了返回 []。"""
    try:
        return store_config.config_problems()
    except Exception:
        return []


def config_error():
    """config.json 整个读不了（JSON 坏了、文件不在）时的原因，正常时 None。窗口里
    店铺树是空的时候要把这句话摆出来，不能空着不说。"""
    try:
        store_config.allowed_stores()
        return None
    except Exception as e:
        return str(e)


def label(region):
    """界面上的字：墨西哥 / 智利 / 全部国家；一串区域用顿号连起来。"""
    if region == ALL:
        return ALL_LABEL
    if is_list(region):
        return "、".join(label(u) for u in split_units(region))
    return store_config.region_name(region)


def visible():
    """下拉框里出现的单元：美客多的国家（和以前一样，都在），加上 config 里有店铺的
    别的平台。没配 Falabella 店的电脑上不会多出一个选了就报错的选项。"""
    out = [r for r in REGIONS if markets.SEP not in r]
    try:
        out += [u for u in store_config.other_units_configured() if u not in out]
    except Exception:
        pass
    return out


def unit_title(unit):
    """"美客多 墨西哥"、"Falabella 智利"：平台和国家都写出来（label 里美客多只写国家，
    那是区域出现之前留下来的写法）。"""
    try:
        return markets.label(markets.unit_market(unit))
    except Exception:
        return label(unit)


def labels():
    return [label(r) for r in visible()]


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
    店铺一起出报表"用的。"""
    r = str((settings or {}).get("region") or "").upper()
    return r if r in REGIONS or r == ALL else ALL


def lineup(choice):
    """选了 choice 之后要跑哪些国家、各几家店：[(区域, [店名])]。

    全部国家：和 run_monthly.region_lineup 同一个顺序（墨西哥第一），一家店都
    没有的国家不出现。读不到 config 返回 [] —— 窗口照常能开。"""
    if is_list(choice):
        return [(u, ss) for u, ss in ((u, stores(u)) for u in split_units(choice)) if ss]
    if choice != ALL:
        return [(choice, stores(choice))]
    try:
        order = list(store_config.regions())
        order += [r for r in store_config.REGION_NAMES if r not in order]
        order += [u for u in store_config.other_units_configured() if u not in order]
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


def platforms_text(groups=None):
    """全部那个任务覆盖哪些平台、哪些国家："美客多（墨西哥、智利、秘鲁）、Falabella（智利）"。

    平台按 markets.PLATFORMS 的顺序，国家按跑的顺序；只说 config 里真有店铺的。
    一家都读不到（config 读不了）时说出代码支持的全部 —— 句子不能是空的。
    界面上写平台和国家的地方都用它，不在窗口里写死平台名。"""
    groups = lineup(ALL) if groups is None else groups
    units = [u for u, _ in groups] or list(REGIONS)
    out = []
    for p in markets.PLATFORMS.values():
        cs = [markets.country_name(markets.split(u)[1]) for u in units
              if markets.unit_platform(u) == p.code]
        if cs:
            out.append("%s（%s）" % (p.name_zh, "、".join(cs)))
    return "、".join(out)


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
    """旧「启动」窗口的标题（现在的统一窗口叫 月度报表）。"""
    return "meli-mx-pipeline —— %s店铺报表" % label(region)


def preview_cmd(py, script, month, stores_, region):
    """「先看用哪批数据」的命令：run_reports --dry-run，带上区域。

    script 是美客多的 run_reports；别的平台的单元（"FA-CL"）换成那个平台自己的
    run_reports，带 --market。"""
    # script：脚本路径，或 ["-m", "包.模块"]
    if markets.SEP in str(region):
        pkg = markets.platform(markets.unit_platform(region)).package
        cmd = [py, "-m", pkg + ".run_reports", "--month", month, "--dry-run",
               "--market", region]
        if stores_:
            cmd += ["--stores"] + list(stores_)
        return cmd
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
        return store_config.unit_of(action["store"])
    except Exception:
        return action.get("region") or DEFAULT


def action_line(action):
    """「需要人工处理」列表里的一行：店名（平台 国家） —— 要做什么。"""
    return "%s（%s） —— %s" % (action["store"], unit_title(action_region(action)),
                               action["title"])


def rerun_plan(picked):
    """「处理好了：重跑」要起的那一轮。返回 dict：

    region   这些店的区域：一个就是它，几个就是 "MX,FA-CL"（run_monthly 一轮
             依次跑这几个，每家店归到它自己的区域）；一家都没有是 None
    stores   要重跑的店（去重、保持顺序）
    useless  重跑也没用的那几条（"店名（原因）"）
    groups   [(区域, [店])]
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
    order = [u for u, _ in lineup(ALL)]
    units = sorted((r for r, _ in groups), key=lambda u: order.index(u) if u in order
                   else len(order))
    return {"region": ",".join(units) or None,
            "stores": names, "useless": useless, "groups": groups}


# ────────────────────────────────────────────────────────── 计划任务

# 以前的版本确实登记过的计划任务名 —— 这是**历史记录**，不是按现在的国家清单
# 算出来的：以后加国家、加平台，都不会凭空多出一个"从没存在过的旧任务"。只有
# 某个版本真的登记过的名字才写进来（按国家的那几个，和 multi-region 版本用过的
# 全部国家任务 ML月度报表）。
LEGACY_TASKS = (("MX", "MX月度报表"), ("CL", "CL月度报表"), ("PE", "PE月度报表"),
                (ALL, "ML月度报表"))
OLD_ALL_TASKS = tuple(n for r, n in LEGACY_TASKS if r == ALL)


def legacy_tasks():
    """所有算"旧任务"的 (区域或 ALL, 任务名)。"""
    return list(LEGACY_TASKS)


def task_name(region):
    # 只跑勾选的几个区域，也还是那一个任务（月度报表）：只是参数不同。
    if region == ALL or is_list(region):
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
    m = re.search(r"--region\s+([\w,-]+)", args or "")
    return m.group(1).upper() if m else DEFAULT


def sched_key(key, region):
    """gui_settings.json 里定时设置的键。墨西哥沿用旧键（sched_day …），
    别的国家加后缀（sched_day_CL …），两边各记各的。"""
    return key if region == DEFAULT else "%s_%s" % (key, region)


# ────────────────────────────────────────────────────────── 每月几号（可以几天）
#
# 一个任务可以每月跑几天（docs/RESULTS_AND_PERIODS.md §6）：每一天都是全部店铺的
# 完整一轮，出的都是上个月；同一个月有几个完整版本时，最新的那个是定稿。
# gui_settings.json 里记成一个列表（sched_days…）；以前只有一个数（sched_day…），
# 照样读得出来（当成只有一天）。

MAX_DAY = 28                 # 2 月也有的日子：29–31 号有的月份没有，那个月就不跑了
DEFAULT_DAYS = [15]


def parse_days(text):
    """"3, 12，26" / "3 12 26" / "15" -> [3, 12, 26]。写错了抛 ValueError（中文说明）。"""
    parts = [p for p in re.split(r"[\s,，、;；]+", str(text or "").strip()) if p]
    if not parts:
        raise ValueError("至少要写一个日子，比如 15，或者 3, 12, 26")
    days = []
    for p in parts:
        if not p.isdigit() or not 1 <= int(p) <= MAX_DAY:
            raise ValueError("「%s」不是 1–%d 之间的日子。几个日子用逗号隔开，比如 3, 12, 26"
                             % (p, MAX_DAY))
        days.append(int(p))
    return sorted(set(days))


def days_text(days):
    """[3, 12, 26] -> "3、12、26"。"""
    return "、".join(str(d) for d in days)


def settings_days(settings, region=ALL):
    """gui_settings.json 里这个任务的几号：新的列表（sched_days…），没有就读旧的
    单个数（sched_day…，再不行 sched_day），都没有是 [15]。坏值跳过。"""
    s = settings or {}
    for key in (sched_key("sched_days", region), "sched_days"):
        v = s.get(key)
        if isinstance(v, (list, tuple)):
            try:
                got = sorted(set(int(x) for x in v if 1 <= int(x) <= MAX_DAY))
            except (TypeError, ValueError):
                got = []
            if got:
                return got
    for key in (sched_key("sched_day", region), "sched_day"):
        try:
            d = int(s.get(key))
        except (TypeError, ValueError):
            continue
        if 1 <= d <= MAX_DAY:
            return [d]
    return list(DEFAULT_DAYS)


def days_settings(days, region=ALL):
    """记进 gui_settings.json 的键值：列表，外加旧键里的第一天（旧版本的窗口读得懂）。"""
    days = sorted(set(int(d) for d in days))
    return {sched_key("sched_days", region): days, sched_key("sched_day", region): days[0]}


def _days_of(st):
    """task_state() 的结果 -> 几号的集合。"""
    d = (st or {}).get("DAYS")
    if d:
        return set(int(x) for x in d)
    return set(int(x) for x in re.findall(r"\d+", str((st or {}).get("DAY", ""))))


def _when(st):
    """task_state() 的结果 -> (几号的集合, 分钟数)。读不出来返回 None。"""
    try:
        h, m = (st or {}).get("TIME", "").split(":")
        days = _days_of(st)
        return (days, int(h) * 60 + int(m)) if days else None
    except (ValueError, TypeError, AttributeError):
        return None


def clash(a, b):
    """两个国家的任务会不会撞车。a、b 是 (几号, "HH:MM") 或 task_state() 的
    结果；有一个没登记就是 None。撞车时返回一句话，否则 None。几号可以是几天
    （[3, 12]），有一天相同就算同一天。

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
            days = set(int(d) for d in (day if isinstance(day, (list, tuple, set)) else [day]))
            return days, int(h) * 60 + int(m)
        except (ValueError, TypeError):
            return None
    wa, wb = norm(a), norm(b)
    if not wa or not wb or not (wa[0] & wb[0]):
        return None
    gap = abs(wa[1] - wb[1])
    if gap >= OVERLAP_HOURS * 60:
        return None
    return ("两个国家都定在每月 %d 号、开始时间只差 %d 小时 %02d 分。两国共用一个"
            "紫鸟，同一时间只能跑一轮：后到点的那个会因为前一个还没跑完而被跳过"
            "（退出码 5）。请把其中一个改到别的日子，或者错开 %d 小时以上。"
            % (min(wa[0] & wb[0]), gap // 60, gap % 60, OVERLAP_HOURS))


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
    days = sorted(_days_of(st))
    return "每月 %s 号 %s" % (days_text(days) if days else st.get("DAY", "?"), st.get("TIME", "?"))


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
    legacy = [(r, n, states.get(n)) for r, n in legacy_tasks() if states.get(n)]
    notes = []
    if main:
        reg = args_region(main.get("ARGS"))
        if reg != ALL:
            left_out = [u for u, _ in lineup(ALL) if u not in split_units(reg)]
            notes.append(("warn" if left_out else "good",
                          "「%s」只跑勾选的：%s。%s以后新加的平台 / 国家不会自动包含，要回来"
                          "勾上再点「设为每月自动运行」。"
                          % (ALL_TASK, "、".join(unit_title(u) for u in split_units(reg)),
                             "没有包含：%s。" % "、".join(unit_title(u) for u in left_out)
                             if left_out else "")))
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
        headline = "● 已设置：%s 自动跑%s" % (
            task_when(main), "全部平台和国家" if reg == ALL else "勾选的平台和国家")
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
