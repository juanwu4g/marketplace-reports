# -*- coding: utf-8 -*-
"""Windows 任务计划程序和这台电脑的检查：生成任务 XML、登记 / 查询 / 删除 /
立即运行任务、睡眠和锁屏自检。

从 schedule_gui.py 挪出来的，因为那边一 import 就要 tkinter；这里只有
subprocess，能在没有 tkinter 的机器上测（XML 长什么样、参数对不对）。

为什么用 XML 而不是 schtasks 的命令行参数、为什么查询走 PowerShell，见下面
各函数的说明（以前写在 schedule_gui.py 文件头）。简单说：有三项设置命令行
没有开关（错过时间就补跑、超时停止、电池），schtasks /query 的输出是本地化的。
"""
import datetime
import io
import os
import re
import subprocess
import tempfile
import threading

from core import region_choice

from core.paths import ROOT
RUNNER = os.path.join(ROOT, "run_monthly.cmd")

# 旧名字留着给还在引用它们的地方：墨西哥任务的名字和参数。
TASK = region_choice.task_name("MX")
REHEARSAL = region_choice.rehearsal_name("MX")
TASK_ARGS = region_choice.task_args("MX")
REHEARSAL_ARGS = region_choice.rehearsal_args("MX")

# 单个国家任务的超时，超了 Windows 会把**整条链**杀掉，报表那一步就再也跑不到。
# 现在这只是兜底：run_monthly 自己在 5 小时（它的 MAX_HOURS）就会掐断下载、
# 接着出报表和写告警。这里必须比那个数大，留出出报表的时间。
# 实测：9/17 那轮带 --month 的批次每家店 8–15 分钟，全程约 3 小时；9/21 凌晨
# 平台库存报表没响应，每家店 26.5 分钟，全程约 5.3 小时。杀掉一轮的代价是整月
# 没报表，多给一小时的代价只在真卡住时才付。
# 全部国家那个任务按国家数放宽：region_choice.time_limit_hours()。
TIME_LIMIT = "PT6H"

MONTHS_XML = "".join("<%s/>" % m for m in (
    "January February March April May June July August September October "
    "November December").split())


# ────────────────────────────────────────────────────────── 系统调用

def _run(argv, timeout=120, encoding="oem"):
    """跑一个命令，返回 (退出码, 输出)。永不抛异常。

    默认按 **oem** 解码。schtasks.exe 这类控制台程序把消息写成 OEM 代码页
    （中文系统是 cp936），按 utf-8 解出来是乱码 —— 而出错时那一行恰恰是唯一
    能说明原因的东西。英文系统的 OEM 是 437，纯 ASCII，两种解法看不出差别，
    所以这个问题只在中文机器上暴露。
    """
    try:
        p = subprocess.run(argv, capture_output=True, timeout=timeout)
        raw = (p.stdout or b"") + (p.stderr or b"")
        try:
            text = raw.decode(encoding, "replace")
        except LookupError:
            text = raw.decode("utf-8", "replace")
        return p.returncode, text
    except Exception as e:
        return -1, "%s: %s" % (e.__class__.__name__, e)


def ps(script, timeout=120):
    """跑一段 PowerShell。

    开头强制 UTF-8：任务名里有中文，不设的话输出按 cp936 回来，Python 这边
    按 utf-8 解会变成乱码。所以这一路按 utf-8 解，不走 oem。
    """
    return _run(["powershell", "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-Command",
                 "$OutputEncoding=[Console]::OutputEncoding="
                 "[Text.Encoding]::UTF8;" + script],
                timeout, encoding="utf-8")


def kv(out):
    """把 `KEY=值` 形式的输出解析成字典。"""
    d = {}
    for line in (out or "").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            d[k.strip()] = v.strip()
    return d


# ────────────────────────────────────────────────────────── 任务读写

def current_user():
    """任务要以谁的身份跑，写成 计算机名\\用户名。

    优先问 whoami：它返回的就是 Windows 认的那个主体。环境变量拼出来的形式
    在微软账号、域账号、改过计算机名的机器上不一定能被 schtasks 解析 ——
    解析不了时 schtasks 只回一句退出码 1 的错误，很难查。
    """
    rc, out = _run(["whoami"], timeout=20)
    who = (out or "").strip().splitlines()[0].strip() if out.strip() else ""
    if rc == 0 and "\\" in who:
        return who
    user = os.environ.get("USERNAME", "")
    domain = os.environ.get("USERDOMAIN", "")
    return ("%s\\%s" % (domain, user)) if domain else user


_STATE_PS = (
    "$t = Get-ScheduledTask -TaskName '{n}' -ErrorAction SilentlyContinue;"
    "if (-not $t) {{ 'EXISTS=0' }} else {{"
    "  'EXISTS=1';"
    "  'STATE=' + $t.State;"
    "  $i = Get-ScheduledTaskInfo -TaskName '{n}' -TaskPath $t.TaskPath;"
    "  'NEXT=' + $i.NextRunTime;"
    "  'LAST=' + $i.LastRunTime;"
    "  'RESULT=' + $i.LastTaskResult;"
    "  'XML=' + ((Export-ScheduledTask -TaskName '{n}') -replace '\\s+',' ')"
    "}};")


def parse_state(d):
    """kv() 读出来的一段 -> 任务现状（补上 ARGS / DAY / TIME）。不存在返回 None。"""
    if not d or d.get("EXISTS") != "1":
        return None
    d = dict(d)
    xml = d.get("XML", "")
    days = [int(x) for x in re.findall(r"<Day>(\d+)</Day>", xml)]
    at = re.search(r"<StartBoundary>[^<]*T(\d\d:\d\d)", xml)
    arg = re.search(r"<Arguments>([^<]*)</Arguments>", xml)
    d["ARGS"] = arg.group(1).strip() if arg else ""
    # DAYS: every run day of the trigger; DAY: the same as text ("15", "3,12,26")
    d["DAYS"] = sorted(set(days))
    d["DAY"] = ",".join(str(x) for x in d["DAYS"]) if days else "?"
    d["TIME"] = at.group(1) if at else "?"
    return d


def task_state(name):
    """任务现状。不存在返回 None。"""
    rc, out = ps(_STATE_PS.format(n=name))
    return parse_state(kv(out)) if rc == 0 else None


def split_states(out):
    """一次查好几个任务的输出（每段以 BEGIN=任务名 开头）-> {任务名: 现状或 None}。"""
    got, cur = {}, None
    for line in (out or "").splitlines():
        if line.startswith("BEGIN="):
            cur = line.split("=", 1)[1].strip()
            got[cur] = {}
        elif cur is not None and "=" in line:
            k, v = line.split("=", 1)
            got[cur][k.strip()] = v.strip()
    return dict((n, parse_state(d)) for n, d in got.items())


def task_states(names):
    """一次 PowerShell 查好几个任务。一个一个查每个要 2 秒，窗口一打开就是 8 秒。"""
    script = "".join("'BEGIN=%s';" % n + _STATE_PS.format(n=n) for n in names)
    rc, out = ps(script, timeout=180)
    if rc != 0:
        raise RuntimeError("查不了计划任务（PowerShell 退出码 %s）：%s"
                           % (rc, (out or "").strip()[:300]))
    got = split_states(out)
    return dict((n, got.get(n)) for n in names)


def run_days(day):
    """task_xml 的 day -> 排好序、去了重的几号：15 -> [15]；[26, 3, 12] -> [3, 12, 26]。
    None -> None（没有触发器）。1–31 以外的数抛 ValueError。"""
    if day is None:
        return None
    days = sorted(set(int(x) for x in (day if isinstance(day, (list, tuple, set)) else [day])))
    if not days or any(not 1 <= x <= 31 for x in days):
        raise ValueError("每月几号要在 1–31 之间：%r" % (day,))
    return days


def task_xml(args=TASK_ARGS, day=None, hour=0, minute=0, time_limit=TIME_LIMIT,
             who=None):
    """生成任务 XML。

    day 为 None 时不带触发器 —— 试跑任务只靠手工触发，不该自己跑起来。
    day 可以是一个数（15）或几个（[3, 12, 26]）：一个 CalendarTrigger，里面几个
    <Day>，每个都在同一个钟点跑（docs/RESULTS_AND_PERIODS.md §6）。
    args 是传给 run_monthly.cmd 的参数，至少要有 --entrance scheduled：
    日志靠它标出"这一轮是定时任务起的"。
    """
    who = who or current_user()
    trigger = ""
    days = run_days(day)
    if days:
        # StartBoundary 的日期部分只决定"从哪天起生效"，具体哪天跑由
        # DaysOfMonth 决定。取今天，避免写一个已经过去很久的日期。
        today = datetime.date.today()
        trigger = (
            "<CalendarTrigger>"
            "<StartBoundary>%04d-%02d-%02dT%02d:%02d:00</StartBoundary>"
            "<Enabled>true</Enabled>"
            "<ScheduleByMonth><DaysOfMonth>%s</DaysOfMonth>"
            "<Months>%s</Months></ScheduleByMonth>"
            "</CalendarTrigger>"
            % (today.year, today.month, today.day, hour, minute,
               "".join("<Day>%d</Day>" % x for x in days), MONTHS_XML))
    return (
        '<?xml version="1.0" encoding="UTF-16"?>'
        '<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">'
        "<RegistrationInfo><Description>"
        "每月自动下载上个月的数据并生成财务报表。由 月度报表.cmd（app.py）创建。"
        "</Description></RegistrationInfo>"
        "<Triggers>%s</Triggers>"
        "<Principals><Principal id=\"Author\">"
        "<UserId>%s</UserId>"
        # InteractiveToken = 「只在用户登录时运行」。紫鸟是带界面的 Electron
        # 程序，换成 Password/S4U 会被丢进 session 0，浏览器根本起不来。
        "<LogonType>InteractiveToken</LogonType>"
        "<RunLevel>LeastPrivilege</RunLevel>"
        "</Principal></Principals>"
        "<Settings>"
        # 上一轮还没跑完就到了下一个触发点：忽略新的。run_monthly 自己也有
        # 锁文件，这里是第二道。
        "<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>"
        "<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>"
        "<StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>"
        "<AllowHardTerminate>true</AllowHardTerminate>"
        "<StartWhenAvailable>true</StartWhenAvailable>"
        "<RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>"
        "<IdleSettings><StopOnIdleEnd>false</StopOnIdleEnd>"
        "<RestartOnIdle>false</RestartOnIdle></IdleSettings>"
        "<AllowStartOnDemand>true</AllowStartOnDemand>"
        "<Enabled>true</Enabled><Hidden>false</Hidden>"
        "<RunOnlyIfIdle>false</RunOnlyIfIdle>"
        "<WakeToRun>true</WakeToRun>"
        "<ExecutionTimeLimit>%s</ExecutionTimeLimit>"
        "<Priority>7</Priority>"
        "</Settings>"
        "<Actions Context=\"Author\"><Exec>"
        "<Command>%s</Command>%s"
        "<WorkingDirectory>%s</WorkingDirectory>"
        "</Exec></Actions></Task>"
        % (trigger, who, time_limit, RUNNER,
           ("<Arguments>%s</Arguments>" % args) if args else "", ROOT))


def install(name, xml):
    """导入任务 XML。返回 (退出码, 说明)。

    必须写成 UTF-16：schtasks /xml 读 UTF-8 文件会报 ERROR: 无效的 XML。

    失败时**不删**那个临时 XML，并把路径、身份、原始报错一起带回去。
    schtasks 失败只给一个退出码 1，不把这些摆出来根本没法查。
    """
    fd, path = tempfile.mkstemp(prefix="mxtask_", suffix=".xml")
    os.close(fd)
    with io.open(path, "w", encoding="utf-16") as f:
        f.write(xml)
    rc, out = _run(["schtasks", "/create", "/tn", name, "/xml", path, "/f"])
    if rc == 0:
        try:
            os.remove(path)
        except OSError:
            pass
        return rc, out
    detail = [
        (out or "").strip() or "(schtasks 没有输出)",
        "",
        "排查用的信息：",
        "  任务身份：%s" % current_user(),
        "  要跑的程序：%s" % RUNNER,
        "  工作目录：%s" % ROOT,
        "  任务定义文件（没删，可以打开看）：%s" % path,
        "",
        "想看更完整的报错，把下面这行贴到「命令提示符」里执行：",
        '  schtasks /create /tn "%s" /xml "%s" /f' % (name, path),
    ]
    return rc, "\n".join(detail)


# ────────────────────────────────────────────────────────── 机器自检

def machine_checks(region="ALL"):
    """返回 [(通过?, 标题, 说明)]。全部只读，不改系统。region 可以是 ALL。"""
    out = []

    venv = os.path.join(ROOT, ".venv", "Scripts", "python.exe")
    out.append((os.path.exists(venv), "程序环境",
                "装好了" if os.path.exists(venv) else
                "没装好（缺 .venv 文件夹）。找技术的人跑一次安装命令。"))

    out.append((os.path.exists(RUNNER), "自动跑的脚本",
                "在" if os.path.exists(RUNNER) else
                "找不到 run_monthly.cmd，程序文件不全。"))

    groups = region_choice.lineup(region)
    out.append((bool(groups), "要跑哪些店（%s）" % region_choice.label(region),
                region_choice.scope_text(groups) if groups
                else "读不出店铺名单，config.json 可能有问题。"))

    # 睡眠：powercfg 的标签是本地化的，但这一段永远是 5 行带 0x 的值，
    # 顺序固定为 最小/最大/步进/交流/电池，所以取倒数第二行（交流）。
    rc, txt = _run(["powercfg", "/q", "SCHEME_CURRENT", "SUB_SLEEP", "STANDBYIDLE"])
    hexes = re.findall(r"0x[0-9a-fA-F]{8}", txt or "")
    if rc == 0 and len(hexes) >= 2:
        ac = int(hexes[-2], 16)
        out.append((ac == 0, "电脑会不会睡着",
                    "设成了永不睡眠" if ac == 0 else
                    "闲置 %d 分钟就睡。睡着了就跑不了 —— 下面可以一键改。"
                    % (ac // 60)))
    else:
        out.append((None, "电脑会不会睡着", "读不出来，请手动确认电源设置。"))

    rc, txt = ps("'V=' + (Get-ItemProperty "
                 "'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon' "
                 "-Name AutoAdminLogon -ErrorAction SilentlyContinue).AutoAdminLogon")
    auto = kv(txt).get("V", "")
    out.append((auto == "1", "开机要不要输密码",
                "已设成开机自动登录" if auto == "1" else
                "要输密码。万一停电重启，没人输密码就没人登录，任务也跑不了。"
                "设置办法：开始菜单搜 netplwiz，取消勾选那个要密码的选项。"))

    rc, txt = ps("$d='HKCU:\\Control Panel\\Desktop';"
                 "'S=' + (Get-ItemProperty $d -Name ScreenSaverIsSecure "
                 "-ErrorAction SilentlyContinue).ScreenSaverIsSecure")
    sec = kv(txt).get("S", "")
    out.append((sec != "1", "会不会自动锁屏",
                "不会自动锁屏" if sec != "1" else
                "屏保会锁屏。锁屏后浏览器可能停住，建议关掉："
                "设置 → 个性化 → 锁屏界面 → 屏幕保护程序，取消「恢复时显示登录屏幕」。"))
    return out


def fix_power():
    """把三项电源超时改成永不。powercfg /change 要管理员，所以走 UAC。"""
    return ps(
        "Start-Process powershell -Verb RunAs -Wait -ArgumentList "
        "'-NoProfile','-Command',"
        "'powercfg /change standby-timeout-ac 0; "
        "powercfg /change hibernate-timeout-ac 0; "
        "powercfg /change monitor-timeout-ac 0'", timeout=180)


def delete_task(name):
    """撤掉一个任务。返回 (退出码, 输出)。"""
    return _run(["schtasks", "/delete", "/tn", name, "/f"])


def run_task(name):
    """让任务立刻跑一次（按它登记的参数）。返回 (退出码, 输出)。"""
    return _run(["schtasks", "/run", "/tn", name])


def all_task_states():
    """全部国家那个任务 + 每个国家的旧任务：{任务名: 现状或 None}。"""
    return task_states([region_choice.ALL_TASK] + [n for _, n in
                                                  region_choice.legacy_tasks()])


def rehearse(region):
    """建一个临时任务只出报表、立刻跑、等它跑完、删掉。返回 ("ok", 结果码) 或 ("err", 说明)。"""
    rehearsal = region_choice.rehearsal_name(region)
    rc, out = install(rehearsal, task_xml(args=region_choice.rehearsal_args(region)))
    if rc != 0:
        return ("err", "建不了临时任务：%s" % (out.strip() or rc))
    rc, out = run_task(rehearsal)
    if rc != 0:
        delete_task(rehearsal)
        return ("err", "临时任务启动不了：%s" % (out.strip() or rc))
    # 轮询到不再是 Running 为止。出 13 份报表实测 60～70 秒；全部国家多几十秒。
    for _ in range(150):
        st = task_state(rehearsal) or {}
        if st.get("STATE") != "Running":
            break
        threading.Event().wait(2)
    st = task_state(rehearsal) or {}
    delete_task(rehearsal)
    return ("ok", st.get("RESULT", "?"))
