# -*- coding: utf-8 -*-
"""定时任务的可视化入口：不用敲一条命令就能把每月自动跑配好。

    双击 定时设置.cmd    （推荐，它会用对解释器）
    或   .venv\\Scripts\\python schedule_gui.py

做五件事：建/改每月的计划任务、试跑一次、看状态、看日志、检查这台机器够不
够格无人值守。**业务逻辑一行都没有** —— 真正跑的是 run_monthly.cmd，这里只是
把 Windows 任务计划程序那一堆勾选项替人点好。

和「启动」窗口跑的是同一个过程
------------------------------
计划任务的动作是 `run_monthly.cmd --entrance scheduled`；「启动」窗口点按钮
起的是同一个 run_monthly.py，只是 `--entrance manual`。两边唯一的区别就是
这里多了"几号几点"。窗口下方的「运行日志」和「启动」窗口里的是同一个控件
（log_view.py），看的是同一批文件（logs\runs），所以在这里也能实时看到、
也能中止任何一轮 —— 包括窗口开着的时候到点自己跑起来的那轮。

两个国家，两个任务
------------------
墨西哥的任务叫「MX月度报表」（生产机器上早就登记了，名字不能改），智利的叫
「CL月度报表」，参数分别带 --region MX / --region CL。窗口上方选哪个国家，
「设为每月自动跑 / 试跑 / 现在就跑 / 取消」就只动那个国家的任务。两国共用
一个紫鸟，锁是全局的：两个任务定在同一天、开始时间离得太近，后到点的那个
会被跳过 —— 窗口会提醒。规则在 region_choice.py，有测试。

**在这次改动之前登记的任务不带 --entrance scheduled**，日志里会被记成
「命令行」。再点一次「设为每月自动跑」就改过来了；状态栏会提醒。

为什么用 XML 而不是 schtasks 的命令行参数
-----------------------------------------
有三项设置命令行根本没有开关，只能在任务计划程序界面里手点：

* 「如果错过计划的开始时间，立即启动任务」（StartWhenAvailable）—— 15 号
  机器关着就补跑，对一台没人管的老 PC 来说这条最要紧
* 「如果任务运行超过 N 小时则停止」（ExecutionTimeLimit）—— run_monthly 自己
  会在 5 小时掐断下载并接着出报表，这一项只是它也卡死时的最后一道兜底
* 电池相关的两项 —— 笔记本上默认会拦住任务

让不写代码的人去 taskschd.msc 里找这三个勾，比让他们敲命令还难。所以这里
直接生成完整的任务 XML 再 `schtasks /create /xml` 导入，一次到位。

为什么查询走 PowerShell 而不是 schtasks
---------------------------------------
`schtasks /query` 的输出是本地化的，中文机器上列名和值都是中文，解析它等于
赌系统语言。`Get-ScheduledTaskInfo` 返回的是对象，属性名与语言无关。
"""
import datetime
import io
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading

import tkinter as tk
from tkinter import ttk, messagebox

import region_choice
import runlog
from log_view import LogView, fit_to_screen
from gui import ROOT, interpreter, store_list, load_settings, save_settings
# 报表输出目录的解析规则只有一份，在 run_monthly 里 —— 定时任务真正跑的是它。
from run_monthly import resolve_out_dir

# 左列说明文字的折行宽度（像素），和「启动」窗口取同一个值。
WRAP = 470

# 墨西哥任务的名字和参数。区域出现之后实际用的是 region_choice.task_name /
# task_args（每个国家一份），这两个常量只是留给还在引用它们的地方。
TASK = region_choice.task_name("MX")
REHEARSAL = region_choice.rehearsal_name("MX")
RUNNER = os.path.join(ROOT, "run_monthly.cmd")
LOG_DIR = os.path.join(ROOT, "logs")

# 计划任务传给 run_monthly.cmd 的参数。和「启动」窗口用的是同一个拼法
# （runlog.runner_argv），只有 entrance 不同；月份、店铺、输出目录都不写，
# 由 run_monthly 取默认值：上个月、全部店铺、主界面里选的输出目录。区域要写
# （--region MX / CL），见 region_choice.task_args。
TASK_ARGS = region_choice.task_args("MX")
REHEARSAL_ARGS = region_choice.rehearsal_args("MX")

# 任务超时，超了 Windows 会把**整条链**杀掉，报表那一步就再也跑不到。
# 现在这只是兜底：run_monthly 自己在 5 小时（它的 MAX_HOURS）就会掐断下载、
# 接着出报表和写告警。这里必须比那个数大，留出出报表的时间。
# 实测：9/17 那轮带 --month 的批次每家店 8–15 分钟，全程约 3 小时；9/21 凌晨
# 平台库存报表没响应，每家店 26.5 分钟，全程约 5.3 小时 —— 超过当时的 5 小时
# 上限被杀，当月没出报表。已经修掉了那晚浪费时间的两处（未开始的账期、下载
# 根本没开始时干等满 240 秒），坏的一晚也回到约 3.2 小时；这里再放宽到 6 小时
# 作为余量。杀掉一轮的代价是整月没报表，多给一小时的代价只在真卡住时才付。
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


def task_state(name):
    """任务现状。不存在返回 None。"""
    rc, out = ps(
        "$t = Get-ScheduledTask -TaskName '%s' -ErrorAction SilentlyContinue;"
        "if (-not $t) { 'EXISTS=0' } else {"
        "  'EXISTS=1';"
        "  'STATE=' + $t.State;"
        "  $i = Get-ScheduledTaskInfo -TaskName '%s' -TaskPath $t.TaskPath;"
        "  'NEXT=' + $i.NextRunTime;"
        "  'LAST=' + $i.LastRunTime;"
        "  'RESULT=' + $i.LastTaskResult;"
        "  'XML=' + ((Export-ScheduledTask -TaskName '%s') -replace '\\s+',' ')"
        "}" % (name, name, name))
    d = kv(out)
    if rc != 0 or d.get("EXISTS") != "1":
        return None
    xml = d.get("XML", "")
    day = re.search(r"<Day>(\d+)</Day>", xml)
    at = re.search(r"<StartBoundary>[^<]*T(\d\d:\d\d)", xml)
    arg = re.search(r"<Arguments>([^<]*)</Arguments>", xml)
    d["ARGS"] = arg.group(1).strip() if arg else ""
    d["DAY"] = day.group(1) if day else "?"
    d["TIME"] = at.group(1) if at else "?"
    return d


def task_xml(args=TASK_ARGS, day=None, hour=0, minute=0):
    """生成任务 XML。

    day 为 None 时不带触发器 —— 试跑任务只靠手工触发，不该自己跑起来。
    args 是传给 run_monthly.cmd 的参数，至少要有 --entrance scheduled：
    日志靠它标出"这一轮是定时任务起的"。
    """
    who = current_user()
    trigger = ""
    if day is not None:
        # StartBoundary 的日期部分只决定"从哪天起生效"，具体哪天跑由
        # DaysOfMonth 决定。取今天，避免写一个已经过去很久的日期。
        today = datetime.date.today()
        trigger = (
            "<CalendarTrigger>"
            "<StartBoundary>%04d-%02d-%02dT%02d:%02d:00</StartBoundary>"
            "<Enabled>true</Enabled>"
            "<ScheduleByMonth><DaysOfMonth><Day>%d</Day></DaysOfMonth>"
            "<Months>%s</Months></ScheduleByMonth>"
            "</CalendarTrigger>"
            % (today.year, today.month, today.day, hour, minute, day, MONTHS_XML))
    return (
        '<?xml version="1.0" encoding="UTF-16"?>'
        '<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">'
        "<RegistrationInfo><Description>"
        "每月自动下载上个月的数据并生成财务报表。由 schedule_gui.py 创建。"
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
        % (trigger, who, TIME_LIMIT, RUNNER,
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

def machine_checks(region="MX"):
    """返回 [(通过?, 标题, 说明)]。全部只读，不改系统。"""
    out = []

    venv = os.path.join(ROOT, ".venv", "Scripts", "python.exe")
    out.append((os.path.exists(venv), "程序环境",
                "装好了" if os.path.exists(venv) else
                "没装好（缺 .venv 文件夹）。找技术的人跑一次安装命令。"))

    out.append((os.path.exists(RUNNER), "自动跑的脚本",
                "在" if os.path.exists(RUNNER) else
                "找不到 run_monthly.cmd，程序文件不全。"))

    names = store_list(region)
    out.append((bool(names), "要跑哪些店（%s）" % region_choice.label(region),
                "%d 家：%s" % (len(names), "、".join(names[:3]) +
                              ("…" if len(names) > 3 else ""))
                if names else "读不出店铺名单，config.json 可能有问题。"))

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


# ────────────────────────────────────────────────────────── 界面

class SchedApp(object):
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.busy = False
        self.busy_label = ""
        fit_to_screen(root, 1680, 900)
        root.minsize(1000, 600)

        s = load_settings()
        # 选的是哪个国家的任务。和「启动」窗口记在同一个键里：在那边选了智利，
        # 打开这边也是智利。
        self.region = region_choice.remembered(s)
        root.title(self._title())
        self.day = tk.IntVar()
        self.hour = tk.IntVar()
        self.minute = tk.IntVar()
        self._load_time(s)
        self.other = None               # 另一个国家的任务现状（查撞车用）

        # 和「启动」窗口同一个排法：左边是设置和按钮，右边整列给运行日志 ——
        # 两个窗口看日志的地方长得一样、一样大。
        self.left = ttk.Frame(root)
        self.left.pack(side="left", fill="y")
        self.right = ttk.Frame(root)
        self.right.pack(side="left", fill="both", expand=True)
        self._build_status()
        self._build_form()
        self._build_actions()
        self._build_log()
        self.root.after(80, self._drain)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.refresh()

    # ---------------------------------------------------------- 区域

    def _title(self):
        return "每月自动出报表 —— 设置（%s）" % region_choice.label(self.region)

    def _task(self):
        return region_choice.task_name(self.region)

    def _other_region(self):
        others = [r for r in region_choice.REGIONS if r != self.region]
        return others[0] if others else None

    def _load_time(self, s):
        """界面上的几号几点：各国各记各的（region_choice.sched_key）。"""
        k = lambda name: region_choice.sched_key(name, self.region)
        self.day.set(int(s.get(k("sched_day"), 15)))
        self.hour.set(int(s.get(k("sched_hour"), 2)))
        self.minute.set(int(s.get(k("sched_minute"), 0)))

    def _save_time(self, **extra):
        k = lambda name: region_choice.sched_key(name, self.region)
        d = dict(load_settings(), **extra)
        d.update({k("sched_day"): self.day.get(), k("sched_hour"): self.hour.get(),
                  k("sched_minute"): self.minute.get()})
        save_settings(d)

    def _pick_region(self):
        region = region_choice.from_label(self.region_box.get())
        if region == self.region:
            return
        self._save_time()                       # 先记下切走之前那个国家的时间
        self.region = region
        s = load_settings()
        save_settings(dict(s, region=region))
        self.root.title(self._title())
        self._load_time(s)
        self.other = None
        self.clash_note.set("")
        self.status.set("正在读取…")
        self.refresh()

    # ---------------------------------------------------------- 布局

    def _build_status(self):
        top = ttk.Frame(self.left)
        top.pack(fill="x", padx=(12, 6), pady=(12, 0))
        ttk.Label(top, text="国家").pack(side="left")
        self.region_box = ttk.Combobox(top, values=region_choice.labels(),
                                       width=10, state="readonly")
        self.region_box.set(region_choice.label(self.region))
        self.region_box.pack(side="left", padx=(6, 12))
        self.region_box.bind("<<ComboboxSelected>>", lambda e: self._pick_region())
        ttk.Label(top, text="每个国家一个任务；下面的按钮只动所选国家的那个。",
                  foreground="#666").pack(side="left")

        f = ttk.LabelFrame(self.left, text=" 现在是什么状态 ", padding=10)
        f.pack(fill="x", padx=(12, 6), pady=(6, 6))
        self.status = tk.StringVar(value="正在读取…")
        self.status_lbl = ttk.Label(f, textvariable=self.status,
                                    font=("Microsoft YaHei UI", 11, "bold"))
        self.status_lbl.pack(anchor="w")
        self.detail = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.detail, foreground="#555",
                  justify="left", wraplength=WRAP).pack(anchor="w", pady=(4, 0))
        # 定时跑是没人看着的：上一轮要是有店铺要人去登录 / 验证，这里要一眼
        # 看得到，而且说清楚去哪儿处理。处理和补跑在「启动」窗口里做。
        self.human_note = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.human_note, foreground="#b42318",
                  justify="left", wraplength=WRAP,
                  font=("", 10, "bold")).pack(anchor="w", pady=(4, 0))
        # 两国任务撞车（同一天、时间离得太近）的提醒。
        self.clash_note = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.clash_note, foreground="#9a6700",
                  justify="left", wraplength=WRAP).pack(anchor="w", pady=(4, 0))
        ttk.Button(f, text="刷新状态", command=self.refresh, width=12).pack(
            anchor="w", pady=(8, 0))

    def _build_form(self):
        f = ttk.LabelFrame(self.left, text=" 什么时候自动跑 ", padding=10)
        f.pack(fill="x", padx=(12, 6), pady=6)
        row = ttk.Frame(f)
        row.pack(anchor="w")
        ttk.Label(row, text="每月").pack(side="left")
        ttk.Spinbox(row, from_=1, to=28, width=4, textvariable=self.day).pack(
            side="left", padx=4)
        ttk.Label(row, text="号").pack(side="left")
        ttk.Spinbox(row, from_=0, to=23, width=4, format="%02.0f",
                    textvariable=self.hour).pack(side="left", padx=(16, 4))
        ttk.Label(row, text="点").pack(side="left")
        ttk.Spinbox(row, from_=0, to=59, width=4, increment=5, format="%02.0f",
                    textvariable=self.minute).pack(side="left", padx=(8, 4))
        ttk.Label(row, text="分").pack(side="left")
        ttk.Label(f, foreground="#555", justify="left", wraplength=WRAP, text=(
            "出的永远是【上个月】的报表。9 月 15 号跑，出的是 8 月。\n"
            "选 15 号是因为上个月的账单到那时已经结清，数字不会再变。\n"
            "选凌晨是因为跑的时候会重开紫鸟浏览器，会关掉别人正开着的窗口。\n"
            "日期最大 28：29～31 号遇上二月就不会触发。")
                  ).pack(anchor="w", pady=(8, 0))

    def _build_actions(self):
        f = ttk.LabelFrame(self.left, text=" 操作 ", padding=10)
        f.pack(fill="x", padx=(12, 6), pady=6)
        rows = [
            ("设为每月自动跑", self.do_install,
             "按上面选的时间登记到 Windows。已经登记过就按新时间改掉。"),
            ("试跑一次（1 分钟）", self.do_rehearse,
             "只出报表，不下载、不开浏览器，跑完自动收拾干净。"
             "用来确认这台电脑到时候真能自动跑起来 —— 别等到 15 号才发现跑不了。"),
            ("现在就跑一次", self.do_run_now,
             "立刻按正式任务跑一遍：下载 + 出报表，约 3 小时。不影响以后的自动跑。"
             "跑起来以后，右边的「运行日志」会实时显示。"),
            ("检查这台电脑", self.do_check,
             "看这台电脑能不能撑住半夜没人管地跑：会不会睡着、会不会锁屏、"
             "开机要不要有人输密码。只看不改。"),
            ("取消每月自动跑", self.do_delete,
             "只撤掉 Windows 里的登记。程序、数据、已出的报表都不动。"),
            ("打开日志文件夹", self.do_logs,
             "每一轮的完整记录都在这里（logs\\runs\\日期_时间_入口.log），"
             "手动跑的和自动跑的放在一起。平时不用开：右边的「运行日志」里就能看、能挑。"),
        ]
        self.buttons = []
        for i, (label, cmd, desc) in enumerate(rows):
            b = ttk.Button(f, text=label, command=cmd, width=18)
            b.grid(row=i, column=0, sticky="w", pady=3)
            ttk.Label(f, text=desc, foreground="#555", justify="left",
                      wraplength=WRAP - 170).grid(
                row=i, column=1, sticky="w", padx=12)
            self.buttons.append(b)

    def _build_log(self):
        # 两页：运行日志是和「启动」窗口共用的那个查看器，看的是 run_monthly
        # 写的日志文件；操作记录是本窗口自己的话（登记成没成、自检结果），
        # 它们不是某一轮运行的输出，不该混进那个标准格式里。
        self.tabs = ttk.Notebook(self.right)
        self.tabs.pack(fill="both", expand=True, padx=(6, 12), pady=12)
        self.view = LogView(self.tabs, who="定时设置窗口", height=10,
                            on_state=self._on_run_state)
        self.tabs.add(self.view, text=" 运行日志（手动、定时共用） ")
        f = ttk.Frame(self.tabs)
        self.tabs.add(f, text=" 本窗口的操作记录 ")
        self.ops_tab = f
        self.txt = tk.Text(f, wrap="word", height=10,
                           font=("Consolas", 9), background="#fbfbfb")
        self.txt.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(f, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set, state="disabled")
        sb.pack(side="right", fill="y")
        for tag, color in (("good", "#1a7f37"), ("warn", "#9a6700"),
                           ("bad", "#b42318"), ("dim", "#666666")):
            self.txt.tag_configure(tag, foreground=color)

    # ---------------------------------------------------------- 工具

    def log(self, line, tag=None):
        self.txt.configure(state="normal")
        self.txt.insert("end", line + "\n", tag or ())
        self.txt.see("end")
        self.txt.configure(state="disabled")
        if line.strip():
            try:
                self.tabs.select(self.ops_tab)     # 有话要说就翻到这一页
            except tk.TclError:
                pass

    def _show_run_log(self):
        """翻到运行日志那页，并回到"跟随当前运行"。"""
        try:
            self.view.follow_current()
            self.tabs.select(self.view)
        except tk.TclError:
            pass

    def _on_run_state(self, st):
        # 一轮跑完了（不管谁起的）：上面的"上次结果"也该跟着变。
        if st and st.get("state") != "running" and not self.busy:
            self.refresh()

    def _lock(self, busy, label=""):
        self.busy = busy
        self.busy_label = label
        for b in self.buttons:
            b.configure(state="disabled" if busy else "normal")

    def _work(self, fn, done, label="", lock=True):
        """把慢活丢到子线程。Tk 控件只在主线程碰。

        lock=False 给只读查询用（读一次任务状态要 2 秒左右）。查询期间按钮
        照常能点，否则刚打开窗口那两秒什么都点不了。
        """
        if lock and self.busy:
            messagebox.showwarning(
                "还在忙",
                "「%s」还没跑完，等它结束再点。" % (self.busy_label or "上一个操作"))
            return
        if lock:
            self._lock(True, label)

        def body():
            try:
                self.q.put((done, fn(), lock))
            except Exception as e:
                self.q.put((done, ("err", "%s: %s" % (e.__class__.__name__, e)), lock))
        threading.Thread(target=body, daemon=True).start()

    def _drain(self):
        try:
            while True:
                fn, payload, locked = self.q.get_nowait()
                # 先解锁再回调。反过来的话，回调里想再发起一个操作（比如建完
                # 任务顺手刷新状态、自检完去修电源）会撞上 busy 还是 True，
                # 于是什么都没做，只弹一句"还在忙"——这正是之前的表现。
                if locked:
                    self._lock(False)
                fn(payload)
        except queue.Empty:
            pass
        self.root.after(80, self._drain)

    # ---------------------------------------------------------- 动作

    def refresh(self):
        # 只读查询，不锁按钮：它要 2 秒，锁上的话刚开窗口那两秒点什么都没反应。
        # 另一个国家的任务也查一次：要看两个任务会不会撞车。
        region, other = self.region, self._other_region()

        def body():
            mine = task_state(region_choice.task_name(region))
            theirs = task_state(region_choice.task_name(other)) if other else None
            return region, mine, theirs
        self._work(body, self._show_states, lock=False)

    def _show_states(self, r):
        if isinstance(r, tuple) and len(r) == 3:
            region, mine, theirs = r
            if region != self.region:
                return                  # 查的时候人已经切到另一个国家了
            self.other = theirs
            self._show_clash()
            self._show_state(mine)
        else:
            self._show_state(r)          # ("err", 说明)

    def _show_clash(self, day=None, hm=None):
        """两国任务撞车时在状态区里说一句。day/hm 给了就按"要登记的时间"算。"""
        mine = (day, hm) if day is not None else None
        if mine is None:
            mine = (self.day.get(), "%02d:%02d" % (self.hour.get(),
                                                   self.minute.get()))
        msg = region_choice.clash(mine, self.other) if self.other else None
        self.clash_note.set(
            "⚠ 和%s的任务撞车：%s" % (region_choice.label(self._other_region()), msg)
            if msg else "")
        return msg

    def _show_human(self):
        try:
            got = runlog.pending_human_actions()
        except Exception:
            got = {"actions": []}
        acts = got.get("actions") or []
        stores = []
        for a in acts:
            if a["store"] not in stores:
                stores.append(a["store"])
        self.human_note.set(
            "⚠ 上一轮有 %d 家店铺需要人工处理：%s。\n双击「启动.cmd」，在左边"
            "「需要人工处理」里打开店铺处理，再点「重跑这些店铺并重出全部报表」。"
            % (len(stores), "；".join(
                "%s（%s）" % (s, "、".join(sorted(set(
                    a["title"] for a in acts if a["store"] == s))))
                for s in stores)) if acts else "")

    def _show_state(self, st):
        self._show_human()
        if isinstance(st, tuple):                    # 出错了
            self.status.set("读不到任务状态")
            self.status_lbl.configure(foreground="#b42318")
            self.log(st[1], "bad")
            return
        if not st:
            self.clash_note.set("")
            self.status.set("● %s还没设置，现在不会自动跑"
                            % region_choice.label(self.region))
            self.status_lbl.configure(foreground="#9a6700")
            self.detail.set("选好时间，点「设为每月自动跑」。")
            return
        code = st.get("RESULT", "")
        # 0–7 是 run_monthly.py 自己的退出码（见它的文件头），267xxx 是
        # Windows 任务计划程序的状态码。都翻成人话，不要让人去查代码表。
        verdict = {"0": "上次跑完，一切正常",
                   "1": "上次跑完了，但有几项要人看一眼（看右边的「运行日志」）",
                   "2": "上次没跑成：浏览器没能启动 / 没有一家店下到数据 / "
                        "有报表没出来",
                   "3": "上次没跑成：店铺名单对不上，要改配置",
                   "4": "上次没跑成：配置文件读不了",
                   "5": "上次没跑：另一轮（可能是手动点的）还没结束，这次被跳过了",
                   "6": "上次被人点了「中止」",
                   "7": "上次下载超过 5 小时被掐断，报表是按已下到的数据出的",
                   "267009": "正在跑",
                   "267011": "还没跑过",
                   "267014": "上次被 Windows 中途停掉了（超过 6 小时，或有人在"
                             "任务计划程序里停的）"}.get(
            code, "上次没跑成（代码 %s），看右边的「运行日志」" % code)
        if region_choice.args_region(st.get("ARGS")) != self.region:
            verdict += ("\n⚠ 这个任务登记的参数（%s）跑的不是%s的店铺：再点一次"
                        "「设为每月自动跑」改过来。"
                        % (st.get("ARGS") or "无", region_choice.label(self.region)))
        if "--entrance" not in st.get("ARGS", ""):
            # 这次改动之前登记的任务。照样能跑，跑的也是同一个过程，只是日志
            # 的"入口"一列会写成「命令行」而不是「定时」。
            verdict += ("\n⚠ 这个任务是旧版登记的：再点一次「设为每月自动跑」，"
                        "日志里才会标成「定时」。")
        self.status.set("● %s已设置：每月 %s 号 %s 自动跑"
                        % (region_choice.label(self.region), st.get("DAY"),
                           st.get("TIME")))
        self.status_lbl.configure(foreground="#1a7f37")
        # 报表写到哪儿必须摆在眼前：定时任务和主界面曾经各用各的目录，
        # 结果报表出了却落在别处，看起来像"只下载没出报表"。
        out, src = resolve_out_dir()
        self.detail.set("下次：%s\n上次：%s\n%s\n报表写到：%s（%s）"
                        % (st.get("NEXT") or "还没跑过",
                           st.get("LAST") or "还没跑过", verdict,
                           os.path.abspath(out), src))
        # 界面上的时间跟已登记的对齐，免得改了没保存的人看错
        try:
            self.day.set(int(st.get("DAY")))
            h, m = st.get("TIME", "02:00").split(":")
            self.hour.set(int(h))
            self.minute.set(int(m))
        except (ValueError, TypeError):
            pass
        self._show_clash()

    def do_install(self):
        if not os.path.exists(RUNNER):
            messagebox.showerror(
                "程序文件不全",
                "找不到 run_monthly.cmd：\n%s\n\n"
                "程序文件不全，找技术的人看一下。" % RUNNER)
            return
        d, h, m = self.day.get(), self.hour.get(), self.minute.get()
        region = self.region
        name = region_choice.label(region)
        clash = self._show_clash(d, "%02d:%02d" % (h, m))
        warn = ("\n⚠ %s\n\n" % clash) if clash else ""
        if not messagebox.askokcancel(
                "设为每月自动跑",
                "以后每月 %d 号 %02d:%02d，这台电脑会自己做两件事：\n"
                "  ① 下载上个月的数据（%s的店铺全部）\n"
                "  ② 出上个月的报表\n\n"
                "这台电脑那时候必须是开着的、已经登录、没睡着、没锁屏。\n"
                "（下载要开浏览器，没人登录就开不起来。）\n%s\n"
                "继续？" % (d, h, m, name, warn)):
            return
        self._save_time()
        task = region_choice.task_name(region)
        args = region_choice.task_args(region)
        self.log("")
        self.log("正在登记「%s」：每月 %d 号 %02d:%02d（%s）" % (task, d, h, m, args))
        self._work(lambda: install(task, task_xml(args, day=d, hour=h, minute=m)),
                   self._after_install, label="设为每月自动跑")

    def _after_install(self, r):
        rc, out = r
        if rc == 0:
            self.log("✓ 登记好了。另外这几项也一并设上了：那天电脑要是关着，"
                     "开机后会补跑；用电池也照跑。下载超过 5 小时会被掐断并"
                     "接着出报表，整轮超过 6 小时 Windows 会强行停掉。", "good")
            self.log("建议现在点一次「试跑一次」，1 分钟就知道到时候跑不跑得起来。",
                     "dim")
            self.refresh()
        else:
            self.log("✗ 没登记成（错误码 %s）" % rc, "bad")
            self.log(out.strip() or "(没有更多信息)", "bad")

    def do_rehearse(self):
        if not messagebox.askokcancel(
                "试跑一次",
                "会让 Windows 真的把程序拉起来跑一遍，但只出报表 ——\n"
                "不下载、不开浏览器，跑完自己收拾干净。约 1 分钟。\n\n"
                "这一步是为了提前确认：到了 15 号半夜没人管的时候，\n"
                "这台电脑确实能把程序跑起来。\n\n开始？"):
            return
        self.log("")
        self.log("=" * 64)
        self.log(">>> 试跑（%s）  %s" % (region_choice.label(self.region),
                                       datetime.datetime.now().strftime("%H:%M:%S")))
        region = self.region
        self._work(lambda: self._rehearse_body(region), self._after_rehearse,
                   label="试跑一次")
        self._show_run_log()

    def _rehearse_body(self, region=None):
        region = region or self.region
        rehearsal = region_choice.rehearsal_name(region)
        rc, out = install(rehearsal, task_xml(
            args=region_choice.rehearsal_args(region)))
        if rc != 0:
            return ("err", "建不了临时任务：%s" % (out.strip() or rc))
        rc, out = _run(["schtasks", "/run", "/tn", rehearsal])
        if rc != 0:
            _run(["schtasks", "/delete", "/tn", rehearsal, "/f"])
            return ("err", "临时任务启动不了：%s" % (out.strip() or rc))
        # 轮询到不再是 Running 为止。出 13 份报表实测 60～70 秒。
        for _ in range(90):
            st = task_state(rehearsal) or {}
            if st.get("STATE") != "Running":
                break
            threading.Event().wait(2)
        st = task_state(rehearsal) or {}
        _run(["schtasks", "/delete", "/tn", rehearsal, "/f"])
        return ("ok", st.get("RESULT", "?"))

    def _after_rehearse(self, r):
        kind, val = r
        if kind == "err":
            self.log("✗ %s" % val, "bad")
            return
        meaning = {"0": ("✓ 试跑通过：Windows 能把程序拉起来，报表也出来了。"
                         "到时候会自动跑。", "good"),
                   "1": ("✓ 试跑通过：Windows 能把程序拉起来，报表出来了，"
                         "但有几项要人看一眼（这在每个月都很常见）。", "good"),
                   "267009": ("⚠ 还没跑完，过一会儿点「刷新」", "warn")}
        msg, tag = meaning.get(
            val, ("✗ 试跑没成功（代码 %s：%s）。原因在「运行日志」那一页"
                  "最近一轮「定时」的记录里。"
                  % (val, runlog.exit_meaning(val)[0]
                     if val.isdigit() and int(val) < 8 else "见日志"), "bad"))
        self.log(msg, tag)
        self.log("临时任务已经删掉了。", "dim")

    def do_run_now(self):
        if not messagebox.askokcancel(
                "现在就跑一次",
                "立刻完整跑一遍：下载 + 出报表，约 3 小时。\n\n"
                "跑的时候浏览器会自己反复开关，别去点它。\n"
                "这个窗口可以关掉，不影响后台继续跑。\n"
                "要中途停下来：点下面「运行日志」右上角的「中止这一轮」。\n\n开始？"):
            return
        cur = runlog.active_run()
        if cur:
            messagebox.showwarning(
                "正在运行",
                "已经有一轮在跑了（入口：%s）。等它结束，或先中止它。"
                % runlog.ENTRANCE_ZH.get(cur.get("entrance"), "?"))
            return
        task = self._task()
        self._work(lambda: self._run_now_body(task), self._after_run_now,
                   label="现在就跑一次")

    def _run_now_body(self, task=None):
        # 查任务在不在要 2 秒，放子线程里做，不然界面会卡住不动。
        task = task or self._task()
        if not task_state(task):
            return (-2, "")
        return _run(["schtasks", "/run", "/tn", task])

    def _after_run_now(self, r):
        rc, out = r
        if rc == -2:
            self.log("✗ %s还没设置过自动跑，先点「设为每月自动跑」。"
                     % region_choice.label(self.region), "bad")
        elif rc == 0:
            self.log("✓ 已经开始跑了，进度看「运行日志」那一页（已经帮你翻过去）。",
                     "good")
            self._show_run_log()
        else:
            self.log("✗ 没能启动：%s" % (out.strip() or rc), "bad")

    def do_check(self):
        self.log("")
        self.log("=" * 64)
        self.log(">>> 检查这台电脑")
        region = self.region
        self._work(lambda: machine_checks(region), self._after_check,
                   label="检查这台电脑")

    def _after_check(self, items):
        bad = 0
        for ok, title, note in items:
            mark, tag = ("  ✓", "good") if ok else (
                ("  ?", "warn") if ok is None else ("  ✗", "bad"))
            if ok is False:
                bad += 1
            self.log("%s  %-16s %s" % (mark, title, note), tag)
        if not bad:
            self.log("全都没问题，这台电脑可以半夜没人管地自己跑。", "good")
            return
        self.log("有 %d 项不合格。不处理的话，到时候可能跑不起来。" % bad, "warn")
        if messagebox.askyesno(
                "要现在设好吗",
                "可以帮你把「不睡眠 / 不休眠 / 屏幕不自动关」一次设好。\n\n"
                "会弹一个 Windows 的管理员确认框，点「是」就行。\n\n"
                "另外两项（开机自动登录、屏保锁屏）需要你手动设，\n"
                "上面列表里写了在哪里设。"):
            self._work(fix_power, self._after_fix, label="设置电源")

    def _after_fix(self, r):
        rc, out = r
        self.log("已设置。再点一次「检查这台电脑」确认。" if rc == 0
                 else "没能设置：%s" % (out.strip() or rc),
                 "good" if rc == 0 else "bad")

    def do_delete(self):
        if not messagebox.askokcancel(
                "取消每月自动跑",
                "以后%s的店铺不再自动跑了（另一个国家的任务不动）。\n\n"
                "只是撤掉 Windows 里的登记 —— 程序、下载好的数据、已经出的\n"
                "报表都不会动，手动跑也照常。想恢复就再点「设为每月自动跑」。\n\n"
                "确定？" % region_choice.label(self.region)):
            return
        task = self._task()
        self._work(lambda: _run(["schtasks", "/delete", "/tn", task, "/f"]),
                   self._after_delete, label="取消每月自动跑")

    def _after_delete(self, r):
        rc, out = r
        self.log("✓ 已取消，以后不会自动跑了。" if rc == 0
                 else "✗ 没能取消：%s" % (out.strip() or rc),
                 "good" if rc == 0 else "bad")
        self.refresh()

    def do_logs(self):
        d = runlog.runs_dir(LOG_DIR)
        if not os.path.isdir(d):
            # 升级前的日志（monthly_年月.log）还在上一层，没跑过新版就开那里。
            d = LOG_DIR
        if not os.path.isdir(d):
            messagebox.showinfo("还没有日志",
                                "还没跑过，所以还没有日志：\n%s" % d)
            return
        if os.name == "nt":
            os.startfile(d)                   # noqa: S606
        else:
            subprocess.Popen(["xdg-open", d])

    def _on_close(self):
        self._save_time(region=self.region)
        self.root.destroy()


def main():
    if os.name != "nt":
        print("这个工具只在 Windows 上有意义（用的是 Windows 任务计划程序）。")
        return 1
    root = tk.Tk()
    try:
        root.call("tk", "scaling", 1.3)
    except tk.TclError:
        pass
    SchedApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
