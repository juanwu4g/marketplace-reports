# -*- coding: utf-8 -*-
"""运行日志的统一格式，以及读它的全部非界面逻辑。

「月度报表」窗口、计划任务和命令行，跑的是同一个 run_monthly.py；它每跑
一次，就在 logs/runs/ 下写一组文件：

    logs/runs/<run_id>.log     给人看的日志，本模块定义的固定六列
    logs/runs/<run_id>.json    摘要（每步退出码、用时、告警内容、命令行）
    logs/runs/<run_id>.stop    有人点了"中止"时才出现，跑完就删
    logs/current.json          正在跑的那一轮；跑完即删

run_id 形如 20261013_140501_scheduled —— 时间加入口，纯 ASCII。

每一行都长这样，六列，用 " | " 隔开：

    2026-10-13 14:05:01 | 定时 | 下载 | BOCINA_TA02 | 警告 | 没有出现新行……

    时间 | 入口（手动/定时/命令行） | 步骤 | 店铺或 - | 级别 | 正文

为什么入口写在**每一行**而不是只写在文件头
------------------------------------------
出问题时被转发的往往是日志中间的十几行。那十几行必须自己说得清"这是半夜
定时跑的那轮，还是下午有人手点的那轮" —— 这两轮以前走的是两条不同的下载
路线，排查时第一个问题永远是这个。

为什么这个模块不碰 tkinter
--------------------------
run_monthly.py（无界面，计划任务里跑）和两个窗口都要 import 它。界面那一半
在 log_view.py，那里只剩 Tk 胶水；找运行记录、跟读文件、解析行、把退出码翻
成人话都在这里，所以能在没有 tkinter 的机器上测。

子进程的输出不改
----------------
run_batch / run_reports 里有五百来个 print，是自由文本。这里不去改它们，而是
由 LineClassifier 逐行"看图说话"：现在是哪一步、哪家店、算什么级别。认不
出来的行一律按「信息」原样写进去 —— 分类错了顶多颜色不对，内容不会丢。
"""
import datetime
import io
import json
import os
import re
import subprocess
import sys
import time

from core.paths import ROOT
LOG_DIR = os.path.join(ROOT, "logs")

ENTRANCES = ("manual", "scheduled", "cli")
ENTRANCE_ZH = {"manual": "手动", "scheduled": "定时", "cli": "命令行"}
ENTRANCE_WHAT = {
    "manual": "有人在「月度报表」窗口（月度报表.cmd）里点按钮跑的",
    "scheduled": "Windows 计划任务拉起的（到点自动跑，或在窗口的「每月自动运行」"
                 "里点了「现在就跑一次」/「试跑一次」）",
    "cli": "有人直接敲命令（或双击 run_monthly.cmd）跑的",
}

STEP_START, STEP_DOWNLOAD, STEP_COLLECT = "启动", "下载", "补收"
STEP_CLEAN, STEP_REPORT, STEP_END = "清洗", "报表", "收尾"
STEPS = (STEP_START, STEP_DOWNLOAD, STEP_COLLECT, STEP_CLEAN, STEP_REPORT,
         STEP_END)

INFO, WARN, ERROR = "信息", "警告", "错误"
LEVELS = (INFO, WARN, ERROR)
LEVEL_RANK = {INFO: 0, WARN: 1, ERROR: 2}

SEP = " | "
NO_STORE = "-"
COMBINED = "全店汇总"          # run_reports 给汇总表起的名字，当作一家"店"来筛

# 契约里约定的机器可读行：只给程序看，不进人看的日志。
MARKERS = ("BATCH_SUMMARY_JSON", "REPORT_CHECKS_JSON")

# 多久没有心跳就认为那一轮已经不在了。PID 会被系统复用：断电重启后
# current.json 还在，里面那个 PID 可能恰好分给了别的程序，只看 PID 的话
# 两个窗口的按钮会永远锁着。
HEARTBEAT_STALE = 180
# 心跳断了但进程还在：算"没有响应"，而不是"没在跑"。超过这么久就不再信那个
# PID 了 —— 断电重启之后同一个号很可能已经是别的程序。
STALL_GIVEUP = 8 * 3600
# 点了中止之后等这么久；运行器还没自己停下，就由窗口直接结束它。
STOP_GRACE = 20


# ════════════════════════════════════════════════════════════ 路径

def runs_dir(log_dir=None):
    return os.path.join(log_dir or LOG_DIR, "runs")


def current_path(log_dir=None):
    return os.path.join(log_dir or LOG_DIR, "current.json")


def run_paths(run_id, log_dir=None):
    d = runs_dir(log_dir)
    return {"log": os.path.join(d, run_id + ".log"),
            "json": os.path.join(d, run_id + ".json"),
            "stop": os.path.join(d, run_id + ".stop")}


def make_run_id(entrance, now=None, log_dir=None):
    """<YYYYMMDD_HHMMSS>_<入口>。同一秒撞名就加 _2、_3（只在测试里见过）。"""
    if entrance not in ENTRANCES:
        raise ValueError("unknown entrance: %r" % (entrance,))
    now = now or datetime.datetime.now()
    base = "%s_%s" % (now.strftime("%Y%m%d_%H%M%S"), entrance)
    rid, n = base, 1
    while os.path.exists(run_paths(rid, log_dir)["log"]):
        n += 1
        rid = "%s_%d" % (base, n)
    return rid


_RUN_ID_RE = re.compile(r"^(\d{8})_(\d{6})_(manual|scheduled|cli)(?:_\d+)?$")


def parse_run_id(run_id):
    """拆出 (开始时间, 入口)。不是本模块起的名字就返回 None。"""
    m = _RUN_ID_RE.match(run_id or "")
    if not m:
        return None
    try:
        when = datetime.datetime.strptime(m.group(1) + m.group(2),
                                          "%Y%m%d%H%M%S")
    except ValueError:
        return None
    return when, m.group(3)


# ════════════════════════════════════════════════════════════ 行格式

def stamp(now=None):
    return (now or datetime.datetime.now()).strftime("%Y-%m-%d %H:%M:%S")


def format_line(when, entrance, step, store, level, message):
    """拼出一行。when 可以是 datetime 或已经格式化好的字符串。"""
    if not isinstance(when, str):
        when = stamp(when)
    return SEP.join([when, ENTRANCE_ZH.get(entrance, entrance), step,
                     store or NO_STORE, level, message])


_TS_RE = re.compile(r"^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d$")


def parse_line(text):
    """把一行拆回六个字段；不是这个格式的返回 None。

    正文里可以有 " | "（比如"路线：all | 白名单：12 家"），所以只切前五刀。
    """
    text = (text or "").rstrip("\r\n")
    parts = text.split(SEP, 5)
    if len(parts) != 6 or not _TS_RE.match(parts[0]):
        return None
    return {"time": parts[0], "entrance": parts[1], "step": parts[2],
            "store": parts[3], "level": parts[4], "message": parts[5],
            "raw": text}


def as_record(text):
    """给查看器用：解析不了的行也要显示，按「信息」原样给。"""
    rec = parse_line(text)
    if rec is None:
        text = (text or "").rstrip("\r\n")
        rec = {"time": "", "entrance": "", "step": "", "store": NO_STORE,
               "level": INFO, "message": text, "raw": text}
    rec["region"], rec["body"] = split_region(rec["message"])
    return rec


def matches(rec, store=None, min_level=None, region=None, text=None):
    """这一行过不过筛选。store=None 不限店铺；min_level 是最低级别；
    region 只看某个国家（--region ALL 那种一轮多国的日志里才有意义）；
    text 是要找的字（不分大小写，店名和正文里找）。"""
    if store and rec.get("store") != store:
        return False
    if min_level and (LEVEL_RANK.get(rec.get("level"), 0)
                      < LEVEL_RANK.get(min_level, 0)):
        return False
    if region and rec.get("region") != region:
        return False
    if text:
        t = text.strip().lower()
        if t and t not in (rec.get("message") or "").lower() \
                and t not in (rec.get("store") or "").lower():
            return False
    return True


# ════════════════════════════════════════════════════════════ 一轮多国

# --region ALL 时，一个国家那一段里每行正文前面带 "[CL] "。只认真有的国家代码：
# 子进程自己打的 "[OK] ..." 不能被当成一个叫 OK 的国家。
_REGION_TAG_RE = re.compile(r"^\[([A-Z]{2,3})\] ")


def _region_codes():
    try:
        from core import store_config             # 不带 selenium / tkinter
        return set(store_config.REGION_NAMES)
    except Exception:
        return {"MX", "CL", "PE"}


def tag_region(region, message):
    """每一行正文前面加上 "[CL] "。多行的话每行都加（RunLog 会按行拆开写）。"""
    lines = str(message).splitlines() or [""]
    return "\n".join("[%s] %s" % (region, x) for x in lines)


def split_region(message):
    """"[CL] 正文" -> ("CL", "正文")；不带国家的行 -> (None, 原文)。"""
    m = _REGION_TAG_RE.match(message or "")
    if m and m.group(1) in _region_codes():
        return m.group(1), message[m.end():]
    return None, message


# ════════════════════════════════════════════════════════════ 逐行分类

_DECOR_RE = re.compile(r"^[=#\-─━—_*~\s]{4,}$")
_MARKER_RE = re.compile(r"^\[(%s)\]\s?(.*)$" % "|".join(MARKERS))
_STORE_HEAD_RE = re.compile(r"^# \[(\d+)/(\d+)\]\s+(\S.*?)\s*$")
_COLLECT_STORE_RE = re.compile(r"^--- 补收：(.+?)（")
_REPORT_STORE_RE = re.compile(r"^单店报表：(\S.*?)\s*$")

_ERROR_MARKS = ("[失败]", "[错误]", "[中止]", "[error]",
                "Traceback (most recent call last)",
                "配置错误：", "店铺不在白名单：", "紫鸟客户端错误：")
_ERROR_RES = (re.compile(r"\bERROR\b"),
              # 异常的最后一行：ValueError: xxx / selenium...TimeoutException: xxx
              re.compile(r"^[\w.]*(Error|Exception)\b[^:：]{0,40}[:：]"),
              # 批次汇总表里的状态列
              re.compile(r"^\S+\s+(failed|not-allowed)\b"))
# 带"失败"二字但其实不是出错的说法。
_NOT_FAILURE = ("部分文件失败", "[日志写入失败]", "失败即拒绝")

_WARN_MARKS = ("[警告]", "[待取]", "[跳过]", "[剔除]", "[warn]", "⚠", "✗",
               "部分文件失败", "[日志写入失败]", "有校验未过", "需要关注",
               "still generating", "仍在生成", "超时", "数据不全")
_WARN_RES = (re.compile(r"\bWARNING\b"),
             re.compile(r"\b\w*Warning\b"),
             re.compile(r"仍欠 [1-9]"),
             re.compile(r"未通过 [1-9]"),
             re.compile(r"^\S+\s+partial\b"))


class LineClassifier(object):
    """喂一行子进程输出，得到 (步骤, 店铺, 级别, 正文)，或者 None（这行不记）。

    纯逻辑，不做任何 IO。步骤先由"现在跑的是哪个子进程"决定，再被输出里的
    横幅细化（run_batch 一个进程里先后是 下载 → 补收 → 清洗）；店铺靠
    `# [3/12] 店名`、`--- 补收：店名（…）`、`单店报表：店名` 这些横幅跟踪。

    不记的行有三种：空行、纯装饰线（===== / #####）、契约里的机器可读行。
    机器可读行的内容留在 self.markers 里给调用方取。
    """

    BASE = {"download": STEP_DOWNLOAD, "collect": STEP_COLLECT,
            "report": STEP_REPORT}

    def __init__(self, child, stores=None):
        self.base = self.BASE[child]
        self.step = self.base
        self.store = NO_STORE
        self.known = set(stores or [])
        self.known.add(COMBINED)
        self.markers = {}
        self.progress = None          # (第几家, 共几家, 店名)，下载阶段才有
        self._tb_indent = None        # 正在一段 Traceback 里时，它的缩进

    # -- 横幅：更新步骤 / 店铺 -------------------------------------------
    def _banner(self, s):
        m = _STORE_HEAD_RE.match(s)
        if m:
            self.step = self.base
            self.store = m.group(3)
            self.known.add(self.store)
            self.progress = (int(m.group(1)), int(m.group(2)), self.store)
            return
        m = _COLLECT_STORE_RE.match(s)
        if m:
            self.step, self.store = STEP_COLLECT, m.group(1)
            self.known.add(self.store)
            return
        m = _REPORT_STORE_RE.match(s)
        if m:
            self.store = m.group(1)
            self.known.add(self.store)
            return
        if (s.startswith("补收阶段") or s.startswith(">>> 补收阶段")
                or s.startswith(">>> 已跳过补收阶段")
                or s.startswith("仅补收模式")):
            self.step, self.store = STEP_COLLECT, NO_STORE
        elif s.startswith(">>> 清洗") or s.startswith(">>> 已跳过清洗"):
            self.step, self.store = STEP_CLEAN, NO_STORE
        elif s.startswith(">>> 对账簿"):
            self.store = NO_STORE
        elif s.startswith("批次汇总") or s.startswith("批次 "):
            self.step, self.store = self.base, NO_STORE
        elif s.startswith("汇总报表："):
            self.store = COMBINED
        elif (s.startswith("报表生成结果") or s.startswith("月度财务报表")
              or s.startswith("未参与（")):
            self.store = NO_STORE

    # -- 级别 -------------------------------------------------------------
    def _level(self, line, s):
        indent = len(line) - len(line.lstrip())
        if self._tb_indent is not None:
            # Traceback 里面的行都比它缩进得深；第一行不更深的就是异常本身。
            if indent <= self._tb_indent:
                self._tb_indent = None
            return ERROR
        if "Traceback (most recent call last)" in s:
            self._tb_indent = indent
            return ERROR
        if any(k in s for k in _ERROR_MARKS) or any(r.search(s) for r in _ERROR_RES):
            return ERROR
        if "失败" in s and not any(k in s for k in _NOT_FAILURE):
            return ERROR
        if any(k in s for k in _WARN_MARKS) or any(r.search(s) for r in _WARN_RES):
            return WARN
        return INFO

    def _store_for(self, s):
        """没有"当前店铺"时（汇总表、目录清单），看行首是不是一个认识的店名。"""
        if self.store != NO_STORE:
            return self.store
        for tok in s.split()[:3]:
            tok = tok.rstrip("：:,，")
            if tok in self.known:
                return tok
        return NO_STORE

    def feed(self, raw):
        line = (raw or "").rstrip()
        s = line.strip()
        if not s:
            return None
        m = _MARKER_RE.match(s)
        if m:
            self.markers[m.group(1)] = m.group(2).strip()
            return None
        if _DECOR_RE.match(s):
            return None
        self._banner(s)
        level = self._level(line, s)
        return self.step, self._store_for(s), level, line


# ════════════════════════════════════════════════════════════ 写日志

class RunLog(object):
    """一轮运行的日志文件。每行写完就 flush。

    每行都 flush：被计划任务超时强杀、或者断电时，缓冲区里的内容全丢，而那
    恰恰是最需要看的几行。也因为这样，查看器才能做到"实时"。
    """

    def __init__(self, run_id, entrance, log_dir=None, echo=True):
        self.run_id = run_id
        self.entrance = entrance
        self.echo = echo
        self.path = run_paths(run_id, log_dir)["log"]
        d = os.path.dirname(self.path)
        if not os.path.isdir(d):
            os.makedirs(d)
        # newline="\n"：Windows 上也只写 \n，查看器按字节数跟读时不用猜换行。
        self._f = io.open(self.path, "a", encoding="utf-8", newline="\n")
        # 有人盯着控制台（双击 run_monthly.cmd）时，屏幕上打短一点的行；
        # 文件里永远是完整的六列。重定向到文件 / 管道时原样打。
        try:
            self._tty = bool(sys.stdout.isatty())
        except Exception:
            self._tty = False

    def write(self, step, level, message, store=NO_STORE):
        for part in (str(message).splitlines() or [""]):
            line = format_line(datetime.datetime.now(), self.entrance, step,
                               store, level, part)
            try:
                self._f.write(line + "\n")
                self._f.flush()
            except Exception:
                pass                  # 日志写不进去不该让跑批停下来
            if self.echo:
                try:
                    sys.stdout.write((console_line(line) if self._tty else line)
                                     + "\n")
                    sys.stdout.flush()
                except Exception:
                    pass              # 没有控制台（窗口拉起的后台进程）

    def info(self, step, message, store=NO_STORE):
        self.write(step, INFO, message, store)

    def warn(self, step, message, store=NO_STORE):
        self.write(step, WARN, message, store)

    def error(self, step, message, store=NO_STORE):
        self.write(step, ERROR, message, store)

    def close(self):
        try:
            self._f.close()
        except Exception:
            pass


def console_line(line):
    """控制台上的一行：`14:05:01  下载  BOCINA_TA02  [警告]  正文`。

    日期和入口在控制台上是多余的（就是现在、就是这个黑框）；级别只在不是
    「信息」时标出来，扫一眼就看得见。"""
    rec = parse_line(line)
    if not rec:
        return line
    parts = [rec["time"][-8:], rec["step"]]
    if rec["store"] != NO_STORE:
        parts.append(rec["store"])
    if rec["level"] != INFO:
        parts.append("[%s]" % rec["level"])
    return "  ".join(parts + [rec["message"]])


def git_commit(root=None):
    """当前代码的 commit（短），读不到返回 None。不调 git，机器上未必装了。"""
    try:
        g = os.path.join(root or ROOT, ".git")
        common = g
        if os.path.isfile(g):                      # worktree：.git 是个指针文件
            with io.open(g, encoding="utf-8") as f:
                g = f.read().split("gitdir:", 1)[1].strip()
            common = g
            c = os.path.join(g, "commondir")
            if os.path.isfile(c):
                with io.open(c, encoding="utf-8") as f:
                    common = os.path.normpath(os.path.join(g, f.read().strip()))
        with io.open(os.path.join(g, "HEAD"), encoding="utf-8") as f:
            head = f.read().strip()
        if not head.startswith("ref:"):
            return head[:7] or None
        ref = head.split(":", 1)[1].strip()
        p = os.path.join(common, *ref.split("/"))
        if os.path.isfile(p):
            with io.open(p, encoding="utf-8") as f:
                return f.read().strip()[:7] or None
        with io.open(os.path.join(common, "packed-refs"), encoding="utf-8") as f:
            for line in f:
                if line.strip().endswith(" " + ref):
                    return line.split()[0][:7]
    except Exception:
        pass
    return None


# ════════════════════════════════════════════════════════════ 小文件读写

def read_json(path):
    """读一个 JSON 文件；不存在、正被改写、内容不完整都返回 None。"""
    try:
        with io.open(path, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else None
    except Exception:
        return None


def write_json(path, doc):
    """先写临时文件再换名，读的人不会读到半截。

    Windows 上别的进程正开着目标文件时 os.replace 会报 PermissionError
    （查看器每秒读一次），所以重试几次，实在不行就直接覆盖写。
    """
    d = os.path.dirname(path)
    if d and not os.path.isdir(d):
        os.makedirs(d)
    text = json.dumps(doc, ensure_ascii=False, indent=2)
    tmp = "%s.%d.tmp" % (path, os.getpid())
    try:
        with io.open(tmp, "w", encoding="utf-8") as f:
            f.write(text)
        for _ in range(5):
            try:
                os.replace(tmp, path)
                return
            except OSError:
                time.sleep(0.05)
    except Exception:
        pass
    try:
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(text)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


# ════════════════════════════════════════════════════════════ 进程在不在

def _win_pid_alive(pid):
    import ctypes
    from ctypes import wintypes
    k = ctypes.WinDLL("kernel32", use_last_error=True)
    k.OpenProcess.restype = wintypes.HANDLE
    k.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    k.GetExitCodeProcess.argtypes = (wintypes.HANDLE,
                                     ctypes.POINTER(wintypes.DWORD))
    k.CloseHandle.argtypes = (wintypes.HANDLE,)
    h = k.OpenProcess(0x1000, False, pid)   # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        err = ctypes.get_last_error()
        if err == 87:                       # ERROR_INVALID_PARAMETER：没这个 PID
            return False
        if err == 5:                        # 拒绝访问：进程在，只是不让看
            return True
        raise OSError(err)
    try:
        code = wintypes.DWORD()
        if not k.GetExitCodeProcess(h, ctypes.byref(code)):
            raise OSError(ctypes.get_last_error())
        return code.value == 259            # STILL_ACTIVE
    finally:
        k.CloseHandle(h)


_TASKLIST_CACHE = {}


def _tasklist_alive(pid):
    # 起一个 tasklist 要几百毫秒，而查看器每秒要问好几次：结果留 3 秒。
    hit = _TASKLIST_CACHE.get(pid)
    if hit and time.time() - hit[0] < 3:
        return hit[1]
    ok = _tasklist_query(pid)
    _TASKLIST_CACHE.clear()
    _TASKLIST_CACHE[pid] = (time.time(), ok)
    return ok


def _tasklist_query(pid):
    try:
        out = subprocess.run(["tasklist", "/FI", "PID eq %d" % pid, "/NH"],
                             capture_output=True, timeout=20,
                             encoding="utf-8", errors="replace").stdout or ""
    except Exception:
        return True
    # 没匹配到时 tasklist 打的是一句本地化的提示，里面不会有这个 PID。
    return str(pid) in out


def pid_alive(pid):
    """这个进程还在吗。判断不了就返回 True（宁可多等，不要抢锁）。

    不能用 os.kill(pid, 0) 探测 Windows 进程 —— 那里 CPython 的 os.kill 走的
    是 TerminateProcess，信号 0 会把目标进程**杀掉**。所以 Windows 上直接问
    内核；问不成再退回 tasklist（慢，要起一个子进程）。
    """
    try:
        pid = int(pid or 0)
    except (TypeError, ValueError):
        pid = 0
    if pid <= 0:
        return True
    if os.name == "nt":
        try:
            return _win_pid_alive(pid)
        except Exception:
            return _tasklist_alive(pid)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except OSError:
        return True
    # 已经死了、只是父进程还没收尸的（僵尸）不算活着：它什么都占不了。
    # 只有 Linux 有 /proc；读不到就按"还在"算。
    try:
        with io.open("/proc/%d/stat" % pid, encoding="ascii", errors="replace") as f:
            if f.read().rsplit(")", 1)[-1].split()[0] == "Z":
                return False
    except Exception:
        pass
    return True


# ════════════════════════════════════════════════════════════ 正在跑的那一轮

def write_current(doc, log_dir=None):
    write_json(current_path(log_dir), doc)


def touch_current(log_dir=None):
    """心跳。运行中每隔几十秒碰一下，见 HEARTBEAT_STALE。"""
    try:
        os.utime(current_path(log_dir), None)
    except OSError:
        pass


def clear_current(run_id=None, log_dir=None):
    """跑完删掉。给了 run_id 就只删自己那份，不动别人的。"""
    p = current_path(log_dir)
    if run_id is not None:
        doc = read_json(p)
        if doc is not None and doc.get("run_id") != run_id:
            return
    try:
        os.remove(p)
    except OSError:
        pass


def active_run(log_dir=None, now=None):
    """现在有没有一轮在跑（不管是哪个入口起的）。有就返回它的 current.json。

    文件在、里面那个 PID 还活着，就算有。心跳也没断是正常情况；PID 活着而心跳
    断了，返回的 dict 里多一个 stalled=True（和 quiet_seconds）：那一轮卡住了。

    卡住的那轮**不能**当成"没在跑"：2026-10-13 实测，运行器挂死之后界面一边
    把按钮放开（心跳断了），一边还显示"正在跑"（进程还在），而"中止"按不动。
    它还占着锁和紫鸟，必须让人看得见、停得掉。
    """
    p = current_path(log_dir)
    doc = read_json(p)
    if not doc or not doc.get("run_id"):
        return None
    try:
        age = (now or time.time()) - os.path.getmtime(p)
    except OSError:
        return None
    if not pid_alive(doc.get("pid")):
        return None
    if age > HEARTBEAT_STALE:
        if age > STALL_GIVEUP:
            return None
        doc = dict(doc, stalled=True, quiet_seconds=int(age))
    return doc


# ════════════════════════════════════════════════════════════ 中止请求

def request_stop(run_id, who="", log_dir=None):
    """请那一轮停下来。它每秒会看一次这个文件，看到就杀掉子进程树并收尾。

    用文件而不是直接杀进程：点"中止"的窗口和起那一轮的往往不是同一个（定时
    跑起来的，在窗口里点「中止这一轮」），手里没有进程句柄；而且让它自己停，锁、
    摘要、日志才收得干净。
    """
    p = run_paths(run_id, log_dir)["stop"]
    d = os.path.dirname(p)
    if not os.path.isdir(d):
        os.makedirs(d)
    with io.open(p, "w", encoding="utf-8") as f:
        f.write("%s %s\n" % (stamp(), who))
    return p


def stop_requested(run_id, log_dir=None):
    return os.path.exists(run_paths(run_id, log_dir)["stop"])


def clear_stop(run_id, log_dir=None):
    try:
        os.remove(run_paths(run_id, log_dir)["stop"])
    except OSError:
        pass


def kill_pid_tree(pid):
    """按 PID 结束一个进程连同它起的所有进程。永不抛异常，返回是否动了手。

    给"运行器自己停不下来"的情况用：手里只有 current.json 里的 PID，没有句柄。
    """
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/T", "/F", "/PID", str(pid)],
                           capture_output=True, timeout=30,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW",
                                                 0x08000000))
        else:
            import signal
            todo, seen = [pid], []
            while todo:                          # 先找全子孙，再从叶子往上杀
                cur = todo.pop()
                seen.append(cur)
                try:
                    out = subprocess.run(["ps", "-o", "pid=", "--ppid", str(cur)],
                                         capture_output=True, timeout=10).stdout
                    todo += [int(x) for x in out.split()]
                except Exception:
                    pass
            for one in reversed(seen):
                try:
                    os.kill(one, signal.SIGKILL)
                except OSError:
                    pass
        return True
    except Exception:
        return False


def force_stop(run_id, who="", log_dir=None):
    """运行器不理中止请求（挂死了）时的最后手段：直接结束它，并替它收尾。

    它自己收不了尾，所以这里补上最要紧的三样：日志里留一行说明是谁、为什么
    强行结束的；摘要改成"已中止"；删掉 current.json 让两个窗口都知道没在跑了。
    锁文件不用管 —— 里面的 PID 已经不在，下一轮会自动接管。
    """
    paths = run_paths(run_id, log_dir)
    summary = read_json(paths["json"]) or {}
    cur = read_json(current_path(log_dir)) or {}
    pid = cur.get("pid") if cur.get("run_id") == run_id else summary.get("pid")
    kill_pid_tree(pid)
    parsed = parse_run_id(run_id)
    entrance = summary.get("entrance") or (parsed[1] if parsed else "cli")
    for _ in range(50):                          # taskkill 返回后进程未必立刻消失
        if not pid_alive(pid):
            break
        time.sleep(0.1)
    if pid_alive(pid):
        # 杀不掉（实测：那一轮是用管理员权限起的，而本窗口不是）。不能假装
        # 它停了 —— 它还占着锁和紫鸟。说清楚，把请求撤掉免得每秒重试一次。
        try:
            with io.open(paths["log"], "a", encoding="utf-8") as f:
                f.write(format_line(datetime.datetime.now(), entrance, STEP_END,
                                    NO_STORE, ERROR,
                                    "想强行结束这一轮，但结束不了进程 %s（多半是"
                                    "权限不够）。请在任务管理器里结束它的"
                                    " python.exe，或重启电脑。" % pid) + "\n")
        except Exception:
            pass
        clear_stop(run_id, log_dir)
        return False
    try:
        with io.open(paths["log"], "a", encoding="utf-8") as f:
            f.write(format_line(datetime.datetime.now(), entrance, STEP_END,
                                NO_STORE, ERROR,
                                "这一轮没有响应中止请求，已被强行结束（%s）。"
                                "已下好的文件保留；没有接着出报表。" % (who or "?"))
                    + "\n")
    except Exception:
        pass
    if summary:
        summary.update(status="aborted", exit=6, ended=stamp(), forced=True)
        try:
            write_json(paths["json"], summary)
        except Exception:
            pass
    clear_current(run_id, log_dir)
    clear_stop(run_id, log_dir)
    return True


def enforce_stop(log_dir=None, now=None):
    """有人点了中止、而运行器过了宽限期还没停（或者它本来就卡住了）：强行结束。

    两个窗口的查看器每秒调一次。正常情况下运行器几秒内自己停下、删掉请求
    文件，这里什么都不做。返回被强行结束的 run_id，没有则 None。
    """
    cur = active_run(log_dir, now)
    if not cur:
        return None
    run_id = cur["run_id"]
    p = run_paths(run_id, log_dir)["stop"]
    try:
        age = (now or time.time()) - os.path.getmtime(p)
    except OSError:
        return None                              # 没人点中止
    if cur.get("stalled") or age > STOP_GRACE:
        who = ""
        try:
            with io.open(p, encoding="utf-8") as f:
                who = f.read().strip()
        except Exception:
            pass
        return run_id if force_stop(run_id, who, log_dir) else None
    return None


# ════════════════════════════════════════════════════════════ 退出码 → 人话

# (说明, 颜色类别)。类别：good / warn / bad。含义以契约为准。
EXIT_MEANING = {
    0: ("全部正常", "good"),
    1: ("跑完了，但有需要关注的项 —— 看日志里标黄标红的行，或报表第 10 页（数据缺口与待办）", "warn"),
    2: ("没跑成：紫鸟客户端出错 / 没有一家店下到数据 / 有报表没生成", "bad"),
    3: ("没跑成：店铺不在 config.json 的白名单里，或不属于所选的国家", "bad"),
    4: ("没跑成：config.json 读不了或格式不对", "bad"),
    5: ("没跑：另一轮还在跑，这次被跳过了", "warn"),
    6: ("被人中止了", "warn"),
    7: ("下载超过时限被掐断，报表是按已下到的数据出的", "bad"),
}


def exit_meaning(code):
    try:
        return EXIT_MEANING[int(code)]
    except (KeyError, TypeError, ValueError):
        return ("异常结束，退出码 %s" % (code,), "bad")


def exit2_text(summary):
    """退出码 2 说的是哪一步坏了。

    2 是"下载硬失败 / 有报表没生成"的合称。2026-10-14 那轮 11 家店都下好了、
    只有 1 家店因为要登录而没有报表，窗口却写着"紫鸟客户端出错 / 没有一家店
    下到数据" —— 两句都不对。按各步骤的退出码说清楚是哪一步。
    """
    if summary.get("region") == "ALL":
        return _exit2_all(summary)
    codes = {st.get("step"): st.get("code") for st in summary.get("steps") or []
             if isinstance(st, dict)}
    dl = codes.get("download", codes.get("collect"))
    rp = codes.get("report")
    if dl is not None and dl >= 2:
        return "下载失败（紫鸟出错，或者没有一家店下到数据），报表没有生成"
    if rp is not None and rp >= 2:
        return ("下载完成，但有报表没生成（有店铺没下到可用的数据，或出报表时"
                "出错）—— 看左边「需要人工处理」和日志末尾")
    return None


def _exit2_all(summary):
    """--region ALL 那一轮的退出码 2：说清楚是哪几个国家、坏在哪一步。"""
    out = []
    for r in summary.get("regions") or []:
        if not isinstance(r, dict) or (r.get("exit") or 0) < 2 or r.get("exit") == 7:
            continue
        text = exit2_text(dict(r, region=r.get("region"))) if r.get("exit") == 2 \
            else exit_meaning(r.get("exit"))[0]
        out.append("%s：%s" % (region_name(r.get("region")),
                              text or exit_meaning(r.get("exit"))[0]))
    return "；".join(out) or None


def region_name(code):
    """MX -> 墨西哥；ALL -> 全部国家。"""
    if code == "ALL":
        return "全部国家"
    try:
        from core import store_config
        return store_config.region_name(code)
    except Exception:
        return str(code or "")


def run_state(run_id, log_dir=None):
    """一轮运行现在是什么状态，翻成界面能直接显示的东西。

    返回 dict：state（running / finished / aborted / skipped / interrupted /
    unknown）、text、kind（run / good / warn / bad）、exit、entrance、month。
    """
    summary = read_json(run_paths(run_id, log_dir)["json"]) or {}
    parsed = parse_run_id(run_id)
    entrance = summary.get("entrance") or (parsed[1] if parsed else "")
    out = {"run_id": run_id, "entrance": entrance,
           "region": summary.get("region") or "MX",
           "month": summary.get("month"), "exit": summary.get("exit"),
           "pid": summary.get("pid"), "summary": summary}
    status = summary.get("status")
    if status == "running":
        cur = active_run(log_dir)
        if cur and cur.get("run_id") == run_id and cur.get("stalled"):
            # 仍然算 running（它占着锁和紫鸟，按钮要继续锁着），但话要说清楚。
            out.update(state="running", stalled=True, kind="bad",
                       text="没有响应：已经 %d 分钟没有任何动静，多半卡住了 —— "
                            "点「中止这一轮」结束它"
                            % (cur.get("quiet_seconds", 0) // 60))
        elif cur and cur.get("run_id") == run_id:
            out.update(state="running", text="正在跑", kind="run", current=cur)
        elif pid_alive(summary.get("pid")) and summary.get("pid"):
            # current.json 还没写出来或刚被删，进程还在：当它在跑。
            out.update(state="running", text="正在跑", kind="run")
        else:
            out.update(state="interrupted", kind="bad",
                       text="中途断了：进程被强行结束（超时、关机或被杀），没有收尾")
        return out
    if status is None:
        out.update(state="unknown", text="没有摘要（旧记录或没写完）", kind="warn")
        return out
    text, kind = exit_meaning(summary.get("exit"))
    if summary.get("exit") == 2:
        text = exit2_text(summary) or text
    state = {"aborted": "aborted", "skipped": "skipped"}.get(status, "finished")
    out.update(state=state, text=text, kind=kind)
    return out


# ════════════════════════════════════════════════════════════ 找运行记录

def list_runs(log_dir=None, limit=60):
    """logs/runs 下的运行记录，最近的在前。每项：run_id / when / entrance。"""
    d = runs_dir(log_dir)
    try:
        names = os.listdir(d)
    except OSError:
        return []
    out = []
    for n in names:
        if not n.endswith(".log"):
            continue
        rid = n[:-4]
        parsed = parse_run_id(rid)
        if parsed:
            out.append({"run_id": rid, "when": parsed[0], "entrance": parsed[1]})
    out.sort(key=lambda r: r["run_id"], reverse=True)
    return out[:limit]


def region_tag(st):
    """"智利 " 这样的前缀，加在月份前面。墨西哥（和区域出现之前的旧记录）是
    空串：两国并存之前的运行记录一行都不变。"""
    r = (st or {}).get("region") or "MX"
    if r == "MX":
        return ""
    return region_name(r) + " "


def run_label(run, log_dir=None):
    """下拉框里的一行：`10-13 14:05  定时  2026-09  全部正常`。

    非墨西哥的那一轮在月份前面带国家：`10-13 14:05  定时  智利 2026-09  …`。"""
    st = run_state(run["run_id"], log_dir)
    return "%s  %s  %s%s  %s" % (
        run["when"].strftime("%m-%d %H:%M"),
        ENTRANCE_ZH.get(run["entrance"], run["entrance"]),
        region_tag(st), st.get("month") or "",
        st["text"].split(" —— ")[0].split("：")[0])


# ════════════════════════════════════════════════════════════ 跟读文件

class Tailer(object):
    """跟着一个还在被追加的文件往下读，只返回**完整**的行。

    每次 poll 都重新打开、读完就关，不长期攥着句柄 —— 写的是另一个进程，
    Windows 上攥着句柄会妨碍它（以及之后删文件、换名）。按字节读、按 \\n
    切，最后半行留到下次：写的人一行还没写完，或者一个汉字的三个字节只到
    了两个，都不会显示成乱码。
    """

    CHUNK = 256 * 1024       # 一次最多读这么多，打开一份大日志时界面不会卡住

    def __init__(self, path):
        self.path = path
        self.offset = 0
        self._rest = b""

    def poll(self):
        """返回 (新读到的行列表, 后面是否还有没读的)。"""
        try:
            size = os.path.getsize(self.path)
        except OSError:
            return [], False
        if size < self.offset:              # 文件被截断或换了一份：从头来
            self.offset, self._rest = 0, b""
        if size == self.offset:
            return [], False
        try:
            with io.open(self.path, "rb") as f:
                f.seek(self.offset)
                data = f.read(self.CHUNK)
        except (OSError, IOError):
            return [], False
        self.offset += len(data)
        buf = self._rest + data
        parts = buf.split(b"\n")
        self._rest = parts.pop()
        lines = [p.decode("utf-8", "replace").rstrip("\r") for p in parts]
        return lines, self.offset < size


class Follower(object):
    """查看器背后的全部逻辑：该看哪一轮、读新行、算状态。不碰 Tk。

    两种模式：
    * 跟随（默认）：有一轮在跑就看它 —— 不管是哪个入口起的，也包括窗口开着
      的时候半夜定时任务自己跑起来的那轮；没有在跑的就看最近一轮。
    * 钉住：用户从列表里挑了一轮过去的，就一直看它，不被新的一轮抢走。

    log_view.py 每隔几百毫秒调一次 tick()，按返回值刷新界面。
    """

    def __init__(self, log_dir=None):
        self.log_dir = log_dir
        self.pinned = None
        self.run_id = None
        self.records = []
        self.stores = []
        self.regions = []             # 日志里出现过的国家代码（一轮多国时才有）
        self._tail = None
        self._state = None
        self._state_at = 0.0

    def follow_current(self):
        self.pinned = None

    def pin(self, run_id):
        self.pinned = run_id

    def _target(self):
        if self.pinned:
            return self.pinned
        cur = active_run(self.log_dir)
        if cur:
            return cur["run_id"]
        if self.run_id and self._state and self._state.get("state") == "running":
            # 刚才还在跑的那轮：先别跳走，等它的摘要落定再说。
            return self.run_id
        runs = list_runs(self.log_dir, limit=1)
        return runs[0]["run_id"] if runs else None

    def _attach(self, run_id):
        self.run_id = run_id
        self.records = []
        self.stores = []
        self.regions = []
        self._state = None
        self._state_at = 0.0
        self._tail = (Tailer(run_paths(run_id, self.log_dir)["log"])
                      if run_id else None)

    def tick(self, now=None):
        """返回 dict：reset（换了一轮，界面要清空）、new（新记录）、
        state（run_state 的结果，可能为 None）、state_changed、more。"""
        now = time.time() if now is None else now
        reset = False
        target = self._target()
        if target != self.run_id:
            self._attach(target)
            reset = True
        new, more = [], False
        if self._tail is not None:
            lines, more = self._tail.poll()
            for text in lines:
                rec = as_record(text)
                new.append(rec)
                s = rec["store"]
                if s and s != NO_STORE and s not in self.stores:
                    self.stores.append(s)
                r = rec.get("region")
                if r and r not in self.regions:
                    self.regions.append(r)
            self.records.extend(new)
        changed = False
        # 状态一秒看一次就够；已经结束的那轮不会再变，就不反复读了。
        settled = self._state is not None and self._state.get("state") not in (
            "running", "unknown")
        if self.run_id and (reset or (not settled
                                      and now - self._state_at >= 1.0)):
            enforce_stop(self.log_dir)
            st = run_state(self.run_id, self.log_dir)
            self._state_at = now
            key = (st.get("state"), st.get("exit"), st.get("stalled"))
            old = self._state and (self._state.get("state"),
                                   self._state.get("exit"),
                                   self._state.get("stalled"))
            changed = key != old
            self._state = st
        return {"reset": reset, "new": new, "more": more,
                "state": self._state, "state_changed": changed or reset}

    def filtered(self, store=None, min_level=None, region=None, text=None):
        return [r for r in self.records
                if matches(r, store, min_level, region, text)]


# ════════════════════════════════════════════════════════════ 查看器的一行

def view_columns(rec, with_region=False):
    """查看器里一行分成哪几列：[时间, (国家,) 步骤, 店铺, 级别, 正文]。

    日志文件里每行都带日期和入口，那是给事后翻文件的人看的；窗口里入口是顶上
    那块色块、日期在「看哪一轮」里，所以只留时分秒。国家只在一轮多国（--region
    ALL）时才有一列，正文前面那个 "[CL] " 挪进这一列。解析不了的行整行放进正文。
    """
    if not rec.get("time"):
        cols = ["", "", "", "", rec.get("raw") or rec.get("message") or ""]
        return cols[:1] + ([""] if with_region else []) + cols[1:]
    store = rec.get("store") or ""
    cols = [rec["time"][-8:], rec.get("step") or "",
            "" if store == NO_STORE else store, rec.get("level") or ""]
    body = rec.get("body") if rec.get("body") is not None else rec.get("message")
    if with_region:
        r = rec.get("region")
        cols.insert(1, region_name(r) if r else "")
    else:
        body = rec.get("message")
    return cols + [body or ""]


# ════════════════════════════════════════════════════════════ 状态栏

def elapsed_text(run_id, now=None):
    """"1:12:05" —— 从运行编号里的开始时间算到现在。算不出来返回 ""。"""
    parsed = parse_run_id(run_id)
    if not parsed:
        return ""
    secs = int(((now or datetime.datetime.now()) - parsed[0]).total_seconds())
    if secs < 0:
        return ""
    return "%d:%02d:%02d" % (secs // 3600, secs % 3600 // 60, secs % 60)


def live_status(st, records=(), now=None, scan=3000):
    """窗口顶上那一句"现在在跑什么"：国家、第几个国家、月份、步骤、店铺（第几家）、
    已经跑了多久。不在跑（或没有记录）返回 None。

    st 是 run_state() 的结果；records 是查看器读到的日志行（as_record 的结果），
    只往回看最后 scan 行 —— 一份几万行的日志每秒从头翻一遍太浪费。
    """
    if not st or st.get("state") != "running":
        return None
    cur = st.get("current") or {}
    summary = st.get("summary") or {}
    region = summary.get("region") or cur.get("region") or "MX"
    parts = []
    if region == "ALL":
        r = cur.get("current_region") or summary.get("current_region")
        idx, n = cur.get("region_index"), cur.get("region_count")
        if r:
            parts.append("%s（%s）%s" % (region_name(r), r,
                                       "，第 %d/%d 个国家" % (idx, n)
                                       if idx and n else ""))
        else:
            parts.append("全部国家")
    else:
        parts.append("%s（%s）" % (region_name(region), region))
    if st.get("month"):
        parts.append(st["month"])
    step = store = prog = None
    for rec in list(records)[-scan:][::-1]:
        if not rec.get("time"):
            continue
        if step is None:
            step, store = rec.get("step"), rec.get("store")
        m = _STORE_HEAD_RE.match((rec.get("body") or rec.get("message") or "").strip())
        if m:
            prog = (int(m.group(1)), int(m.group(2)), m.group(3))
            break
    if step:
        parts.append(step)
    if store and store != NO_STORE:
        parts.append(store + ("（第 %d/%d 家）" % prog[:2]
                              if prog and prog[2] == store else ""))
    took = elapsed_text(st.get("run_id"), now)
    if took:
        parts.append("已运行 " + took)
    return "  ·  ".join(parts)


# ════════════════════════════════════════════════════════════ 起一轮

def human_actions_for_run(run_id, log_dir=None):
    """某一轮下载之后需要人处理的店（见 download_diag.human_actions）。

    读这一轮摘要里记的批次汇总；没有下载这一步（只出报表）的返回 None ——
    和"下载了、没有要人处理的"（返回 []）区分开。
    """
    groups = human_action_groups(run_id, log_dir)
    return groups[0] if groups else None


def _region_batches(summary):
    """一轮摘要里的批次汇总：[(区域, 批次汇总路径, 是不是这个国家的全部店铺)]。

    单个国家的一轮只有一项；--region ALL 的一轮每个国家一项（都是全部店铺）。"""
    if summary.get("region") == "ALL":
        return [(r.get("region") or "MX", r.get("batch_summary"), True)
                for r in summary.get("regions") or [] if isinstance(r, dict)]
    return [(summary.get("region") or "MX", summary.get("batch_summary"),
             summary.get("stores") in ("all", None))]


def human_action_groups(run_id, log_dir=None):
    """某一轮里每个国家下载之后需要人处理的店：[{run_id, month, region,
    stores, actions, full}]。没有下载这一步的国家不出现。"""
    summary = read_json(run_paths(run_id, log_dir)["json"]) or {}
    out = []
    for region, path, full in _region_batches(summary):
        batch = read_json(path) if path else None
        if not batch:
            continue
        from core import download_diag            # 纯 Python，不带 selenium / tkinter
        out.append({"run_id": run_id, "month": summary.get("month"),
                    "region": region, "full": full,
                    "stores": [r.get("store") for r in batch.get("stores") or []
                               if isinstance(r, dict)],
                    "actions": download_diag.human_actions(batch.get("stores"))})
    return out


def pending_human_actions(log_dir=None, look_back=30):
    """现在还有哪些店需要人处理：每家店以**最近一次**下载到它的那一轮为准。

    重跑了一家店（子集）之后，这家店以重跑那轮为准，其余的店仍以之前那一轮
    为准 —— 所以重跑好了的店会从清单里消失，没重跑的不会被漏掉。往回看到
    一轮全店下载（或 look_back 轮）为止。返回 {"month", "actions"}。

    区域：「一轮全店下载」只让**同一个国家**更早的记录过时。以前看到一轮全店
    就停；现在墨西哥跑完一整轮，往前那轮智利的待办不能跟着消失。每条待办带
    region（那一轮的区域），界面补跑时按它传 --region。
    """
    seen, actions, month = set(), [], None
    full = set()                        # 已经见过一轮全店下载的区域
    for run in list_runs(log_dir, limit=look_back):
        if read_json(run_paths(run["run_id"], log_dir)["json"]) is None:
            continue
        # --region ALL 的一轮里每个国家各算一段，和单独跑那个国家一样对待。
        for got in human_action_groups(run["run_id"], log_dir):
            region = got["region"]
            if region in full:
                continue                # 这个国家更晚有过一轮全店：这段过时了
            fresh = [s for s in got["stores"] if s not in seen]
            for a in got["actions"]:
                if a["store"] in fresh:
                    actions.append(dict(a, run_id=got["run_id"], region=region))
                    month = month or got["month"]
            seen.update(got["stores"])
            if got["full"]:
                full.add(region)        # 一轮全店下载：这个国家再往前的都过时了
    return {"month": month, "actions": actions}


def runner_argv(entrance, step="both", month=None, stores=None, out_dir=None,
                reports_for=None, region=None):
    """给 run_monthly.py 的参数。两个窗口都用它拼，区别只有 entrance。

    month / stores / out_dir 不给就不写 —— 由 run_monthly 自己取默认值（上个
    月、全部店铺、主界面里选的输出目录）。定时任务就是这样：只带 --entrance
    和 --region。region 不给也不写，run_monthly 按墨西哥跑（旧任务就是这样）。
    """
    a = ["--entrance", entrance]
    if region:
        a += ["--region", region]
    if step and step != "both":
        a += ["--step", step]
    if month:
        a += ["--month", month]
    if stores:
        a += ["--stores"] + list(stores)
    if out_dir:
        a += ["--out-dir", out_dir]
    if reports_for:
        a += ["--reports-for", reports_for]
    return a


def spawn_detached(cmd, cwd=None, stderr_path=None):
    """起一个不跟着本窗口一起死的后台进程。返回 Popen。

    日志文件才是输出，所以 stdout 直接丢掉。stderr 另存一个小文件：解释器
    在写出第一行日志之前就崩了（比如依赖没装好）的话，那是唯一的线索。

    Windows：CREATE_NO_WINDOW 而不是 DETACHED_PROCESS —— 后者让进程完全没有
    控制台，它再起的每个控制台子进程（run_batch、taskkill、chromedriver）都
    会各弹一个黑窗口；前者给它一个看不见的控制台，子孙共用。
    CREATE_NEW_PROCESS_GROUP 让它不吃到发给本窗口的 Ctrl+C / Ctrl+Break。
    stdin 也接到 DEVNULL：pythonw 起的进程没有有效的标准句柄。
    """
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    err = subprocess.DEVNULL
    fh = None
    if stderr_path:
        try:
            d = os.path.dirname(stderr_path)
            if d and not os.path.isdir(d):
                os.makedirs(d)
            fh = io.open(stderr_path, "wb")
            err = fh
        except Exception:
            fh, err = None, subprocess.DEVNULL
    kw = {}
    if os.name == "nt":
        kw["creationflags"] = (getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
                               | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP",
                                         0x00000200))
    else:
        kw["start_new_session"] = True
    try:
        return subprocess.Popen(cmd, cwd=cwd or ROOT, env=env,
                                stdin=subprocess.DEVNULL,
                                stdout=subprocess.DEVNULL, stderr=err,
                                close_fds=True, **kw)
    finally:
        if fh is not None:
            fh.close()            # 子进程已经拿到自己那份句柄了


def launch_error_path(entrance, log_dir=None):
    return os.path.join(runs_dir(log_dir), "_launch_%s.err" % entrance)
