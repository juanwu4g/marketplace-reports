# -*- coding: utf-8 -*-
"""唯一的运行器：下载某个月的数据，再出那个月的报表。

    python run_monthly.py                    # 上个月，下载 + 出报表
    python run_monthly.py --month 2026-07    # 指定月份（补做）
    python run_monthly.py --step report      # 只重出报表，不重新下载
    python run_monthly.py --step collect     # 只补取欠着的报表
    python run_monthly.py --stores A B       # 只跑这几家（不出全店汇总）
    python run_monthly.py --region CL        # 智利店铺（默认 MX 墨西哥）
    python run_monthly.py --region ALL       # 全部国家依次跑（墨西哥第一）
    python run_monthly.py --dry-run          # 只打印会执行什么

文件名还叫 run_monthly 是因为生产机器上已经登记的计划任务调的是
run_monthly.cmd，改名会让它们全部失效。实际上它现在是**每一次运行**的入口：

    ML报表.cmd           -> ml_app.py -> run_monthly.py --entrance manual ...
    计划任务（在同一个窗口里设置） -> run_monthly.cmd --entrance scheduled
    命令行                          -> run_monthly.py（--entrance 默认 cli）

三条路只差 --entrance 这一个参数，它只影响日志里"入口"那一列。锁、日志、
摘要、告警、下载失败后还出不出报表、超时，全部在这里，只有一份。

以前不是这样：界面自己起 run_batch / run_reports，默认月份时**不带** --month；
定时任务永远带。run_batch 带不带 --month 走的是两条不同的下载路线（Ventas
窗口、进账单页的方式、下载目录后缀都不同），所以"手动跑没事、定时跑出事"
这类问题根本没法在界面上复现。现在这里**永远**显式传 --month。

为什么是子进程而不是 import
---------------------------
和 transform.py 同一个理由：下载要驱动浏览器跑几个小时，出报表要吃 pandas。
任何一边崩了都不该把另一边连坐。子进程最坏也就是返回一个非零退出码。

例外是文件开头那句 `from run_reports import prev_month` —— 它是**故意**放在
最前面的：出报表排在几小时的下载之后，等跑到那一步才发现 pandas 装坏了太
晚了。这句 import 失败就立刻退出，那时候什么都还没做。

下载失败了还出不出报表
----------------------
看下载那步的退出码：

* 0（全成）、1（有店铺要关注 / 还欠报表 / 清洗部分失败）→ **出**。缺的会体现
  在报表的【⑩ 数据缺口与待办】里，有总比没有强。
* 2/3/4（紫鸟错误或一家店都没下到 / 白名单 / 配置）→ **不出**。这时候出报表
  会拿旧目录的数据生成一份看起来正常的报表，比没有报表更危险。

自己的时限
----------
计划任务的"运行超过 6 小时则停止"会把**整条链**杀掉 —— 报表不出，告警也不
写，第二天什么痕迹都没有。所以下载这一步自己掐表（--max-hours，默认 5）：
到点就结束下载子进程（连同它起的紫鸟和浏览器），**接着照样出报表、写告警**，
最后以退出码 7 结束。计划任务那个 6 小时只剩兜底的作用。

中止
----
任何一个窗口都能叫停任何一轮：往 logs/runs/<run_id>.stop 写个文件，这里每秒
看一次，看到就杀掉子进程树、把状态记成 aborted、放锁、以退出码 6 结束。

退出码
------
    0  全部成功
    1  完成，但有需要关注的项（个别店铺失败，或报表有校验未过）
    2  硬失败：紫鸟客户端出错 / 没有一家店下到数据 / 有报表没生成
    3  店铺不在白名单，或 --stores 里有不属于 --region 的店
    4  配置错误
    5  另一轮还在跑（锁没释放）
    6  被人中止
    7  下载超过时限被掐断（报表已按现有数据出）

区域（国家）
------------
--region MX|CL，默认 MX。一轮只跑一个区域：两国共用一个紫鸟客户端，锁是全局
的，墨西哥和智利不能同时跑。

* MX：子进程命令和区域出现之前**一字不差**（测试钉死了）。生产上的计划任务、
  告警文件名 `！报表异常_<年月>.txt`、全店汇总表名都不变。
* 其他区域：run_batch **永远**带显式的 --stores（这个区域的全部店铺，或者
  --stores 给的子集）。run_batch 不带店铺时走的是 config 里墨西哥那一批，
  智利一轮要是漏传了店铺，就会去下载墨西哥的 12 家店 —— 不能让这种事靠
  "别忘了传"来保证。run_reports 带 --region，汇总表叫 CL_ML_全店汇总_<月>.xlsx，
  告警文件叫 `！报表异常_CL_<年月>.txt`。
* --stores 里有不属于这个区域的店：退出码 3，什么都不跑。

全部国家：--region ALL
----------------------
一个进程、一把锁、一份运行日志，按 store_config.regions() 的顺序（墨西哥第一）
把每个有店铺的国家**依次**跑一遍。每个国家的子进程命令和单独跑它时一字不差
（就是同一个 build_plan），各出各的全店汇总表、各写各的告警文件。

* 日志里每个国家一段，段首一行醒目的分隔；这一段里的每一行正文前面都带
  `[MX] ` / `[CL] ` 这样的国家代码 —— 只转发中间十几行也分得清是哪个国家。
  不在任何国家里的行（文件头、最后的总结）不带。
* 一个国家下载硬失败、配置有问题，只影响它自己：记下、告警，接着跑下一个。
  被人中止则整轮停下，后面的国家不跑。
* 没有店铺的国家跳过，日志里记一笔。
* 退出码取最坏的那个：0 < 1 < 7 < 2 < 3 < 4（和单个国家时"7 让位给更糟的
  报表失败"同一个排法）；中止仍是 6。
* 不能和 --stores 一起用（店铺子集只在一个国家里有意义）：退出码 3。
* 下载时限 --max-hours 按国家各算各的，和单独跑时一样。

以前每个国家一个计划任务，两个任务时间挨得近就会撞上这把全局锁、后到的被
跳过（退出码 5）。一个 --region ALL 的任务从根上没有这个问题。
"""
import argparse
import datetime
import getpass
import io
import json
import os
import platform
import queue
import re
import signal
import subprocess
import sys
import threading
import time
import traceback

from core import console  # noqa: F401  中文输出编码保护
from core import runlog
# 见上：故意的前置 import，用来在下载之前就验证报表侧能不能加载。
from platforms.mercadolibre.run_reports import prev_month, OUT_DIR

from core.paths import ROOT

# 目前只有一个平台。第二个平台接进来时，这里换成"按店铺所属平台选包"——
# 见 docs/ARCHITECTURE.md 的 "Extension points"。
PLATFORM_PKG = "platforms.mercadolibre"
LOG_DIR = os.path.join(ROOT, "logs")
LOCK = os.path.join(LOG_DIR, "monthly.lock")

# 锁超过这个时长就认为是上次被强杀留下的残骸，直接接管。计划任务的超时是
# 6 小时，这里留足余量。
STALE_HOURS = 8

# 下载这一步最多跑多久，见文件头「自己的时限」。必须小于 win_tasks 里的
# TIME_LIMIT（6 小时），差出来的那一小时留给出报表和写告警。
MAX_HOURS = 5.0

# --region ALL：依次跑全部国家。
ALL = "ALL"

# 子进程在跑的时候，多久看一次中止请求和时限。
POLL = 0.5
HEARTBEAT = 20


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


GUI_SETTINGS = os.path.join(ROOT, "gui_settings.json")


def resolve_out_dir(cli=None):
    """报表输出目录，返回 (目录, 来源)。

    优先级：命令行 --out-dir > 主界面里选的 > 默认目录。

    以前这里只认命令行，缺省就是写死的 downloads\\data\\reports\\financial。
    而定时任务调用本脚本时**不带这个参数** —— 于是在主界面把输出目录指到
    共享盘的人，定时任务出的报表却落进了默认目录，去平常的地方看什么都没有，
    看起来就像"只下载、没出报表"。告警文件也跟着写错了地方。

    直接读 gui_settings.json 而不 import gui：那会把 tkinter 拖进一个无界面的
    计划任务里。统一窗口从这里 import 同一个函数，规则只有一份。
    """
    if cli:
        return cli, "命令行 --out-dir"
    try:
        with io.open(GUI_SETTINGS, encoding="utf-8") as f:
            d = (json.load(f) or {}).get("out_dir")
        if d and str(d).strip():
            return str(d).strip(), "主界面里选的，存在 gui_settings.json"
    except Exception:
        pass
    return OUT_DIR, "默认"


def config():
    """读 config.json 里的 monthly 段。读不到就当没配，不影响主流程。"""
    try:
        with io.open(os.path.join(ROOT, "config.json"), encoding="utf-8") as f:
            return json.load(f).get("monthly") or {}
    except Exception:
        return {}


def region_stores(region):
    """非墨西哥区域的"全部店铺"。墨西哥不走这里（见文件头「区域」）。"""
    from core import store_config
    return list(store_config.stores_for_region(region))


def wrong_region(region, stores):
    """--stores 里不属于 region 的店：[(店名, 它的区域)]。"""
    from core import store_config
    return [(s, store_config.region_of(s)) for s in stores or []
            if store_config.region_of(s) != region]


def region_label(region):
    from core import store_config
    if region == ALL:
        return "全部国家"
    return "%s（%s）" % (store_config.region_name(region), region)


def region_lineup():
    """--region ALL 依次跑哪些国家。返回 (要跑的, 跳过的)。

    要跑的：[{"region", "stores", "problem"}]，顺序是 store_config.regions()
    （墨西哥第一）。problem 和 check_region 的一样是 (退出码, 一句话)：这个国家
    的 batch 写坏了 / 有店不在白名单，它自己不跑、告警，别的国家照跑。
    跳过的：store_config.REGION_NAMES 里一家店都没有的国家（记一笔就好）。

    config.json 本身读不了时抛 StoreConfigError —— 那就哪个国家都定不了。
    """
    from core import store_config
    order = list(store_config.regions())
    order += [r for r in store_config.REGION_NAMES if r not in order]
    run, skipped = [], []
    for r in order:
        try:
            names = list(store_config.stores_for_region(r))
        except store_config.StoreNotAllowed as e:
            run.append({"region": r, "stores": [],
                        "problem": (3, "店铺不在白名单：%s" % e)})
            continue
        except store_config.StoreConfigError as e:
            run.append({"region": r, "stores": [],
                        "problem": (4, "config.json 有问题，定不了%s有哪些店铺：%s"
                                    % (region_label(r), e))})
            continue
        if names:
            run.append({"region": r, "stores": names, "problem": None})
        else:
            skipped.append(r)
    return run, skipped


# 退出码的轻重，--region ALL 取最坏的那个。7（下载超时、报表照出）比 1 重，
# 但让位给 2–4 —— 和单个国家里 final = worst if worst >= 2 else 7 一致。
SEVERITY = (0, 1, 7, 2, 3, 4)


def worst_code(codes):
    """几个国家的退出码合成一个。认不出来的码按 2 算。"""
    worst = 0
    for c in codes:
        c = c if c in SEVERITY else 2
        if SEVERITY.index(c) > SEVERITY.index(worst):
            worst = c
    return worst


class RegionLog(object):
    """--region ALL 时给一个国家用的日志：每行正文前面加 `[CL] `。

    和 runlog.RunLog 同样的接口（write/info/warn/error/path），run_child、
    run_steps、conclude 拿到哪个都一样用。"""

    def __init__(self, log, region):
        self.log = log
        self.region = region
        self.path = log.path

    def write(self, step, level, message, store=runlog.NO_STORE):
        self.log.write(step, level, runlog.tag_region(self.region, message), store)

    def info(self, step, message, store=runlog.NO_STORE):
        self.write(step, runlog.INFO, message, store)

    def warn(self, step, message, store=runlog.NO_STORE):
        self.write(step, runlog.WARN, message, store)

    def error(self, step, message, store=runlog.NO_STORE):
        self.write(step, runlog.ERROR, message, store)


def alert_name(month, region="MX"):
    """告警文件名。墨西哥的名字生产上有人（和脚本）在认，不能改。"""
    ym = month.replace("-", "")
    return ("！报表异常_%s.txt" % ym if region == "MX"
            else "！报表异常_%s_%s.txt" % (region, ym))


def check_region(region, stores):
    """这一轮的店铺和区域对不对得上。返回 (区域全部店铺或 None, 问题或 None)。

    问题是 (退出码, 一句话)：--stores 里有别的区域的店 -> 3（和"不在白名单"
    同一类：店铺名单不对）；非墨西哥区域一家店都没配、或 config 读不了 -> 4。

    墨西哥不给 --stores 时这里什么都不读，和区域出现之前一样。墨西哥给了
    --stores 但 config 读不了：不在这里拦，交给 run_batch 按原来的方式报 4。
    """
    from core import store_config
    try:
        if stores:
            wrong = wrong_region(region, stores)
            if wrong:
                return None, (3, "这些店铺不属于%s，不能放在这一轮里：%s。"
                                 "一轮只跑一个国家；换个区域再跑，或者检查 "
                                 "config.json 的 stores.regions。"
                              % (region_label(region), "、".join(
                                  "%s（%s）" % (s, region_label(r))
                                  for s, r in wrong)))
            return None, None
        if region == "MX":
            return None, None
        names = region_stores(region)
        if not names:
            return None, (4, "config.json 里没有%s的店铺（stores.batches.%s 为空，"
                             "stores.regions 里也没有标成 %s 的店）。"
                          % (region_label(region), region, region))
        return names, None
    except store_config.StoreConfigError as e:
        if region == "MX":
            return None, None
        return None, (4, "config.json 有问题，定不了%s有哪些店铺：%s"
                      % (region_label(region), e))
    except store_config.StoreNotAllowed as e:
        return None, (3, "店铺不在白名单：%s" % e)


def known_stores():
    """白名单里的店名，只用来帮日志认出"这一行说的是哪家店"。读不到就算了。"""
    try:
        from core import store_config
        return list(store_config.batch_stores())
    except Exception:
        return []


# ---------------------------------------------------------------- 命令行

def build_plan(step, month, stores, out_dir, py=None, reports_all=False,
               region="MX", all_stores=None):
    """这一轮要依次跑哪些子进程：[{"key", "label", "cmd"}]。

    纯函数，入口（手动 / 定时 / 命令行）**不是**它的参数 —— 这就是"两个入口
    跑的是同一个过程"的字面意思。

    * run_batch 永远带显式的 --month（原因见文件头）。
    * 给了店铺子集，两个子进程都带 --stores，并且 run_reports 加
      --no-combined：汇总表的文件名写死是 MX_ML_全店汇总_<月>.xlsx，跑 2 家店
      也叫这个名字，会用一份"2 家店的汇总"**覆盖掉**真正 12 家店的那份，而
      文件名还写着"全店"。这条规则以前只有界面知道。
    * reports_all：只**下载**这几家店，报表却给**全部**店铺重出、并重出汇总
      表。给"有几家店要人登录，登录完补跑"用：别的店用它们最近一批数据，
      补跑的店用刚下的，汇总表才是完整的 12 家。
    * region 不是 MX 时，all_stores 是这个区域的全部店铺：没给 stores 就把它
      当成下载的店铺清单显式传下去（见文件头「区域」），run_reports 加
      --region。一家都没有直接报错，绝不退回"不带店铺"。MX 什么都不加。
    """
    py = py or sys.executable
    sub = (["--stores"] + list(stores)) if stores else []
    dsub = sub
    if region != "MX" and not stores:
        if not all_stores:
            raise ValueError("区域 %s 没有店铺，不能不带 --stores 起 run_batch"
                             % region)
        dsub = ["--stores"] + list(all_stores)
    reg = ["--region", region] if region != "MX" else []
    # 子进程按包里的模块起（-m），cwd 是仓库根目录（见 launch 里的 Popen）。
    batch = [py, "-u", "-m", PLATFORM_PKG + ".run_batch"]
    dl = batch + ["--month", month] + dsub
    co = batch + ["--collect-only", "--month", month] + dsub
    rsub = [] if reports_all else sub
    rp = ([py, "-u", "-m", PLATFORM_PKG + ".run_reports", "--month", month]
          + reg + rsub + (["--no-combined"] if rsub else [])
          + ["--out-dir", out_dir])
    plan = []
    if step in ("both", "download"):
        plan.append({"key": "download", "label": "下载 %s 的数据" % month,
                     "cmd": dl})
    if step == "collect":
        plan.append({"key": "collect", "label": "补取欠着的报表", "cmd": co})
    if step in ("both", "report"):
        plan.append({"key": "report", "label": "生成 %s 的报表%s"
                     % (month, "（全部店铺）" if reports_all and stores else ""),
                     "cmd": rp})
    return plan


def show(cmd):
    """命令行写成能直接贴进终端的样子（带空格的参数加引号）。"""
    return " ".join(('"%s"' % c) if (" " in c or not c) else c
                    for c in (str(x).replace("\r", "").replace("\n", "\\n")
                              for x in cmd))


# ---------------------------------------------------------------- 锁

def alive(pid):
    """这个进程还在吗。判断不了就返回 True（宁可多等，不要抢锁）。"""
    return runlog.pid_alive(pid)


def acquire(say=None):
    """防重入。手动跑和计划任务撞上时，两个进程同时驱动同一个紫鸟客户端，
    后启动的那个会把前一个的浏览器连同会话一起重启掉。

    锁里记了 PID，进程没了就立刻接管。只看文件年龄是不够的：跑批中途被
    强杀（关窗口、任务超时、断电）时锁会留下，而下一次定时运行往往就在几
    分钟后 —— 实测过一次 10:01 被杀、10:10 的定时任务拿到退出码 5，
    再按 STALE_HOURS 还要空等 8 小时。

    say(级别, 文字) 用来把经过写进日志；不给就 print。
    """
    def tell(level, text):
        if say:
            say(level, text)
        else:
            print(text)

    lock_dir = os.path.dirname(LOCK)
    if not os.path.isdir(lock_dir):
        os.makedirs(lock_dir)
    body = "pid=%d 起于 %s" % (os.getpid(), now())
    # 两个入口现在共用这把锁，手点和定时任务可能在同一秒起来。O_EXCL 保证
    # "看见没有锁"和"建锁"是一步；先查后写的话两边都会以为自己拿到了。
    for _ in range(2):
        try:
            fd = os.open(LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except OSError:
            fd = None
        if fd is not None:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(body)
            return True
        try:
            age = (datetime.datetime.now()
                   - datetime.datetime.fromtimestamp(os.path.getmtime(LOCK)))
            with io.open(LOCK, encoding="utf-8") as f:
                who = f.read().strip()
        except Exception:
            continue                         # 刚好被对方删了：再抢一次
        m = re.search(r"pid=(\d+)", who)
        pid = int(m.group(1)) if m else 0
        if not alive(pid):
            tell(runlog.INFO, "[提示] 锁是 %s 留下的，但那个进程已经不在了"
                              "（上次被中断），本次接管。" % (who or "上一次运行"))
        elif age < datetime.timedelta(hours=STALE_HOURS):
            tell(runlog.WARN, "另一轮还在跑（%s），本次跳过。" % (who or "?"))
            tell(runlog.INFO, "确认没有在跑的话，删掉 %s 再试。" % LOCK)
            return False
        else:
            tell(runlog.WARN, "[警告] 锁已存在 %.1f 小时且进程仍在，超过 %d 小时"
                              "判为异常，接管。"
                 % (age.total_seconds() / 3600.0, STALE_HOURS))
        try:
            os.remove(LOCK)
        except OSError:
            pass
    tell(runlog.WARN, "锁文件 %s 抢不到，本次跳过。" % LOCK)
    return False


def release():
    try:
        os.remove(LOCK)
    except OSError:
        pass


# ---------------------------------------------------------------- 执行

def child_env():
    env = os.environ.copy()
    # 店名和提示文案都是中文；子进程的 stdout 是管道，默认会按 cp936 编码
    # 写入并抛 UnicodeEncodeError（GUIDE.md 6.4）。
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def kill_tree(p):
    """结束子进程**连同它起的所有进程**。永不抛异常。

    只杀 run_batch 本身是不够的：它起的紫鸟客户端、chromedriver 和浏览器会
    留在那儿，下一轮一上来就撞上一个半死的会话。

    Windows：taskkill /T 顺着父子关系杀整棵树。
    POSIX：子进程是用 start_new_session 起的，自成一个进程组，杀组。
    """
    if p.poll() is not None:
        return
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/T", "/F", "/PID", str(p.pid)],
                           capture_output=True, timeout=30)
        else:
            os.killpg(p.pid, signal.SIGTERM)
            try:
                p.wait(timeout=5)
            except subprocess.TimeoutExpired:
                os.killpg(p.pid, signal.SIGKILL)
    except Exception:
        pass
    try:
        p.wait(timeout=10)
    except Exception:
        # 上面那招没成（taskkill 不在 PATH、权限不够）：至少把它本身杀掉。
        try:
            p.kill()
            p.wait(timeout=10)
        except Exception:
            pass


def _pump(stream, q):
    """子线程：把子进程的输出一行行塞进队列。主线程才有空看表和中止请求。"""
    try:
        for raw in iter(stream.readline, b""):
            q.put(raw)
    except Exception:
        pass
    q.put(None)


def run_child(log, key, cmd, run_id, deadline=None, stores=None,
              log_dir=None):
    """跑一个子进程，把它的输出逐行整理进日志。

    返回 dict：code（退出码）、seconds、ended（"exit" 正常结束 / "aborted"
    被中止 / "deadline" 超时被掐）、markers（机器可读行）、progress。
    deadline 是 time.time() 意义上的截止时刻，None 表示不限时。
    """
    cls = runlog.LineClassifier(key, stores)
    began = time.time()
    kw = {"start_new_session": True} if os.name != "nt" else {}
    p = subprocess.Popen(cmd, cwd=ROOT, stdin=subprocess.DEVNULL,
                         stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                         env=child_env(), **kw)
    q = queue.Queue()
    threading.Thread(target=_pump, args=(p.stdout, q), daemon=True).start()

    ended, eof, exited_at = "exit", False, None
    last_check = last_beat = 0.0

    def emit(raw):
        got = cls.feed(raw.decode("utf-8", "replace").rstrip("\r\n"))
        if got:
            step, store, level, msg = got
            log.write(step, level, msg, store)

    try:
        while True:
            try:
                raw = q.get(timeout=POLL)
                if raw is None:
                    eof = True
                else:
                    emit(raw)
            except queue.Empty:
                pass
            t = time.time()
            if p.poll() is not None:
                if eof:
                    break
                # 进程没了但管道还没到头：它起的某个孙进程还攥着管道。等几秒
                # 把尾巴读完，不能无限等 —— 紫鸟客户端是会一直活着的。
                exited_at = exited_at or t
                if t - exited_at > 3:
                    break
            if t - last_check >= POLL:
                last_check = t
                if runlog.stop_requested(run_id, log_dir):
                    ended = "aborted"
                    break
                if deadline is not None and t > deadline:
                    ended = "deadline"
                    break
                if t - last_beat >= HEARTBEAT:
                    last_beat = t
                    runlog.touch_current(log_dir)
    except KeyboardInterrupt:
        # 控制台里按了 Ctrl+C：当成中止处理，不能把子进程丢在那儿。
        ended = "aborted"

    if ended != "exit":
        kill_tree(p)
        # 被杀之前已经打出来的那几行往往正好说明它卡在哪儿，读完再走。
        stop_at = time.time() + 3
        while time.time() < stop_at:
            try:
                raw = q.get(timeout=0.2)
            except queue.Empty:
                continue
            if raw is None:
                break
            emit(raw)
    try:
        code = p.wait(timeout=10)
    except Exception:
        code = -1
    # 只有读到了头才关管道。没到头说明某个孙进程（实测是 chromedriver）还攥着
    # 写端，读线程正卡在 readline 里、握着这个流的锁 —— 这时 close() 会等那把
    # 锁，永远等不到：2026-10-13 那轮就是这样，下载 15:46 结束，本脚本一直挂到
    # 被人发现，报表那一步始终没开始。不关没有代价：读线程是 daemon，进程退
    # 出时一起走。
    if eof:
        try:
            p.stdout.close()
        except Exception:
            pass
    return {"code": code, "seconds": int(time.time() - began), "ended": ended,
            "markers": cls.markers, "progress": cls.progress}


def norm(code):
    """子进程的退出码收进契约的 0–4。

    被信号杀掉是负数，Windows 上崩溃是 0xC0000005 这样的大数 —— 都算硬失败
    （2），否则它们会原样变成本脚本的退出码，谁也看不懂。
    """
    return code if code in (0, 1, 2, 3, 4) else 2


# ---------------------------------------------------------------- 告警

HARD = {2: "紫鸟客户端出错，或者没有一家店下到数据",
        3: "有店铺不在 config.json 的白名单里",
        4: "config.json 读不了或格式不对"}


# 这两种原因是"平台还没做完"，不是哪里坏了：只提示，不告警。
WAITING = ("STILL_GENERATING", "PLATFORM_DELAYED")


def human_needed(path):
    """这一趟下载之后需要人处理的店（download_diag.human_actions）。读不到汇总返回 []。"""
    doc = runlog.read_json(path) if path else None
    if not doc:
        return []
    from core import download_diag
    return download_diag.human_actions(doc.get("stores"))


def human_line(need):
    """告警文件里的那一句。"""
    stores = []
    for a in need:
        if a["store"] not in stores:
            stores.append(a["store"])
    return ("%d 家店铺需要人工处理：%s。处理完双击 ML报表.cmd，在「需要人工处理」"
            "那一页点「重跑这些店铺并重出全部报表」。"
            % (len(stores), "；".join(
                "%s（%s）" % (s, "、".join(sorted(set(
                    a["title"] for a in need if a["store"] == s))))
                for s in stores)))


def report_human(log, step, need):
    """日志里单独一段，标成错误级：滚过去也能在「级别：错误」里筛出来。"""
    log.error(step, "需要人工处理的店铺（脚本自己解决不了）：")
    for a in need:
        log.error(step, "  %s —— %s：%s%s" % (
            a["store"], a["title"], a["how"],
            "" if a["rerun"] else "（处理之前重跑没用）"), store=a["store"])


def describe_batch(path):
    """读 run_batch 写的批次汇总，说清楚下载那步**到底**出了什么事。

    返回 (urgent, attention)：urgent 要写进告警文件，attention 只进日志。
    读不到汇总（旧版 run_batch 不打标记行、或者它没跑到写汇总那一步）时两个
    都是空，由调用方退回一句笼统的话。

    * 店铺没下到数据 / 有报表没下成（脚本、页面、账号的问题）→ urgent：报表里
      会缺数据，而且要有人去处理。
    * 只是平台还没把报表生成好（STILL_GENERATING / PLATFORM_DELAYED）→
      attention。这是常态：2026-10-09 那轮 12 家店里 7 家有 MercadoPago 报表
      当晚没生成完。每个月都为它写一个"报表异常"，两个月后那个文件就没人
      看了 —— 和 alert() 里说的是同一个道理。稍后补取即可。
    * 只是清洗有文件没过 → attention：报表读的是原始导出，不受影响。
    """
    urgent, attention = [], []
    doc = runlog.read_json(path) if path else None
    if not doc:
        return urgent, attention
    rows = [r for r in (doc.get("stores") or []) if isinstance(r, dict)]

    def names(rs):
        ns = [str(r.get("store")) for r in rs]
        return "、".join(ns[:8]) + ("…" if len(ns) > 8 else "")

    def missing(r):
        return [m for m in (r.get("missing") or []) if isinstance(m, dict)]

    def harmless(m):
        """缺了这张不会让报表少数据：平台还在生成，或者报表根本不读它。"""
        from core import download_diag
        return m.get("cause") in WAITING or not download_diag.affects_report(m)

    def only_waiting(r):
        """这家店缺的全是不影响报表的：平台还在生成的，或者报表不读的那几种。

        路线报错（errors）只要说的都是这些报表，也一样不算。账单路线的报错是
        一行汇总（"billing: 2 张未下载 —— billing/2026-09/BILL_MP_V3 …"），
        报表名在中间，所以按"提到了"算，不按开头算。"""
        ms = missing(r)
        if not ms or not all(harmless(m) for m in ms):
            return False
        forms = [m.get("form") or "" for m in ms]
        return all(any(f in str(e) for f in forms if f)
                   for e in r.get("errors") or [])

    failed = [r for r in rows if r.get("status") not in ("ok", "partial")]
    waiting = [r for r in rows if r.get("status") == "partial" and only_waiting(r)]
    partial = [r for r in rows if r.get("status") == "partial"
               and not only_waiting(r)]
    if failed:
        urgent.append("%d/%d 家店铺没下到数据：%s。"
                      % (len(failed), len(rows), names(failed)))
        for r in failed[:5]:
            errs = r.get("errors") or []
            if errs:
                urgent.append("    %s：%s" % (r.get("store"), str(errs[0])[:160]))
    if partial:
        urgent.append("%d 家店铺有部分报表没下成：%s。"
                      % (len(partial), names(partial)))
        # 原因码是给程序读的；写给财务看的告警里换成人话，码留在括号里方便
        # 对照日志和各店的 REPORT.md。没有逐张记录（旧版汇总）就退回总数。
        counts = {}
        for r in partial + failed:
            for m in missing(r):
                c = m.get("cause") or "UNKNOWN"
                counts[c] = counts.get(c, 0) + 1
        if not counts and isinstance(doc.get("missing_by_cause"), dict):
            counts = dict(doc["missing_by_cause"])
        for k, v in sorted(counts.items()):
            urgent.append("    缺 %s 张：%s（%s）" % (v, _cause_text(k), k))
    if waiting:
        n = sum(len(missing(r)) for r in waiting)
        attention.append("%d 张不影响报表的文件没下到（平台还在生成，或是报表不读的"
                         "那几种）：%s。需要的话过几个小时在 ML报表.cmd 窗口点「补取欠的"
                         "报表」。" % (n, names(waiting)))
    tr = doc.get("transform")
    if isinstance(tr, dict) and not tr.get("skipped") and tr.get("ok") is False:
        attention.append("清洗有文件没通过，其余已写入（不影响报表，报表读的是"
                         "原始导出）。" if tr.get("partial")
                         else "清洗失败（不影响报表，报表读的是原始导出）。")
    return urgent, attention


def _cause_text(code):
    """原因码 -> 一句人话。download_diag 不依赖 selenium / tkinter，可以放心导入；
    真导不进来（文件不全）也不该让告警写不出去。"""
    try:
        from core import download_diag
        return download_diag.CAUSES[code][0]
    except Exception:
        return "原因见日志"


def alert_path(out_dir, month, region="MX"):
    return os.path.join(os.path.abspath(out_dir), alert_name(month, region))


def alert(month, lines, out_dir, entrance="cli", log_path=None, say=None,
          region="MX"):
    """把失败摆到财务看得见的地方。

    计划任务会把 stdout 丢掉，没人会主动去翻 logs/。报表输出目录是财务每个月
    一定会打开的地方，所以告警文件放那儿。手动跑的也一样写 —— 点完按钮就去
    开会的人，回来看的同样是那个目录。

    **只在真出事的时候写。** 报表侧退出码 1（有校验未过）是常态 —— 桥B
    「账单应付 vs 实际扣款」每家店每个月都不过，那是已知的口径问题。真按退
    出码 1 就告警，财务每个月都会看到一个"报表异常"，两个月后这个文件就彻底
    没人看了。会写这个文件的只有两类：**店铺没下到数据**，和**报表根本没生
    成**。报表本身有校验没过，报表的【⑩ 数据缺口与待办】里写着，不劳告警。
    """
    def tell(level, text):
        if say:
            say(level, text)
        else:
            print(text)

    body = ["%s%s 月报生成异常" % ("" if region == "MX"
                                  else region_label(region) + " ", month),
            "时间：%s" % now(),
            "这一轮是：%s（%s）" % (runlog.ENTRANCE_ZH.get(entrance, entrance),
                                 runlog.ENTRANCE_WHAT.get(entrance, "")), ""]
    body += lines
    body += ["", "完整日志：%s" % (log_path or runlog.runs_dir(LOG_DIR)),
             "也可以双击 ML报表.cmd，在窗口右边的「运行日志」里挑这一轮来看。",
             "处理办法见 README.md 第 12 节《出问题先做这三步》。"]
    text = "\n".join(body)
    try:
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir)
        p = os.path.join(out_dir, alert_name(month, region))
        with io.open(p, "w", encoding="utf-8") as f:
            f.write(text)
        tell(runlog.WARN, "已写出告警文件：%s" % p)
    except Exception as e:
        tell(runlog.WARN, "[警告] 告警文件写失败：%s" % str(e)[:120])

    url = config().get("webhook")
    if not url:
        return
    try:
        import requests
        requests.post(url, json={"msgtype": "text",
                                 "text": {"content": text}}, timeout=15)
        tell(runlog.INFO, "已推送 webhook")
    except Exception as e:
        tell(runlog.WARN, "[警告] webhook 推送失败：%s" % str(e)[:120])


# ---------------------------------------------------------------- 主流程

def parse_args(argv=None):
    from core import store_config
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", default=None, metavar="YYYY-MM",
                    help="目标月份，默认上个月")
    ap.add_argument("--step", choices=["both", "download", "report", "collect"],
                    default="both",
                    help="只跑其中一步。collect = 只补取欠着的报表")
    ap.add_argument("--stores", nargs="+", default=None, metavar="NAME",
                    help="只跑这几家店。给了就不出全店汇总表（避免覆盖真正的那份）")
    ap.add_argument("--region", type=str.upper,
                    choices=list(store_config.REGION_NAMES) + [ALL],
                    default="MX",
                    help="哪个国家的店铺：MX 墨西哥（默认）/ CL 智利 / PE 秘鲁；"
                         "ALL = 全部国家依次跑（墨西哥第一），同一把锁、同一份日志")
    ap.add_argument("--reports-for", choices=["stores", "all"], default="stores",
                    help="和 --stores 一起用：all = 只下载这几家店，报表给全部"
                         "店铺重出（含全店汇总）。给人工处理完补跑用")
    ap.add_argument("--out-dir", default=None,
                    help="报表输出目录。不给就用主界面里选的，再不行用默认目录")
    ap.add_argument("--entrance", choices=list(runlog.ENTRANCES), default="cli",
                    help="谁起的这一轮，只影响日志里的「入口」一列："
                         "manual ML 报表窗口 / scheduled 计划任务 / cli 命令行")
    ap.add_argument("--max-hours", type=float, default=MAX_HOURS, metavar="N",
                    help="下载这一步最多跑 N 小时，到点掐断并接着出报表"
                         "（默认 %g，0 = 不限）" % MAX_HOURS)
    ap.add_argument("--dry-run", action="store_true",
                    help="只打印会执行什么")
    return ap.parse_args(argv)


def header(log, args, month, run_id, plan, out_dir, out_src):
    """日志开头那一段：这一轮是谁起的、要干什么、在什么环境里。

    写得啰嗦是故意的。出问题时被转来的往往只有这一个文件，它得自己回答
    "哪台机器、哪个版本、跑的到底是哪条命令"。
    """
    S = runlog.STEP_START
    e = args.entrance
    log.info(S, "入口　　：%s —— %s" % (runlog.ENTRANCE_ZH[e],
                                    runlog.ENTRANCE_WHAT[e]))
    log.info(S, "运行编号：%s" % run_id)
    log.info(S, "目标月份：%s" % month)
    log.info(S, "区域　　：%s" % region_label(args.region))
    if args.stores and args.reports_for == "all":
        what = "、".join(args.stores) + "（只下载这几家；报表给全部店铺重出，含全店汇总）"
    elif args.stores:
        what = "、".join(args.stores) + "（子集，不出全店汇总）"
    else:
        what = ("全部（config.json 白名单）" if args.region == "MX" else
                "全部（config.json 里%s的店铺）"
                % region_label(args.region))
    log.info(S, "店铺　　：%s" % what)
    log.info(S, "步骤　　：%s" % " → ".join(s["label"] for s in plan))
    for i, s in enumerate(plan, 1):
        log.info(S, "命令 %d　：%s" % (i, show(s["cmd"])))
    _header_tail(log, args, out_dir, out_src,
                 any(s["key"] in ("download", "collect") for s in plan))


def _header_tail(log, args, out_dir, out_src, downloads, per_region=False):
    """文件头的后半段：输出目录、时限、机器、版本。单个国家和全部国家共用。"""
    S = runlog.STEP_START
    log.info(S, "输出目录：%s（%s）" % (os.path.abspath(out_dir), out_src))
    if downloads:
        log.info(S, "下载时限：%s" % ("%g 小时%s，到点掐断下载、接着出报表"
                                  % (args.max_hours,
                                     "（每个国家各算各的）" if per_region else "")
                                  if args.max_hours > 0 else "不限"))
    try:
        user = getpass.getuser()
    except Exception:
        user = "?"
    log.info(S, "机器　　：%s　用户：%s" % (platform.node() or "?", user))
    log.info(S, "Python　：%s（%s）" % (platform.python_version(), sys.executable))
    log.info(S, "代码版本：%s" % (runlog.git_commit(ROOT) or "读不到 .git"))
    log.info(S, "日志文件：%s" % log.path)


def run_steps(log, args, plan, run_id, stores, rec, st, save):
    """依次跑一个国家的子进程，把结果记进 rec、把轻重记进 st。

    rec：steps / urgent / attention / batch_summary / human_actions 这几个键，
         单个国家时就是摘要 doc 本身，全部国家时是 doc["regions"] 里那一项。
    st ：{"worst", "cut", "aborted"}，就地改 —— 中途抛异常时调用方还拿得到
         已经跑出来的结果。
    异常不在这里接，由调用方写日志、告警、放锁。
    """
    steps, bad, urgent = rec["steps"], rec["attention"], rec["urgent"]
    for s in plan:
        key = s["key"]
        step = runlog.LineClassifier.BASE[key]
        limited = key in ("download", "collect") and args.max_hours > 0
        deadline = time.time() + args.max_hours * 3600 if limited else None
        log.info(step, "开始：%s" % s["label"])
        r = run_child(log, key, s["cmd"], run_id, deadline=deadline,
                      stores=stores, log_dir=LOG_DIR)
        code = r["code"]
        steps.append({"step": key, "code": code, "seconds": r["seconds"],
                      "ended": r["ended"]})

        if r["ended"] == "aborted":
            st["aborted"] = True
            log.warn(step, "收到中止请求：已结束「%s」和它起的所有进程"
                           "（用时 %d 秒）。已经下好的文件都还在。"
                     % (s["label"], r["seconds"]))
            bad.append("被人中止于「%s」。" % s["label"])
            return

        if r["ended"] == "deadline":
            st["cut"] = True
            where = ""
            if r["progress"]:
                where = "，当时在第 %d/%d 家（%s）" % r["progress"]
            msg = ("下载超过 %g 小时被掐断%s。报表按**已经下到**的数据出；"
                   "没轮到的店铺可能用的是更早一批数据，或者没有报表 —— "
                   "看报表的【⑩ 数据缺口与待办】。"
                   % (args.max_hours, where))
            log.error(step, msg)
            urgent.append(msg)
            bad.append(msg)
            save()
            continue                      # 接着出报表，见文件头

        log.write(step, runlog.INFO if code == 0 else
                  runlog.WARN if code == 1 else runlog.ERROR,
                  "结束：%s，退出码 %d，用时 %d 分 %02d 秒"
                  % (s["label"], code, r["seconds"] // 60, r["seconds"] % 60))
        code = norm(code)
        st["worst"] = max(st["worst"], code)

        if key in ("download", "collect"):
            summary = r["markers"].get("BATCH_SUMMARY_JSON")
            rec["batch_summary"] = summary
            u, a = describe_batch(summary)
            need = human_needed(summary)
            rec["human_actions"] = need
            if need:
                report_human(log, step, need)
                # 放在告警最前面：这是唯一一类"不找人就永远好不了"的事。
                urgent.append(human_line(need))
            if code >= 2:
                urgent.append("下载失败（退出码 %d：%s），报表未生成。"
                              % (code, HARD.get(code, "")))
                urgent.extend(u)
                bad.extend(urgent)
                bad.extend(a)
                save()
                if len(plan) > 1:
                    log.error(step, "下载是硬失败，后面的出报表不跑了 —— "
                                    "不能拿旧数据出一份看起来正常的报表。")
                return
            if code == 1:
                if not u and not a:
                    # 没有批次汇总可读：只能笼统地说。
                    u = ["下载这一步有需要关注的项（没读到批次汇总，细节"
                         "看日志）。报表已按现有数据生成，缺口见报表的"
                         "【⑩ 数据缺口与待办】。"]
                urgent.extend(u)
                bad.extend(u + a)
        else:
            if code == 1:
                # 报表都写出来了，只是校验或 Ventas 源数据要人核对。不告警。
                bad.append("报表都生成了，但需要核对：有校验未过，或 Ventas "
                           "源数据不完整。看报表【⑩ 数据缺口与待办】。")
            elif code >= 2:
                why = {2: "至少有一份报表没有生成（崩溃 / 超时），或者没有"
                          "一家店有可用的下载目录"}.get(code, HARD.get(code, ""))
                urgent.append("报表步骤失败（退出码 %d）：%s。" % (code, why))
                bad.append(urgent[-1])
        save()


def internal_error(log, rec, st):
    """运行器自己出了没料到的错。写进日志、记进告警，算硬失败。"""
    log.error(runlog.STEP_END, "运行器内部出错：\n" + traceback.format_exc())
    rec["urgent"].append("运行器内部出错，这一轮没跑完。细节在日志最后。")
    rec["attention"].append(rec["urgent"][-1])
    st["worst"] = max(st["worst"], 2)


def conclude(log, args, month, out_dir, region, st, rec, log_path):
    """一个国家（没被中止）跑完之后：算出它的退出码，列出要关注的项，该告警就告警。"""
    E = runlog.STEP_END
    bad, urgent = rec["attention"], rec["urgent"]
    # 时限被掐是 7，除非还有更糟的事（报表也没出来）。
    final = st["worst"] if (st["worst"] >= 2 or not st["cut"]) else 7
    if final == 0:
        log.info(E, "%s 全部成功。" % month)
    else:
        log.write(E, runlog.WARN if final == 1 else runlog.ERROR,
                  ("%s 跑完了，有需要关注的项：" if final == 1
                   else "%s 这一轮出了问题：") % month)
        for b in bad:
            log.write(E, runlog.WARN, "  · %s" % b)
        if urgent:
            alert(month, urgent, out_dir, entrance=args.entrance,
                  log_path=log_path, region=region,
                  say=lambda level, text: log.write(E, level, text))
    return final


def execute(args, plan=None):
    """跑一轮。plan 给测试用：换掉真正的子进程命令，流程本身不变。

    --region ALL 走 execute_all；那边的 plan 是 {区域: 计划}。"""
    if args.region == ALL:
        return execute_all(args, plan)
    month = args.month or prev_month()
    out_dir, out_src = resolve_out_dir(args.out_dir)
    region = args.region
    all_stores, problem = check_region(region, args.stores)
    if problem:
        real = []
    else:
        real = build_plan(args.step, month, args.stores, out_dir,
                          reports_all=args.reports_for == "all",
                          region=region, all_stores=all_stores)
    plan = real if plan is None else plan

    if args.dry_run:
        print("入口　　：%s" % runlog.ENTRANCE_ZH[args.entrance])
        print("区域　　：%s" % region_label(region))
        print("目标月份：%s" % month)
        if problem:
            print("\n[错误] %s" % problem[1])
            return problem[0]
        print("日志目录：%s" % runlog.runs_dir(LOG_DIR))
        print("输出目录：%s（%s）" % (os.path.abspath(out_dir), out_src))
        print("\n[dry-run] 会依次执行：")
        for s in plan:
            print("  " + show(s["cmd"]))
        return 0

    run_id = runlog.make_run_id(args.entrance, log_dir=LOG_DIR)
    paths = runlog.run_paths(run_id, LOG_DIR)
    log = runlog.RunLog(run_id, args.entrance, LOG_DIR)
    S, E = runlog.STEP_START, runlog.STEP_END

    # attention（bad）：写进日志和摘要，供人翻查
    # urgent        ：还要写成告警文件推到财务眼前。判断标准见 alert() 的说明。
    steps, bad, urgent = [], [], []
    doc = {"run_id": run_id, "entrance": args.entrance, "pid": os.getpid(),
           "region": region, "month": month, "step": args.step, "stores": args.stores or "all",
           "began": now(), "ended": None, "status": "running",
           "steps": steps, "exit": None, "urgent": urgent, "attention": bad,
           "log": paths["log"], "out_dir": os.path.abspath(out_dir),
           "out_dir_source": out_src, "max_hours": args.max_hours,
           "commands": [{"step": s["key"], "cmd": s["cmd"]} for s in plan],
           "batch_summary": None, "human_actions": []}

    def save(**kw):
        """摘要在开跑时、每一步之后、以及每一条退出路径上都落盘。

        以前只在最后写一次，于是两种情况都查不出来：下载硬失败走的是早退，
        直接跳过了写摘要那段；被计划任务超时强杀时进程根本走不到最后。
        现在被杀的那次会留下 status=running、而且缺某一步的记录 —— 事后一看
        就知道是断在哪儿，而不是像从来没跑过。
        """
        doc.update(kw)
        try:
            runlog.write_json(paths["json"], doc)
        except Exception as e:
            log.warn(E, "[警告] 摘要写不进去：%s" % str(e)[:100])

    def finish(status, code):
        save(status=status, exit=code, ended=now())
        text, _ = runlog.exit_meaning(code)
        if code == 2:
            text = runlog.exit2_text(doc) or text   # 说清楚是哪一步坏了
        level = (runlog.INFO if code == 0 else
                 runlog.WARN if code in (1, 5, 6) else runlog.ERROR)
        log.write(E, level, "本轮结束：退出码 %d —— %s" % (code, text))
        log.close()
        return code

    header(log, args, month, run_id, plan, out_dir, out_src)

    if problem:
        # 在拿锁之前就停：什么都没跑，也不碰 current.json。照样告警 —— 定时
        # 任务撞上这个时没人在看，config 里智利那一批是空的就得有人知道。
        log.error(S, problem[1])
        urgent.append(problem[1])
        bad.append(problem[1])
        alert(month, urgent, out_dir, entrance=args.entrance,
              log_path=paths["log"], region=region,
              say=lambda level, text: log.write(E, level, text))
        return finish("failed", problem[0])

    if not acquire(lambda level, text: log.write(S, level, text)):
        # 没拿到锁：不碰 current.json —— 那是正在跑的那一轮的。
        return finish("skipped", 5)

    runlog.clear_stop(run_id, LOG_DIR)
    runlog.write_current({"run_id": run_id, "pid": os.getpid(),
                          "entrance": args.entrance, "log": paths["log"],
                          "started": doc["began"], "month": month,
                          "region": region,
                          "step": args.step}, LOG_DIR)
    save()
    stores = args.stores or all_stores or known_stores()
    st = {"worst": 0, "cut": False, "aborted": False}
    try:
        run_steps(log, args, plan, run_id, stores, doc, st, save)
    except Exception:
        internal_error(log, doc, st)
    finally:
        release()
        runlog.clear_current(run_id, LOG_DIR)
        runlog.clear_stop(run_id, LOG_DIR)

    if st["aborted"]:
        # 人为中止不告警：停的人自己知道。
        return finish("aborted", 6)
    final = conclude(log, args, month, out_dir, region, st, doc, paths["log"])
    return finish("done" if final < 2 else "failed", final)


# ---------------------------------------------------------------- 全部国家

def region_banner(i, n, region, stores):
    """每个国家那一段的段首。宽一点、带上第几个，滚动的时候一眼看得见。"""
    return "━━━━━━━━ 第 %d/%d 个国家：%s，%d 家店铺 ━━━━━━━━" % (
        i, n, region_label(region), len(stores))


def execute_all(args, plans=None):
    """--region ALL：全部国家依次跑，见文件头「全部国家」。

    plans 给测试用：{区域: 计划}，换掉那个国家真正的子进程命令。没给的国家
    照常用 build_plan —— 和单独跑它时同一个函数、同样的参数。
    """
    month = args.month or prev_month()
    out_dir, out_src = resolve_out_dir(args.out_dir)
    problem, lineup, skipped = None, [], []
    if args.stores:
        problem = (3, "选了全部国家（--region ALL）就不能再指定店铺：店铺子集只在"
                      "一个国家里有意义。只跑几家店请选具体的国家。")
    else:
        from core import store_config
        try:
            lineup, skipped = region_lineup()
        except store_config.StoreConfigError as e:
            problem = (4, "config.json 有问题，定不了每个国家有哪些店铺：%s" % e)
        if not problem and not lineup:
            problem = (4, "config.json 里一个国家的店铺都没有（stores.allowed / "
                          "batch / batches 都是空的）。")
    for item in lineup:
        item["plan"] = []
        if item["problem"]:
            continue
        if plans is not None and item["region"] in plans:
            item["plan"] = plans[item["region"]]
        else:
            item["plan"] = build_plan(args.step, month, None, out_dir,
                                      region=item["region"],
                                      all_stores=item["stores"])

    if args.dry_run:
        print("入口　　：%s" % runlog.ENTRANCE_ZH[args.entrance])
        print("区域　　：全部国家，依次跑")
        print("目标月份：%s" % month)
        if problem:
            print("\n[错误] %s" % problem[1])
            return problem[0]
        print("日志目录：%s" % runlog.runs_dir(LOG_DIR))
        print("输出目录：%s（%s）" % (os.path.abspath(out_dir), out_src))
        for i, item in enumerate(lineup, 1):
            print("\n[dry-run] 第 %d/%d 个国家：%s，%d 家店铺"
                  % (i, len(lineup), region_label(item["region"]),
                     len(item["stores"])))
            if item["problem"]:
                print("  [错误] %s（这个国家不跑，别的国家照跑）" % item["problem"][1])
            for s in item["plan"]:
                print("  " + show(s["cmd"]))
        for r in skipped:
            print("\n[dry-run] %s：config.json 里没有店铺，跳过" % region_label(r))
        return 0

    run_id = runlog.make_run_id(args.entrance, log_dir=LOG_DIR)
    paths = runlog.run_paths(run_id, LOG_DIR)
    log = runlog.RunLog(run_id, args.entrance, LOG_DIR)
    S, E = runlog.STEP_START, runlog.STEP_END

    regions_doc = []
    doc = {"run_id": run_id, "entrance": args.entrance, "pid": os.getpid(),
           "region": ALL, "month": month, "step": args.step, "stores": "all",
           "began": now(), "ended": None, "status": "running",
           "steps": [], "exit": None, "urgent": [], "attention": [],
           "log": paths["log"], "out_dir": os.path.abspath(out_dir),
           "out_dir_source": out_src, "max_hours": args.max_hours,
           "commands": [{"region": it["region"], "step": s["key"], "cmd": s["cmd"]}
                        for it in lineup for s in it["plan"]],
           "regions": regions_doc, "skipped_regions": skipped,
           "current_region": None,
           "batch_summary": None, "human_actions": []}

    extra = []                  # 不属于哪个国家的问题（config 读不了、运行器出错）

    def save(**kw):
        doc.update(kw)
        # 平铺的几项（steps / urgent / attention）从各国那里汇总，旧的读法
        # （exit2_text、界面）照样能用；每项都标上国家。
        doc["steps"] = [dict(x, region=r["region"]) for r in regions_doc
                        for x in r["steps"]]
        doc["urgent"] = extra + [runlog.tag_region(r["region"], x)
                                 for r in regions_doc for x in r["urgent"]]
        doc["attention"] = extra + [runlog.tag_region(r["region"], x)
                                    for r in regions_doc for x in r["attention"]]
        doc["human_actions"] = [dict(a, region=r["region"]) for r in regions_doc
                                for a in r.get("human_actions") or []]
        try:
            runlog.write_json(paths["json"], doc)
        except Exception as e:
            log.warn(E, "[警告] 摘要写不进去：%s" % str(e)[:100])

    def finish(status, code):
        save(status=status, exit=code, ended=now(), current_region=None)
        text, _ = runlog.exit_meaning(code)
        if code == 2:
            text = runlog.exit2_text(doc) or text
        level = (runlog.INFO if code == 0 else
                 runlog.WARN if code in (1, 5, 6) else runlog.ERROR)
        log.write(E, level, "本轮结束：退出码 %d —— %s" % (code, text))
        log.close()
        return code

    # ---- 文件头
    e = args.entrance
    log.info(S, "入口　　：%s —— %s" % (runlog.ENTRANCE_ZH[e], runlog.ENTRANCE_WHAT[e]))
    log.info(S, "运行编号：%s" % run_id)
    log.info(S, "目标月份：%s" % month)
    log.info(S, "区域　　：全部国家，依次跑：%s" % (
        " → ".join(region_label(it["region"]) for it in lineup) or "（无）"))
    for it in lineup:
        log.info(S, "店铺　　：%s %s" % (
            region_label(it["region"]),
            "全部 %d 家：%s" % (len(it["stores"]), "、".join(it["stores"]))
            if it["stores"] else "（定不了，见下）"))
    for r in skipped:
        log.info(S, "店铺　　：%s config.json 里没有店铺，跳过" % region_label(r))
    if lineup and lineup[0]["plan"]:
        log.info(S, "步骤　　：每个国家 %s" % " → ".join(
            s["label"] for s in lineup[0]["plan"]))
    for it in lineup:
        for i, s in enumerate(it["plan"], 1):
            log.info(S, "命令 %s-%d：%s" % (it["region"], i, show(s["cmd"])))
    _header_tail(log, args, out_dir, out_src,
                 any(s["key"] in ("download", "collect")
                     for it in lineup for s in it["plan"]), per_region=True)

    if problem:
        log.error(S, problem[1])
        extra.append(problem[1])
        # 哪个国家都定不了：写进墨西哥那份告警（大家看惯了的那个文件）。
        alert(month, [problem[1]], out_dir, entrance=args.entrance,
              log_path=paths["log"], region="MX",
              say=lambda level, text: log.write(E, level, text))
        return finish("failed", problem[0])

    if not acquire(lambda level, text: log.write(S, level, text)):
        return finish("skipped", 5)

    runlog.clear_stop(run_id, LOG_DIR)
    current = {"run_id": run_id, "pid": os.getpid(), "entrance": args.entrance,
               "log": paths["log"], "started": doc["began"], "month": month,
               "region": ALL, "regions": [it["region"] for it in lineup],
               "step": args.step}
    runlog.write_current(current, LOG_DIR)
    save()
    aborted = False
    n = len(lineup)
    try:
        for i, it in enumerate(lineup, 1):
            r = it["region"]
            rlog = RegionLog(log, r)
            rec = {"region": r, "stores": it["stores"], "status": "running",
                   "exit": None, "began": now(), "ended": None, "steps": [],
                   "urgent": [], "attention": [], "batch_summary": None,
                   "human_actions": [], "alert": None}
            regions_doc.append(rec)
            rlog.info(S, region_banner(i, n, r, it["stores"]))
            runlog.write_current(dict(current, current_region=r,
                                      region_index=i, region_count=n), LOG_DIR)
            save(current_region=r)

            if it["problem"]:
                code, msg = it["problem"]
                rlog.error(S, msg + "（这个国家不跑，别的国家照跑）")
                rec["urgent"].append(msg)
                rec["attention"].append(msg)
                alert(month, rec["urgent"], out_dir, entrance=args.entrance,
                      log_path=paths["log"], region=r,
                      say=lambda level, text, _l=rlog: _l.write(E, level, text))
                rec.update(status="failed", exit=code, ended=now(),
                           alert=alert_path(out_dir, month, r))
                save()
                continue

            st = {"worst": 0, "cut": False, "aborted": False}
            try:
                run_steps(rlog, args, it["plan"], run_id, it["stores"], rec, st, save)
            except Exception:
                internal_error(rlog, rec, st)
            if st["aborted"]:
                aborted = True
                rec.update(status="aborted", exit=6, ended=now())
                save()
                break
            final = conclude(rlog, args, month, out_dir, r, st, rec, paths["log"])
            rec.update(status="done" if final < 2 else "failed", exit=final,
                       ended=now(), alert=alert_path(out_dir, month, r)
                       if final and rec["urgent"] else None)
            text, _ = runlog.exit_meaning(final)
            rlog.write(E, runlog.INFO if final == 0 else
                       runlog.WARN if final == 1 else runlog.ERROR,
                       "%s这一段结束：退出码 %d —— %s"
                       % (region_label(r), final, text))
            save()
    except Exception:
        # 国家之间的那点衔接出了错（不是哪个国家里面的）：记下，照样收尾。
        log.error(E, "运行器内部出错：\n" + traceback.format_exc())
        extra.append("运行器内部出错，这一轮没跑完。细节在日志最后。")
    finally:
        release()
        runlog.clear_current(run_id, LOG_DIR)
        runlog.clear_stop(run_id, LOG_DIR)

    # ---- 总结：每个国家一行
    done = [rec["region"] for rec in regions_doc]
    log.info(E, "全部国家的结果：")
    for rec in regions_doc:
        code = rec.get("exit")
        text, _ = runlog.exit_meaning(code)
        log.write(E, runlog.INFO if code == 0 else
                  runlog.WARN if code in (1, 6) else runlog.ERROR,
                  "  %s：退出码 %s —— %s" % (region_label(rec["region"]), code, text))
    for it in lineup:
        if it["region"] not in done:
            log.warn(E, "  %s：没跑（前面被中止了）" % region_label(it["region"]))
    for r in skipped:
        log.info(E, "  %s：没有店铺，跳过" % region_label(r))

    if aborted:
        return finish("aborted", 6)
    final = worst_code([rec.get("exit") for rec in regions_doc]
                       + ([2] if extra else []))
    if final == 0:
        log.info(E, "%s 全部国家都成功。" % month)
    return finish("done" if final < 2 else "failed", final)


def main(argv=None):
    return execute(parse_args(argv))


if __name__ == "__main__":
    sys.exit(main())
