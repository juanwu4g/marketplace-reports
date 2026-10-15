# -*- coding: utf-8 -*-
"""
调用 report_build.py 生成月度财务报表。

`report_build.py` 是一个独立的报表生成器（不 import 项目里任何模块），本文件是
它与下载流程之间的接缝：负责挑出每家店该用哪个下载目录、拼出命令、跑起来、把结
果讲清楚。它刻意不了解报表内部怎么算。

沿用 transform.py 的两条约定：

* **子进程，不是 import。** pandas / openpyxl 抛异常不能反过来打断调用方；报表
  里的公式重算本来也要外调 LibreOffice。最坏情况只是打印一段错误。
* **绝不抛异常。** 报表没生成不该让已经下载好的文件跟着失败。
"""
import io
import json
import datetime
import glob
import os
import re
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(ROOT, "report_build.py")
DOWNLOADS = os.path.join(ROOT, "downloads")

# 没写出报表时的现场记录放这里，不放报表输出目录：那是财务每个月要打开的文件
# 夹，以前每出一份报表（包括成功的）就往里面的 diagnostics/ 丢一个日志，从不
# 清理，一年下来几百个。logs/ 本来就是放运行记录的地方（.gitignore 已排除）。
DIAG_DIR = os.path.join(ROOT, "logs", "report_diag")

# report_build.py 的退出码。照抄自那边的 EXIT_*（本文件刻意不 import 它，见文
# 件头），两边要一起改。
RC_OK = 0             # 已写出，error 级校验全过
RC_CHECKS = 1         # 已写出，有 error 级校验未过或公式错误
RC_NO_STORE = 2       # 没有一家店可用，没写
RC_STRICT_ABORT = 10  # --strict 中止，没写
RC_CRASH = 70         # 未处理的异常

# build() 结果里的 outcome 取值。前四个表示**文件确实是本次写出来的**；其余
# 全部表示这份报表没有产出。run_reports 的退出码就按这条线划：前者最多是 1
# （要看一眼），后者是 2（报表没出来）。
WRITTEN_OUTCOMES = ("ok", "data_incomplete", "checks_failed", "formula_errors")
OUTCOME_LABELS = {
    "ok": "成功",
    "data_incomplete": "数据不全⚠",        # 已写出；平台声明 Ventas 导出不完整
    "checks_failed": "有校验未过",          # 已写出；有 error 级校验没通过
    "formula_errors": "公式错误",           # 已写出；重算后文件里有公式报错
    "strict_aborted": "严格模式未写出",     # --strict 下校验未过，按要求不写
    "no_usable_store": "无可用数据",        # 每家店的数据都用不了
    "stale": "写入失败",                    # 旧文件还在，本次没写进去
    "crashed": "生成出错",                  # report_build 抛了异常
    "timeout": "超时",
    "not_written": "未写出",                # 其他原因没有产出文件
    "skipped": "跳过",                      # 根本没启动（缺脚本 / 没有店铺）
}

# 报表必需的输入。缺 Ventas 或账单就算不出损益表，这样的目录不能用 —— 光取"最新
# 目录"是不够的：一次只跑了 --only ventas 的运行会留下一个看着很新、其实缺账单的
# 目录（BOCINA_SM 的 20260903_110457 就是）。
# 每项是 (几个可接受的文件名模式, 名称)：智利的销售报表叫 Ventas_CL，秘鲁叫 Ventas_PE。
VENTAS_PATTERNS = ("*Ventas_MX*.xlsx", "*Ventas_CL*.xlsx", "*Ventas_PE*.xlsx")
REQUIRED = (
    (VENTAS_PATTERNS, "Ventas"),
    (("Reporte_Facturacion_MercadoLibre_*.xlsx",), "账单"),
)


# 账单文件名里的月份缩写。与 report_build.FILE_MONTH_TOKENS 同源；这里单独留一份，
# 是因为 financial_report 刻意不 import report_build（两者靠子进程相隔，见文件头）。
FILE_MONTH_TOKENS = {
    "ene": 1, "enero": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
    "set": 9,       # 秘鲁的 Setiembre（实测文件名仍是 Sept，防平台改拼法）
}

# 账单按自然月出的站点（report_build 里 BILL_CYCLE_DAY == 1 的那些）。秘鲁账单页
# 写"Cierre 30/set/2026"，IDG_PE 实测 Sept 账单计费日 09-01..09-30。
CALENDAR_BILL_SITES = ("mx", "pe")


def sales_span(folder):
    """这个目录的销售报表覆盖到多早。返回 pd 可比的 (最早, 最晚) 字符串 YYYY-MM。

    只读文件名解决不了 —— 销售报表的文件名里只有导出日期，没有窗口。所以读
    run_meta.json 里记下的实际窗口；老目录没有这个文件时返回 None，宁可不判断
    也不猜。
    """
    meta_path = os.path.join(folder, "run_meta.json")
    if not os.path.isfile(meta_path):
        return None
    try:
        with io.open(meta_path, encoding="utf-8") as fh:
            meta = json.load(fh) or {}
        n, at = meta.get("sales_window_months"), meta.get("downloaded_at")
        if not n or not at:
            return None
        d = datetime.date(int(at[:4]), int(at[5:7]), int(at[8:10]))
    except Exception:
        return None
    y, m = d.year, d.month - int(n)
    while m <= 0:
        y, m = y - 1, m + 12
    return "%04d-%02d" % (y, m), "%04d-%02d" % (d.year, d.month)


def sales_reaches(folder, month):
    """这个目录的销售数据**确实**含有早于 month 月初的订单吗？

    run_meta.json 给的是请求的窗口，最准，但老目录没有。退而求其次直接翻文件：
    只要里面存在早于目标月月初的订单，就**证明**窗口盖过了这个月 —— 这是证据，
    不是推测。反过来不成立（没有更早的订单也可能只是这家店当月才开张），所以
    这个函数只用来**肯定**，不用来否定。

    读 xlsx 有代价，所以只在账单匹配失败后才调用。
    """
    span = sales_span(folder)
    if span:
        return span[0] <= month <= span[1]
    hits = []
    for pat in VENTAS_PATTERNS + ("*Ventas*.xlsx",):
        hits = sorted(glob.glob(os.path.join(folder, pat)))
        if hits:
            break
    if not hits:
        return False
    p = hits[0]
    try:
        import pandas as pd
        d = pd.read_excel(p, sheet_name=0, header=5, usecols=[1])
        col = d.columns[0]
        first = min((_spanish_year_month(v) for v in d[col] if isinstance(v, str)),
                    default=None)
        return bool(first and first < month)
    except Exception:
        return False


_ES_MONTHS = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5,
              "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9,
              "octubre": 10, "noviembre": 11, "diciembre": 12}


def _spanish_year_month(v):
    """'3 de agosto de 2026 09:24 hs.' -> '2026-08'；认不出返回 None。"""
    m = re.match(r"\s*\d{1,2}\s+de\s+([A-Za-zÁ-úá-ú]+)\s+de\s+(\d{4})", v)
    if not m:
        return None
    mo = _ES_MONTHS.get(m.group(1).lower())
    return "%s-%02d" % (m.group(2), mo) if mo else None


def billing_months(folder):
    """这个目录里有哪几个月的账单。返回 {(年, 月)}。

    账单文件名自带月份（Reporte_Facturacion_MercadoLibre_Ago2026.xlsx），
    不用打开文件就能知道覆盖了哪几期。
    """
    out = set()
    for f in glob.glob(os.path.join(folder, "Reporte_*.xlsx")):
        m = re.search(r"_([A-Za-z]{3,5})(\d{4})\.xlsx$", os.path.basename(f))
        if m and m.group(1).lower() in FILE_MONTH_TOKENS:
            out.add((int(m.group(2)), FILE_MONTH_TOKENS[m.group(1).lower()]))
    return out


def folder_site(folder):
    """下载目录记的站点（run_meta.json 的 "site"）。老目录没有这个字段，按 mx。"""
    try:
        with io.open(os.path.join(folder, "run_meta.json"), encoding="utf-8") as fh:
            return str((json.load(fh) or {}).get("site") or "mx").lower()
    except Exception:
        return "mx"


def bills_needed(folder, want):
    """做 want（(年, 月)）这个会计月要哪几期账单。

    墨西哥、秘鲁账单是自然月，只要当月那一期。智利账单不按自然月切（各店切期日不同，
    如 8 号或 24 号），一个自然月横跨"当月"和"次月"两期账单 —— 两期都要有，
    报表才拿得全本月的费用（report_build.bill_cycles 是同一条规则）。
    """
    y, m = want
    if folder_site(folder) in CALENDAR_BILL_SITES:
        return {want}
    ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
    return {want, (ny, nm)}


def available():
    return os.path.isfile(SCRIPT)


def has_libreoffice():
    """报表里的公式需要 LibreOffice 重算出缓存值，没有就只能跳过。"""
    if os.environ.get("SOFFICE"):
        return True
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.isfile(p):
            return True
    for d in os.environ.get("PATH", "").split(os.pathsep):
        for exe in ("soffice.exe", "soffice"):
            if d and os.path.isfile(os.path.join(d, exe)):
                return True
    return False


def _usable(folder):
    """这个下载目录是否够算一份报表。"""
    import glob
    missing = [label for pats, label in REQUIRED
               if not any(glob.glob(os.path.join(folder, p)) for p in pats)]
    return (not missing), missing


def latest_run(store, base=None, month=None):
    """这家店最近一个**能用**的下载目录。

    返回 (目录, 说明)。没有可用目录时返回 (None, 原因)。
    按时间戳倒序找第一个文件齐全的，而不是无脑取最新 —— 见 REQUIRED 的注释。

    给了 month（"YYYY-MM"）时，**优先**挑账单里确实含那个月的目录。做历史月份
    报表时这一步是必须的：最新那个目录只有最近两期账单，拿它做 7 月报表，
    佣金全是 0，代扣税会被算成佣金的好几倍（实测 28%，正常 9%）。
    找不到含该月账单的目录时退回原来的规则，并在说明里讲清楚。
    """
    want = None
    if month:
        try:
            want = (int(month[:4]), int(month[5:7]))
        except Exception:
            want = None
    base = base or DOWNLOADS
    d = os.path.join(base, store)
    if not os.path.isdir(d):
        return None, "没有下载目录"
    runs = sorted((x for x in os.listdir(d)
                   if os.path.isdir(os.path.join(d, x)) and x[:2] == "20"),
                  reverse=True)
    if not runs:
        return None, "目录下没有任何运行记录"
    usable = []
    skipped = []
    for r in runs:
        folder = os.path.join(d, r)
        ok, missing = _usable(folder)
        if ok:
            usable.append((r, folder))
        else:
            skipped.append(r)

    if want:
        wm = "%04d-%02d" % want
        # 第一优先：确实含该月账单的目录。没有账单，费用侧整块是空的。
        # 智利要两期账单都在（见 bills_needed）；先找两期都在的，找不到再退一步。
        for r, folder in usable:
            need = bills_needed(folder, want)
            if len(need) > 1 and need <= billing_months(folder):
                return folder, "%s（含 %s 账单及次月账单）" % (r, wm)
        for r, folder in usable:
            if want in billing_months(folder):
                return folder, "%s（含 %s 账单）" % (r, wm)
        # 第二优先：账单没有，但销售窗口盖得住该月的目录。这样至少收入侧是对的，
        # 总比退回最新目录（两头都缺）强。必须把"为什么选它、还缺什么"写进说明 ——
        # 选了一个不完美的目录而不告诉人，比直接失败更糟。
        for r, folder in usable:
            if sales_reaches(folder, wm):
                span = sales_span(folder)
                how = ("销售窗口 %s~%s" % span) if span else "销售数据含更早的订单"
                return folder, ("%s（⚠ 无 %s 账单，但%s，盖得住该月："
                                "收入可用，平台费用会缺失，代扣税率会失真）"
                                % (r, wm, how))

    for r, folder in usable:
        note = r if not skipped else "%s（跳过 %d 个不完整的更新目录）" % (r, len(skipped))
        if want:
            note += ("  ⚠ 没有任何目录含 %s 的账单，也没有目录的销售窗口盖得住该月 —— "
                     "这份报表的收入和费用都会缺失，仅供参考" % wm)
        return folder, note
    return None, "%d 个目录都缺文件（最新的缺：%s）" % (
        len(runs), "、".join(_usable(os.path.join(d, runs[0]))[1]))


def classify(r):
    """把一次 build 的结果归成一个 outcome（取值见 OUTCOME_LABELS）。

    单独成函数、只看传进来的 dict，是为了让"什么算写出来了"只有一处定义：
    build() 用它，run_reports 的汇总和退出码也用它，测试可以不起子进程直接喂。

    顺序有讲究 —— 先排除"没产出"，再谈校验：
    * 崩溃排在最前。report_build 可能在写完 xlsx 之后才抛异常（重算、导指标），
      那时文件在、修改时间也新，但这次生成并没有走完，不能算写出来了。
    * `wrote` 为假的一律是没产出，不管 checks_failed 是不是真 —— --strict 中止
      时两者同时成立，以前汇总里会显示成"有校验未过"，配一个并不存在的文件。
    """
    if r.get("skipped"):
        return "skipped"
    rc = r.get("returncode")
    if r.get("timed_out"):
        return "timeout"
    # 报表正被 Excel/WPS 打开时，report_build 是以 PermissionError 崩掉的。这
    # 种情况说"写入失败"比说"生成出错"有用得多 —— 关掉文件重跑就行。
    if r.get("stale") and "PermissionError" in (r.get("output") or ""):
        return "stale"
    if r.get("crashed") or rc == RC_CRASH:
        return "crashed"
    if r.get("stale"):
        return "stale"
    if not r.get("wrote"):
        if rc == RC_STRICT_ABORT:
            return "strict_aborted"
        if rc == RC_NO_STORE:
            return "no_usable_store"
        return "not_written"
    if r.get("formula_errors"):
        return "formula_errors"
    if r.get("checks_failed") or rc == RC_CHECKS:
        return "checks_failed"
    if rc not in (None, RC_OK):
        # 文件是新的，但退出码不是任何一种"写完了"的码。宁可报没产出。
        return "not_written"
    if r.get("data_incomplete"):
        return "data_incomplete"
    return "ok"


def written(r):
    """这份报表本次是否确实写出来了（不论校验过没过）。"""
    return (r.get("outcome") or classify(r)) in WRITTEN_OUTCOMES


def _write_diag(out, month, pairs, rc, secs, output):
    """把一次没产出报表的 build 的完整输出存进 logs/report_diag/。

    返回 (路径, 错误)。只在没写出报表时调用：成功的 build 其输出已经在控制台
    和月度日志里，再存一份只是重复。写不进去不算事 —— 返回错误文字让调用方
    提一句即可，不能因为记日志失败再丢一个异常出去。
    """
    try:
        if not os.path.isdir(DIAG_DIR):
            os.makedirs(DIAG_DIR)
        stem = os.path.splitext(os.path.basename(out))[0]
        path = os.path.join(DIAG_DIR, "%s.%s.log" % (
            stem, datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
        with io.open(path, "w", encoding="utf-8") as fh:
            fh.write("会计月：%s\n输出：%s\nPython：%s\n退出码：%s\n用时：%ds\n" % (
                month, os.path.abspath(out), sys.executable, rc, secs))
            for name, folder in pairs:
                fh.write("输入：%s = %s\n" % (name, os.path.abspath(folder)))
            fh.write("\n" + output)
        return path, None
    except (OSError, IOError) as e:
        return None, str(e)


def build(pairs, month, out, metrics_out=None, strict=False,
          no_recalc=None, timeout=3600, echo=True, site=None):
    """生成一份报表。pairs 是 [(店铺名, 下载目录), ...]。

    一份报表可以覆盖多家店 —— 汇总报表就是把 12 家一起传进来。
    绝不抛异常。返回 dict，调用方该看的是：

        outcome         这次的结局，取值见 OUTCOME_LABELS；由 classify() 得出
        wrote           文件是否**本次**写出（存在且修改时间变新）
        ok              wrote 且 report_build 退出码为 0
        checks          report_build 给出的全部校验（[REPORT_CHECKS_JSON]）
        checks_failed   是否有 error 级校验未过
        data_incomplete 平台是否声明 Ventas 导出不完整（warn 级）
        formula_errors  重算后文件里是否有公式报错
        returncode      report_build 的退出码（见 RC_*）；没跑起来或超时为 None
        diagnostic_log  没写出报表时保存的完整输出的路径，否则为 None

    "写没写出来"和"校验过没过"是两件事：checks_failed 为真时文件多半仍已写出
    （除非 --strict），所以不要用 checks_failed 判断有没有文件，看 outcome。
    """
    if not available():
        return {"ok": False, "skipped": "找不到 report_build.py",
                "outcome": "skipped", "wrote": False,
                "returncode": None, "seconds": 0, "out": out, "output": ""}
    if not pairs:
        return {"ok": False, "skipped": "没有可用的店铺", "returncode": None,
                "outcome": "skipped", "wrote": False,
                "seconds": 0, "out": out, "output": ""}

    if no_recalc is None:
        no_recalc = not has_libreoffice()

    d = os.path.dirname(os.path.abspath(out))
    if d and not os.path.isdir(d):
        os.makedirs(d)

    cmd = [sys.executable, SCRIPT, "--month", month, "--out", out]
    for name, folder in pairs:
        cmd += ["--store", "%s=%s" % (name, folder)]
    if metrics_out:
        cmd += ["--metrics-out", metrics_out]
    if strict:
        cmd.append("--strict")
    if no_recalc:
        cmd.append("--no-recalc")
    if site:
        # 国家（mx / cl / pe）。不给时 report_build 读下载目录 run_meta.json 的 site。
        cmd += ["--site", str(site).lower()]

    env = dict(os.environ)
    # 报表全程中文输出，管道里必须是 UTF-8，否则一句 print 就会 UnicodeEncodeError
    env["PYTHONIOENCODING"] = "utf-8"

    # 记下调用前的修改时间：文件"存在"不等于"本次写成功了"。报表被 Excel 打开时
    # openpyxl 会 PermissionError，旧文件原封不动留在原地 —— 只判断存在就会把
    # 写失败报成成功。
    before_mtime = os.path.getmtime(out) if os.path.isfile(out) else None

    began = time.time()
    timed_out = False
    try:
        p = subprocess.run(cmd, cwd=ROOT, env=env, timeout=timeout,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           encoding="utf-8", errors="replace")
        output, rc = p.stdout or "", p.returncode
    except subprocess.TimeoutExpired as e:
        partial = e.stdout or ""
        if isinstance(partial, bytes):
            partial = partial.decode("utf-8", errors="replace")
        output, rc = partial + "\n[错误] 超时（%d 秒）" % timeout, None
        timed_out = True
    except Exception as e:
        output, rc = "%s: %s" % (e.__class__.__name__, e), None

    secs = int(time.time() - began)
    checks = []
    for line in output.splitlines():
        if line.startswith("[REPORT_CHECKS_JSON] "):
            try:
                checks = json.loads(line.split(" ", 1)[1])
            except (ValueError, TypeError):
                pass
    if echo and output.strip():
        for line in output.rstrip().splitlines():
            if not line.startswith("[REPORT_CHECKS_JSON] "):
                print("    %s" % line)

    # "文件在不在"和"校验过没过"是两件事，分别报。退出码见文件头的 RC_*：
    # 0 / 1 都表示写出来了；其余每一种都表示没写出来。但退出码只是 report_build
    # 自己的说法，这里还要拿磁盘上的文件对一遍（见上面 before_mtime 的注释）。
    exists = os.path.isfile(out)
    wrote = exists and (before_mtime is None
                        or os.path.getmtime(out) > before_mtime)
    stale = exists and not wrote          # 旧文件还在，但本次没写进去

    if stale:
        print("    [错误] 未能写入 %s" % os.path.basename(out))
        print("           磁盘上是上一次的旧文件，不是本次结果。")
        if "PermissionError" in output:
            print("           原因：文件正被 Excel/WPS 打开，请关闭后重跑。")

    r = {"ok": rc == RC_OK and wrote, "skipped": None, "returncode": rc,
         "seconds": secs, "out": out, "wrote": wrote, "stale": stale,
         "timed_out": timed_out,
         # rc 为 None 且不是超时：子进程根本没起来，和崩溃一样没有产出。
         "crashed": rc == RC_CRASH or (rc is None and not timed_out),
         "checks_failed": any(not c["ok"] and c["severity"] == "error" for c in checks),
         "data_incomplete": any(c["key"] == "ventas_incomplete" and not c["ok"]
                                for c in checks),
         "checks": checks, "diagnostic_log": None,
         "formula_errors": "[警告] 文件里存在公式错误" in output,
         "output": output[-6000:]}
    r["outcome"] = classify(r)

    # 只有没产出报表时才留现场。路径在这里不打印 —— run_reports 的汇总里会在
    # 那份报表的那一行下面给出，一次就够。
    if r["outcome"] not in WRITTEN_OUTCOMES:
        r["diagnostic_log"], log_error = _write_diag(out, month, pairs, rc, secs, output)
        if log_error:
            print("    [日志写入失败] %s" % log_error)
    return r
