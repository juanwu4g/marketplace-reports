# -*- coding: utf-8 -*-
"""「月度报表」—— 全部平台、全部国家月度报表的统一窗口。

「立即运行」「每月自动运行」「期间报表」三页的范围都是同一棵 平台 → 国家（→ 店铺）勾选树
（core/scope.py + core/scope_picker.py），从 markets.py 和 config.json 算出来：加平台、
加国家不用改这个文件。

    双击 月度报表.cmd        （打开「立即运行」页；唯一的双击入口）
    或   .venv\\Scripts\\pythonw -m core.app [--tab run|human|schedule|period]

    ┌ 顶部横幅：现在在跑什么（国家 · 月份 · 步骤 · 店铺 · 已运行多久）  [中止这一轮] ┐
    ├ 左：三页                         │ 右：运行日志（任何入口起的都能看）        │
    │   立即运行 / 需要人工处理 / 每月自动运行 │   看哪一轮 · 国家/店铺/级别/查找 · 跟到最新 │
    └ 底部：报表写到哪 · 用的哪个 Python                                          ┘

这个窗口只做三件事：把参数拼成 run_monthly.py 的命令行（或把它登记成计划
任务）、把那一轮的日志文件实时贴出来、把结果说清楚。**所有真正的逻辑都在
原来的脚本里**；窗口里能写成纯函数的规则都在 entrance.py / region_choice.py
/ runlog.py，有测试钉着。这里只剩摆控件。

以前是两个窗口（启动 / 定时设置），各有一份日志查看器，按国家各登记一个计划
任务。现在一个窗口、一个推荐的计划任务（月度报表，--region ALL）。旧的按
国家任务（生产上的 MX月度报表）只显示、不自动删，由人点「替换」或「删除」。

几个刻意的设计（从旧窗口继承）：

* **绝不吞掉校验结果。** 跑完状态横幅按退出码把话说全，未过的行在日志里标色。
* **看的是日志文件，不是管道。** 起的那一轮是个独立的后台进程，窗口关了它
  照样跑；重新打开窗口，日志会自己接上。半夜定时跑起来的那轮也看得到、停得掉。
* **有一轮在跑就锁住按钮 —— 不管是谁起的。** 下载一启动就会重启紫鸟客户端，
  再起一轮会把正在跑的那轮杀掉。判断依据是 logs/current.json 加进程是否还活着。
"""
import datetime
import io
import os
import queue
import subprocess
import sys
import threading
import traceback

if __name__ == "__main__" and "--check" in sys.argv:
    # 月度报表.cmd 打开窗口之前的自检，必须在 import tkinter 之前：缺 tkinter 时
    # 要说一句人话，而不是甩一屏 traceback。见 entrance.self_check。
    from core import console  # noqa: F401  中文输出编码保护
    from core import entrance
    sys.exit(entrance.self_check())

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from core import entrance as E
from core import region_choice as RC
from core import runlog
from core import scope as SC
from core import ui_style as S
from core import win_tasks as W
from core.log_view import LogView, fit_to_screen
from core import period_report as PR
from core.period_report import choice as PC
from core.scope_picker import ScopePicker
# 报表输出目录的解析规则只有一份，在 run_monthly 里 —— 定时任务真正跑的是它。
from core.run_monthly import resolve_out_dir

ROOT = E.ROOT
LOG_DIR = runlog.LOG_DIR
TITLE = "月度报表"
LEFT_W = 560          # 左边那一栏的宽度（像素）
WRAP = LEFT_W - 85    # 左边说明文字的折行宽度（留出分组框边距和页面滚动条）
WHO = "月度报表窗口"     # 写进中止请求：是哪个窗口点的


def hint(parent, text="", var=None, wrap=WRAP, **kw):
    """灰色的说明文字。"""
    return ttk.Label(parent, text=text, textvariable=var, style="Hint.TLabel",
                     wraplength=wrap, justify="left", **kw)


def section(parent, title):
    """一组控件：带标题的分组框。"""
    f = ttk.LabelFrame(parent, text=" %s " % title, padding=(S.PAD, S.GAP))
    f.pack(fill="x", pady=(0, S.PAD))
    return f


def _scale_widths(root):
    """LEFT_W / WRAP 是按 100% 缩放写的像素。Windows 显示缩放 125%/150% 时字变大、
    像素宽度不变，左栏的按钮和说明就被截掉 —— 按 Tk 的缩放系数放大。"""
    global LEFT_W, WRAP
    try:
        k = float(root.tk.call("tk", "scaling")) / (96.0 / 72.0)
    except Exception:
        k = 1.0
    if k > 1.05:
        LEFT_W = int(LEFT_W * k)
        WRAP = LEFT_W - int(85 * k)


class App(object):
    def __init__(self, root, tab="run"):
        self.root = root
        root.title(TITLE)
        S.apply(root)
        _scale_widths(root)
        fit_to_screen(root, 1560, 900)
        root.minsize(1100, 640)

        s = E.load_settings()
        # 平台 → 国家 → 店铺 这棵树（markets.py + config.json）和两页各自的勾选。
        self.tree = SC.tree()
        old = RC.remembered_choice(s)          # 以前记的是一个国家或 ALL
        self.sel = SC.from_settings(s.get("run_scope"), self.tree,
                                    default=SC.from_region_arg(old, self.tree))
        self.results = SC.last_results(LOG_DIR)
        self.sched_dirty = False               # 计划任务那页的勾选改过、还没登记
        self.months = E.recent_months()
        self.launch = None            # 刚起、日志还没出现的那一轮
        self.busy = None
        self.active = None            # runlog.active_run() 的结果
        self.human = {"month": None, "actions": []}
        self.human_told = None
        self.human_at = 0.0
        self.q = queue.Queue()        # 计划任务那页的慢活（schtasks 要 2 秒）
        self.sched_busy = ""
        self.overview = None

        self._build_banner()
        self._build_footer()
        panes = ttk.PanedWindow(root, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=S.PAD, pady=(S.GAP, 0))
        left = ttk.Frame(panes, width=LEFT_W)
        right = ttk.Frame(panes)
        panes.add(left, weight=0)
        panes.add(right, weight=1)
        left.pack_propagate(False)      # 左栏按 LEFT_W 定宽，不随里面的控件缩小

        self.tabs = ttk.Notebook(left)
        self.tabs.pack(fill="both", expand=True, padx=(0, S.GAP))
        self.tab_run = self._scrollable_tab("  立即运行  ")
        self.tab_human = self._scrollable_tab("  需要人工处理  ")
        self.tab_sched = self._scrollable_tab("  每月自动运行  ")
        self.tab_period = self._scrollable_tab("  期间报表  ")
        self._build_run(self.tab_run)
        self._build_human(self.tab_human)
        self._build_schedule(self.tab_sched)
        self._build_period(self.tab_period)

        self.view = LogView(right, log_dir=None, on_state=self._on_run_state,
                            on_tick=self._on_view_tick, who=WHO)
        self.view.pack(fill="both", expand=True, padx=(S.GAP, 0))

        self.show_tab(tab)
        self._load_scope()
        self._refresh_human(announce=False)
        self.root.after(200, self._watch)
        self.root.after(80, self._drain)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.refresh_schedule()

        if E.interpreter() == sys.executable and not os.path.exists(
                os.path.join(ROOT, ".venv")):
            self._footer_note("⚠ 没找到 .venv，正在用 %s —— 依赖版本可能不对"
                              % sys.executable)

    # ================================================================ 框架

    def _scrollable_tab(self, title):
        """一页。内容多（小屏幕）时能上下滚，不会把按钮挤到屏幕外面。"""
        outer = ttk.Frame(self.tabs)
        self.tabs.add(outer, text=title)
        canvas = tk.Canvas(outer, highlightthickness=0, bd=0,
                           background=ttk.Style().lookup("TFrame", "background")
                           or "#f0f0f0")
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, padding=(S.PAD, S.PAD, S.PAD, 0))
        win = canvas.create_window(0, 0, window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)

        def resize(_e=None):
            canvas.itemconfigure(win, width=canvas.winfo_width())
            canvas.configure(scrollregion=canvas.bbox("all"))
            need = inner.winfo_reqheight() > canvas.winfo_height()
            if need and not sb.winfo_ismapped():
                sb.pack(side="right", fill="y")
            elif not need and sb.winfo_ismapped():
                sb.pack_forget()
        inner.bind("<Configure>", resize)
        canvas.bind("<Configure>", resize)

        def wheel(e):
            # 滚轮只管鼠标底下是这一页的时候（右边的日志自己会滚）。用路径判断
            # 而不是 Enter/Leave：鼠标移到页里的控件上时，这一页会收到 Leave。
            if not str(getattr(e, "widget", "")).startswith(str(outer)):
                return
            if inner.winfo_reqheight() > canvas.winfo_height():
                canvas.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1),
                                    "units")
        canvas.bind_all("<MouseWheel>", wheel, add="+")
        inner.outer = outer
        return inner

    def show_tab(self, name):
        tab = {"run": self.tab_run, "human": self.tab_human,
               "schedule": self.tab_sched, "period": self.tab_period}.get(name, self.tab_run)
        try:
            self.tabs.select(tab.outer)
        except tk.TclError:
            pass

    def _build_banner(self):
        """顶上那条：现在在跑什么。颜色跟着状态变。"""
        self.banner = tk.Frame(self.root, background=S.BANNER_BG[None])
        self.banner.pack(fill="x")
        self.banner_var = tk.StringVar(value="正在读取运行记录…")
        self.banner_lbl = tk.Label(self.banner, textvariable=self.banner_var,
                                   font=S.FONT_BANNER, anchor="w", justify="left",
                                   background=S.BANNER_BG[None], fg=S.FG[None],
                                   padx=S.PAD, pady=S.GAP)
        self.banner_lbl.pack(side="left", fill="x", expand=True)
        self.stop_btn = ttk.Button(self.banner, text="中止这一轮", width=12,
                                   command=self.do_stop, state="disabled")
        self.stop_btn.pack(side="right", padx=S.PAD, pady=S.SMALL)

    def _set_banner(self, text, kind):
        bg = S.BANNER_BG.get(kind, S.BANNER_BG[None])
        if self.banner_var.get() != text:
            self.banner_var.set(text)
        self.banner.configure(background=bg)
        self.banner_lbl.configure(background=bg, fg=S.FG.get(kind, S.FG[None]))

    def _build_footer(self):
        bar = ttk.Frame(self.root)
        bar.pack(side="bottom", fill="x", padx=S.PAD, pady=(S.SMALL, S.GAP))
        self.footer_var = tk.StringVar(value="")
        hint(bar, var=self.footer_var, wrap=900).pack(side="left")
        hint(bar, text="Python：%s" % E.interpreter(), wrap=600).pack(side="right")

    def _footer_note(self, text):
        self.footer_var.set(text)

    # ================================================================ 立即运行

    def _build_run(self, page):
        f = section(page, "范围")
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="月份").grid(row=0, column=0, sticky="w")
        self.month = ttk.Combobox(f, values=self.months, width=14, state="readonly")
        self.month.set(self.months[0])
        self.month.grid(row=0, column=1, sticky="w", padx=(S.GAP, 0))
        hint(f, "默认是上一个完整月份。当月还没过完，做月报没有意义。"
             ).grid(row=1, column=1, sticky="w", padx=(S.GAP, 0))
        # 平台 → 国家 → 店铺：点一行勾选 / 取消，点展开标记看店铺。平台和国家从
        # config / markets.py 来，加平台、加国家不用改窗口。
        self.run_picker = ScopePicker(f, with_stores=True, on_change=self._scope_changed,
                                      height=10)
        self.run_picker.grid(row=2, column=0, columnspan=2, sticky="we", pady=(S.GAP, 0))
        row = ttk.Frame(f)
        row.grid(row=3, column=0, columnspan=2, sticky="we", pady=(S.SMALL, 0))
        ttk.Button(row, text="全选", width=6,
                   command=lambda: self.run_picker.set_all(True)).pack(side="left")
        ttk.Button(row, text="全不选", width=7,
                   command=lambda: self.run_picker.set_all(False)
                   ).pack(side="left", padx=(S.SMALL, 0))
        self.scope_var = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.scope_var, style="Section.TLabel",
                  wraplength=WRAP, justify="left").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(S.SMALL, 0))
        hint(f, "点一行勾选或取消；点左边的展开标记（⊞）看店铺。勾上一个平台就是它下面的全部国家，"
                "勾上一个国家就是它的全部店铺。一个国家只勾了部分店铺时，这个国家不出"
                "全店汇总表（避免覆盖真正的那份）。"
             ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(2, 0))

        f = section(page, "报表输出到")
        f.columnconfigure(0, weight=1)
        self.outdir = tk.StringVar(value=E.load_settings().get("out_dir")
                                   or E.DEFAULT_REPORT_DIR)
        ttk.Entry(f, textvariable=self.outdir).grid(row=0, column=0, sticky="we")
        ttk.Button(f, text="浏览…", command=self._pick_outdir, width=8).grid(
            row=0, column=1, padx=(S.SMALL, 0))
        ttk.Button(f, text="恢复默认", command=self._reset_outdir, width=9).grid(
            row=0, column=2, padx=(S.SMALL, 0))
        hint(f, "可以直接指向共享盘。目录不存在会自动建。每月自动运行也写到这里。"
             ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(S.SMALL, 0))

        f = section(page, "运行")
        row = ttk.Frame(f)
        row.pack(fill="x")
        self.btn_run = ttk.Button(row, text="▶  下载并出报表", style="Accent.TButton",
                                  command=lambda: self.do_start("both"))
        self.btn_run.pack(side="left")
        self.estimate_var = tk.StringVar(value="")
        ttk.Label(row, textvariable=self.estimate_var, style="Hint.TLabel"
                  ).pack(side="left", padx=(S.PAD, 0))
        hint(f, "平时就点这一个：先下载所选店铺的数据，下完接着出报表，中间不用管。"
             ).pack(anchor="w", pady=(S.SMALL, S.GAP))
        row = ttk.Frame(f)
        row.pack(fill="x")
        self.btn_download = ttk.Button(row, text="只下载数据",
                                       command=lambda: self.do_start("download"))
        self.btn_report = ttk.Button(row, text="只出报表",
                                     command=lambda: self.do_start("report"))
        self.btn_collect = ttk.Button(row, text="补取欠的报表",
                                      command=lambda: self.do_start("collect"))
        for b in (self.btn_download, self.btn_report, self.btn_collect):
            b.pack(side="left", padx=(0, S.GAP))
        hint(f, "美客多每家店出两份报表：<店>_<月>_账单基准.xlsx（以账单为准，会计入账用）和 "
                "<店>_<月>.xlsx（以订单为准，经营分析参考）；全店汇总表也各一份。\n"
                "只出报表：用已下载的数据重出，约 1 分钟，不开浏览器，可以随便重跑。"
                "报表写着“暂定”（智利账单还没关账）时，等关账后点它重出。\n"
                "补取欠的报表：MercadoPago 有些报表要生成几十分钟；只打开确实欠着报表"
                "的店，不重新下载，和上面勾的店铺无关。"
             ).pack(anchor="w", pady=(S.SMALL, 0))

        f = section(page, "工具")
        row = ttk.Frame(f)
        row.pack(fill="x")
        self.btn_preview = ttk.Button(row, text="先看用哪批数据", command=self.do_preview)
        self.btn_preview.pack(side="left", padx=(0, S.GAP))
        ttk.Button(row, text="打开报表目录", command=self.do_open_dir
                   ).pack(side="left", padx=(0, S.GAP))
        hint(f, "先看用哪批数据：不生成任何文件，只列出每家店会用哪个下载目录、缺不缺"
                "该月账单。秒出，任何时候都能点。").pack(anchor="w", pady=(S.SMALL, 0))

    def _load_scope(self):
        """（重新）把树和勾选摆到「立即运行」页上。"""
        self.run_picker.load(self.tree, self.sel, self.results)
        self._scope_changed()

    def _scope_changed(self):
        self.sel = self.run_picker.selection()
        self.scope_var.set("已选：" + SC.scope_text(self.sel, self.tree))
        self._update_estimate()
        self._lock(self.busy, force=True)

    def _update_estimate(self):
        n = SC.store_count(self.sel, self.tree)
        self.estimate_var.set("%d 家店，%s" % (n, E.estimate("both", n))
                              if n else "一家店都没勾")

    def _pick_outdir(self):
        d = filedialog.askdirectory(title="选择报表输出目录",
                                    initialdir=self.outdir.get() or ROOT)
        if d:
            self.outdir.set(os.path.normpath(d))
            E.update_settings(out_dir=self.outdir.get())

    def _reset_outdir(self):
        self.outdir.set(E.DEFAULT_REPORT_DIR)
        E.update_settings(out_dir=E.DEFAULT_REPORT_DIR)

    def _out_dir(self):
        return (self.outdir.get() or "").strip() or E.DEFAULT_REPORT_DIR

    def _region(self):
        """勾选 -> run_monthly 的 (--region, --stores)。一个都没勾：提示，返回 (None, None)。"""
        region = SC.region_arg(self.sel, self.tree)
        if not region:
            messagebox.showwarning("没选店铺", "一家店都没勾。\n在「范围」里点平台、国家"
                                             "或店铺把它勾上，或者点「全选」。")
            return None, None
        return region, SC.subset(self.sel, self.tree)

    def do_start(self, step):
        region, subset = self._region()
        if not region:
            return
        month = self.month.get()
        if step != "collect":
            title, text = E.confirm_scope(step, SC.groups(self.sel, self.tree), month,
                                          latest=self.months[0])
            if not messagebox.askokcancel(title, text):
                return
        self.start(step, E.STEP_TITLES[step], subset if step != "collect" else None,
                   month=month, region=region)

    def start(self, step, title, stores, month=None, reports_for=None, region=None):
        """起一轮：run_monthly.py --entrance manual ...，独立的后台进程。"""
        region = region or SC.region_arg(self.sel, self.tree)
        cur = runlog.active_run()
        if cur or self.launch:
            messagebox.showwarning(
                "正在运行", "已经有一轮在跑了（入口：%s）。\n请等它结束，或先点"
                            "顶上的「中止这一轮」。"
                % runlog.ENTRANCE_ZH.get((cur or {}).get("entrance"), "手动"))
            return
        # 手输（而非"浏览…"选）的目录也先记下来：定时任务读的是这份设置。
        E.update_settings(out_dir=self._out_dir())
        argv = E.runner_cmd(step, month or self.month.get(), region, stores,
                            self._out_dir(), reports_for=reports_for)
        try:
            proc = runlog.spawn_detached(
                argv, cwd=ROOT, stderr_path=runlog.launch_error_path("manual"))
        except Exception as e:
            self._set_banner("✗ 启动失败：%s" % e, "bad")
            return
        title = "%s：%s" % ("全部平台和国家" if region == RC.ALL else "、".join(
            RC.unit_title(u) for u in RC.split_units(region)), title)
        self.launch = {"proc": proc, "at": datetime.datetime.now(), "title": title}
        self.view.follow_current()
        self._set_banner("● 正在启动：%s …（几秒后右边开始出日志）" % title, "run")
        self._lock(True)

    def do_stop(self):
        cur = runlog.active_run()
        if not cur:
            return
        if not messagebox.askokcancel(
                "中止这一轮",
                "确定要结束正在跑的这一轮吗？（入口：%s，%s）\n\n"
                "会连同它打开的紫鸟和浏览器一起关掉。\n"
                "已经下好的文件会保留；中止后不会接着出报表，后面的国家也不跑。"
                % (runlog.ENTRANCE_ZH.get(cur.get("entrance"), "?"),
                   runlog.region_name(cur.get("region") or "MX"))):
            return
        try:
            runlog.request_stop(cur["run_id"], WHO)
            self._set_banner("● 正在中止…（几秒内生效）", "warn")
            self.stop_btn.configure(state="disabled")
        except Exception as e:
            self._set_banner("✗ 中止请求写不进去：%s" % e, "bad")

    def do_preview(self):
        """只读预览：直接调 run_reports.py --dry-run，不经过运行器、不拿锁、
        不留运行记录；有一轮正在跑的时候也能点。全部国家时每个国家看一遍。"""
        region, _ = self._region()
        if not region:
            return
        month = self.month.get()
        script = ["-m", "platforms.mercadolibre.run_reports"]
        cmds = [(u.unit, RC.preview_cmd(E.interpreter(), script, month,
                                        ss if part else None, u.unit))
                for u, ss, part in SC.groups(self.sel, self.tree)]
        box = {}

        def body():
            kw = {}
            if os.name == "nt":
                # 本窗口是 pythonw 起的，没有控制台；不加这个会闪一个黑窗口。
                kw["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW",
                                              0x08000000)
            parts = []
            for region, cmd in cmds:
                try:
                    p = subprocess.run(
                        cmd, cwd=ROOT, stdin=subprocess.DEVNULL,
                        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        timeout=300, env=dict(os.environ, PYTHONIOENCODING="utf-8",
                                              PYTHONUTF8="1"), **kw)
                    out = (p.stdout or b"").decode("utf-8", "replace")
                except Exception as e:
                    out = "没跑成：%s: %s" % (e.__class__.__name__, e)
                parts.append("━━━━━━━━ %s ━━━━━━━━\n%s" % (RC.unit_title(region), out))
            box["out"] = "\n".join(parts)

        def wait():
            if "out" not in box:
                self.root.after(150, wait)
                return
            self.btn_preview.configure(state="normal")
            self._show_text("先看用哪批数据 —— %s" % month, box["out"])

        self.btn_preview.configure(state="disabled")
        threading.Thread(target=body, daemon=True).start()
        wait()

    def _show_text(self, title, text):
        w = tk.Toplevel(self.root)
        w.title(title)
        w.geometry("960x560")
        t = tk.Text(w, wrap="none", font=S.FONT_LOG, padx=8, pady=6, bd=0)
        sb = ttk.Scrollbar(w, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        t.pack(side="left", fill="both", expand=True)
        t.tag_config(runlog.WARN, foreground=S.LOG_WARN_FG, background=S.LOG_WARN_BG)
        t.tag_config(runlog.ERROR, foreground=S.LOG_ERROR_FG, background=S.LOG_ERROR_BG)
        t.tag_config("banner", font=S.FONT_BOLD)
        cls = runlog.LineClassifier("report", [s for _, ss, _ in SC.groups(self.sel, self.tree)
                                               for s in ss])
        for line in (text or "").splitlines():
            if line.startswith("━━"):
                t.insert("end", line + "\n", "banner")
                continue
            got = cls.feed(line)
            lv = got[2] if got else runlog.INFO
            t.insert("end", line + "\n", lv if lv != runlog.INFO else ())
        t.configure(state="disabled")

    def do_open_dir(self):
        d = self._out_dir()
        if not os.path.isdir(d):
            messagebox.showinfo("还没有报表", "目录还不存在：\n%s\n\n先点「只出报表」。" % d)
            return
        open_path(d)

    # ================================================================ 需要人工处理

    def _build_human(self, page):
        self.human_head = tk.StringVar(value="")
        self.human_head_lbl = ttk.Label(page, textvariable=self.human_head,
                                        font=S.FONT_BOLD, wraplength=WRAP,
                                        justify="left")
        self.human_head_lbl.pack(anchor="w", pady=(0, S.GAP))
        hint(page, "登录 / 验证、账号语言不对、店铺打不开、缺报表权限 —— 这些脚本自己"
                   "解决不了，放着不管每个月都会原样失败一次。"
             ).pack(anchor="w", pady=(0, S.GAP))
        # 平台 → 国家 → 店铺，和「立即运行」同一个勾选树；右边一列写这家店要做什么。
        # 一开始一家都不勾：处理哪几家、重跑哪几家，由人勾。
        self.human_picker = ScopePicker(page, with_stores=True, on_change=self._show_human,
                                        height=8, store_note=self._human_note,
                                        last_heading="要做什么", open_units=True)
        self.human_picker.pack(fill="x")
        row = ttk.Frame(page)
        row.pack(fill="x", pady=(S.SMALL, 0))
        ttk.Button(row, text="全选", width=6,
                   command=lambda: self.human_picker.set_all(True)).pack(side="left")
        ttk.Button(row, text="全不选", width=7,
                   command=lambda: self.human_picker.set_all(False)
                   ).pack(side="left", padx=(S.SMALL, 0))
        self.human_how = tk.StringVar(value="")
        # 固定高度：说明文字随选中项变长变短时，下面的按钮不能跟着上下跳 ——
        # 实测选中一家店后按钮上移，下一下就点空了。
        tk.Label(page, textvariable=self.human_how, fg=S.HINT, justify="left",
                 anchor="nw", height=5, wraplength=WRAP, font=S.FONT_SMALL
                 ).pack(anchor="w", fill="x", pady=(S.GAP, S.GAP))
        self.human_open = ttk.Button(page, text="① 打开勾选的店铺（去登录 / 处理）",
                                     command=self.do_open_stores)
        self.human_open.pack(anchor="w", fill="x")
        self.human_rerun = ttk.Button(page, text="② 处理好了：重跑勾选的店铺并重出报表",
                                      command=self.do_rerun_human)
        self.human_rerun.pack(anchor="w", fill="x", pady=(S.GAP, 0))
        hint(page, "两个按钮都只管勾上的店：可以只勾一两家先处理、先重跑，别的留着。"
                   "重跑＝只重新下载勾上的店，然后给它们所在的平台/国家的全部店铺"
                   "重出报表（含全店汇总）。可以一次勾好几个平台、国家的店。"
             ).pack(anchor="w", pady=(S.GAP, 0))

    def _refresh_human(self, announce=True):
        """重读清单。announce：有新的一轮带来了要处理的店，就弹窗提醒一次。"""
        try:
            self.human = runlog.pending_human_actions()
        except Exception as e:
            self.human = {"month": None, "actions": []}
            self.human_how.set("读不到上一轮的结果：%s" % e)
        acts = self.human["actions"]
        # 重读时保留人已经勾上的店（还在清单里的）；新来的店不勾。
        keep = set(self.human_picker.picked_stores()) if self.human_picker.t else set()
        tree = SC.tree(self._human_groups(acts))
        sel = {}
        for u in SC.units(tree):
            mine = [st for st in u.stores if st in keep]
            if mine:
                sel[u.unit] = mine
        self.human_picker.load(tree, sel)
        text, kind = E.human_headline(self.human)
        self.human_head.set(("⚠" if acts else "✓") + text)
        self.human_head_lbl.configure(foreground=S.FG[kind])
        try:
            self.tabs.tab(self.tab_human.outer, text="  %s  " % E.human_title(self.human))
        except tk.TclError:
            pass
        self._show_human()
        self._lock(self.busy, force=True)
        newest = max((a.get("run_id") or "" for a in acts), default="")
        if announce and acts and newest and newest != self.human_told:
            self.human_told = newest
            stores = sorted(set(a["store"] for a in acts))
            self.show_tab("human")
            messagebox.showwarning(
                "需要人工处理",
                "这一轮有 %d 家店铺需要人去处理，脚本自己解决不了：\n\n%s\n\n"
                "在「需要人工处理」页勾选要处理的店铺（可以只勾几家），点「打开勾选的店铺」；\n"
                "处理完点「重跑勾选的店铺并重出报表」。"
                % (len(stores), "\n".join("· %s" % RC.action_line(a) for a in acts)))
        elif not announce:
            self.human_told = newest or None

    @staticmethod
    def _human_groups(acts):
        """待办 -> [(区域, [店])]：平台、国家的顺序和「立即运行」一样。"""
        order = [u for u, _ in RC.lineup(RC.ALL)]
        groups = {}
        for a in acts:
            ss = groups.setdefault(RC.action_region(a), [])
            if a["store"] not in ss:
                ss.append(a["store"])
        return sorted(groups.items(), key=lambda g: order.index(g[0]) if g[0] in order
                      else len(order))

    def _human_note(self, unit, store):
        return "；".join(a["title"] for a in self.human["actions"] if a["store"] == store)

    def _picked_human(self):
        picked = set(self.human_picker.picked_stores())
        return [a for a in self.human["actions"] if a["store"] in picked]

    def _show_human(self):
        picked = self._picked_human()
        n = len(set(a["store"] for a in picked))
        self.human_open.configure(text="① 打开勾选的店铺（%s）" % (
            "%d 家" % n if n else "先在上面勾选"))
        self.human_rerun.configure(text="② 处理好了：重跑勾选的店铺并重出报表%s" % (
            "（%d 家）" % n if n else ""))
        if self.human["actions"]:
            self.human_how.set(E.human_detail(picked))
        elif not self.human_how.get().startswith("读不到"):
            self.human_how.set("")
        self._lock(self.busy, force=True)

    def do_open_stores(self):
        """把所选的店铺一家家打开，留给人去登录 / 处理。窗口开着不关。"""
        stores = []
        for a in self._picked_human():
            if a["store"] not in stores:
                stores.append(a["store"])
        if not stores:
            messagebox.showinfo("打开店铺", "先在列表里勾选要处理的店铺。")
            return
        if runlog.active_run():
            messagebox.showwarning("正在运行", "有一轮正在跑，等它结束再打开店铺。")
            return
        self._footer_note("正在打开 %s …（紫鸟窗口出来后，在里面完成登录 / 验证）"
                          % "、".join(stores))
        self.human_open.configure(state="disabled")

        def body():
            kw = {}
            if os.name == "nt":
                kw["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
            failed = []
            for s in stores:
                try:
                    p = subprocess.run(
                        [E.interpreter(), "-u", os.path.join(ROOT, "open_store.py"),
                         "--store", s], cwd=ROOT, capture_output=True,
                        timeout=900, env=dict(os.environ, PYTHONUTF8="1",
                                              PYTHONIOENCODING="utf-8"), **kw)
                    if p.returncode != 0:
                        failed.append("%s（退出码 %s）" % (s, p.returncode))
                except Exception as e:
                    failed.append("%s（%s）" % (s, e))
            self.root.after(0, lambda: self._opened(stores, failed))

        threading.Thread(target=body, daemon=True).start()

    def _opened(self, stores, failed):
        self._lock(self.busy, force=True)
        if failed:
            self._footer_note("✗ 没能打开：%s" % "、".join(failed))
        else:
            self._footer_note("已打开 %s。在紫鸟窗口里处理好以后，回来点「重跑」。"
                              % "、".join(stores))

    def do_rerun_human(self):
        """只重新下载这几家店，然后给它们所在国家的**全部**店铺重出报表。"""
        plan = RC.rerun_plan(self._picked_human())
        month = self.human["month"] or self.month.get()
        ok, title, text = E.rerun_confirm(plan, month)
        if not ok:
            messagebox.showinfo(title, text)
            return
        if not messagebox.askokcancel(title, text):
            return
        self.start("both", "重跑 %s 并重出全部报表" % "、".join(plan["stores"]),
                   plan["stores"], month=month, reports_for="all",
                   region=plan["region"])

    # ================================================================ 每月自动运行

    def _build_schedule(self, page):
        ttk.Label(page, text="推荐：一个任务，每月自动跑全部平台和国家", style="Section.TLabel"
                  ).pack(anchor="w")
        # 平台和国家从 config / markets.py 来，不在这里写死：加平台、加国家不用改窗口。
        hint(page, "任务名「%s」：%s，一个跑完接着跑下一个，同一份日志；不会再有两个任务"
                   "撞车、后一个被跳过的问题。" % (RC.ALL_TASK, RC.platforms_text())
             ).pack(anchor="w", pady=(2, S.GAP))

        f = section(page, "现在是什么状态")
        self.sched_head = tk.StringVar(value="正在读取…")
        self.sched_head_lbl = ttk.Label(f, textvariable=self.sched_head,
                                        font=S.FONT_BOLD, wraplength=WRAP,
                                        justify="left")
        self.sched_head_lbl.pack(anchor="w")
        self.sched_detail = tk.StringVar(value="")
        hint(f, var=self.sched_detail).pack(anchor="w", pady=(S.SMALL, 0))
        self.sched_notes = ttk.Frame(f)
        self.sched_notes.pack(fill="x", pady=(S.SMALL, 0))
        ttk.Button(f, text="刷新", command=self.refresh_schedule, width=8
                   ).pack(anchor="w", pady=(S.GAP, 0))

        f = section(page, "自动跑哪些")
        self.sched_picker = ScopePicker(f, with_stores=False,
                                        on_change=self._sched_scope_changed, height=12,
                                        fit=True)
        self.sched_picker.pack(fill="x")
        self.sched_picker.load(self.tree, SC.from_settings(
            E.load_settings().get("sched_scope"), self.tree), self.results)
        self.sched_scope_var = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.sched_scope_var, style="Section.TLabel",
                  wraplength=WRAP, justify="left").pack(anchor="w", pady=(S.SMALL, 0))
        hint(f, "全都勾上 = 以后新加的平台、国家也会自动跑（推荐）。只勾一部分 = 只跑勾上"
                "的，以后新加的要回来勾。改了勾选要点下面的「设为每月自动运行」才生效。"
             ).pack(anchor="w", pady=(2, 0))
        self._sched_scope_changed(dirty=False)

        f = section(page, "什么时候自动跑")
        s = E.load_settings()
        k = lambda name: RC.sched_key(name, RC.ALL)
        # 每月几号：可以写几天（3, 12, 26）。旧设置只有一个数，照样读得出来。
        self.days = tk.StringVar(value=", ".join(str(d) for d in RC.settings_days(s, RC.ALL)))
        self.hour = tk.IntVar(value=int(s.get(k("sched_hour"), s.get("sched_hour", 2))))
        self.minute = tk.IntVar(value=int(s.get(k("sched_minute"),
                                                s.get("sched_minute", 0))))
        row = ttk.Frame(f)
        row.pack(anchor="w")
        ttk.Label(row, text="每月").pack(side="left")
        ttk.Entry(row, width=12, textvariable=self.days).pack(side="left", padx=S.SMALL)
        ttk.Label(row, text="号").pack(side="left")
        ttk.Spinbox(row, from_=0, to=23, width=4, format="%02.0f",
                    textvariable=self.hour).pack(side="left", padx=(S.PAD, S.SMALL))
        ttk.Label(row, text=":").pack(side="left")
        ttk.Spinbox(row, from_=0, to=59, width=4, increment=5, format="%02.0f",
                    textvariable=self.minute).pack(side="left", padx=S.SMALL)
        hint(f, "出的永远是上个月的报表（15 号跑，出上个月）。15 号上个月的账单已经"
                "结清；选凌晨是因为跑的时候会重开紫鸟浏览器。日期最大 28。\n"
                "可以写几天，用逗号隔开，比如 3, 12, 26：每一天都是全部店铺完整跑一轮、出上个月的报表；"
                "后面那次有完整版本时，以最新的完整版本为准（早出一份，账单关账后再出一份定稿）。"
             ).pack(anchor="w", pady=(S.SMALL, 0))

        f = section(page, "操作")
        grid = ttk.Frame(f)
        grid.pack(fill="x")
        self.sched_buttons = []
        rows = [("设为每月自动运行", self.do_install,
                 "按上面的时间登记（已登记就按新时间改）。有按国家的旧任务会先问你要不要替换。"),
                ("试跑一次（1 分钟）", self.do_rehearse,
                 "让 Windows 真的拉起程序，只出报表、不下载，跑完自动删掉。提前确认到时候跑得起来。"),
                ("现在就跑一次", self.do_run_now,
                 "立刻按正式任务完整跑一遍（全部国家）。不影响以后的自动跑。"),
                ("取消每月自动运行", self.do_delete,
                 "只撤掉 Windows 里的登记，程序、数据、报表都不动。")]
        for i, (text, cmd, note) in enumerate(rows):
            b = ttk.Button(grid, text=text, command=cmd, width=18,
                           style="Accent.TButton" if i == 0 else "TButton")
            b.grid(row=i, column=0, sticky="w", pady=2)
            hint(grid, note, wrap=WRAP - 190).grid(row=i, column=1, sticky="w",
                                                    padx=(S.GAP, 0))
            self.sched_buttons.append(b)

        f = section(page, "旧任务")
        hint(f, "这个系统早期的版本按国家登记过计划任务。这台电脑上要是还留着，会列在下面；"
                "不会自动删除 —— 换成上面那个任务之后再删。").pack(anchor="w", pady=(0, S.SMALL))
        self.legacy_box = ttk.Frame(f)
        self.legacy_box.pack(fill="x")
        self.legacy_none = ttk.Label(self.legacy_box, text="这台电脑上没有旧任务。",
                                     style="Hint.TLabel")
        self.legacy_rows = {}
        self.legacy_frames = {}
        for _, name in RC.legacy_tasks():
            row = ttk.Frame(self.legacy_box)        # 只在这台电脑上真有这个任务时才显示
            ttk.Label(row, text=name, width=14).pack(side="left")
            var = tk.StringVar(value="…")
            ttk.Label(row, textvariable=var, style="Hint.TLabel").pack(
                side="left", fill="x", expand=True)
            b = ttk.Button(row, text="删除", width=6, state="disabled",
                           command=lambda name=name: self.do_delete_legacy(name))
            b.pack(side="right")
            self.legacy_rows[name] = (var, b)
            self.legacy_frames[name] = row

        f = section(page, "这台电脑")
        row = ttk.Frame(f)
        row.pack(fill="x")
        b = ttk.Button(row, text="检查这台电脑", command=self.do_check)
        b.pack(side="left", padx=(0, S.GAP))
        self.sched_buttons.append(b)
        ttk.Button(row, text="打开日志文件夹", command=self.do_logs
                   ).pack(side="left")
        hint(f, "检查：会不会睡着、会不会锁屏、开机要不要输密码。只看不改。"
             ).pack(anchor="w", pady=(S.SMALL, 0))

        f = section(page, "操作记录")
        self.ops = tk.Text(f, height=7, wrap="word", font=S.FONT_SMALL, bd=0,
                           background="#fbfbfc", padx=6, pady=4)
        self.ops.pack(fill="x")
        for tag in ("good", "warn", "bad"):
            self.ops.tag_configure(tag, foreground=S.FG[tag])
        self.ops.tag_configure("dim", foreground=S.HINT)
        self.ops.configure(state="disabled")

        if os.name != "nt":
            for b in self.sched_buttons:
                b.configure(state="disabled")

    def log_op(self, line, tag=None):
        self.ops.configure(state="normal")
        self.ops.insert("end", line + "\n", tag or ())
        self.ops.see("end")
        self.ops.configure(state="disabled")

    def _work(self, fn, done, label="", lock=True):
        """慢活丢到子线程，Tk 控件只在主线程碰。lock=False 给只读查询用。"""
        if lock and self.sched_busy:
            messagebox.showwarning("还在忙", "「%s」还没做完，等它结束再点。"
                                   % self.sched_busy)
            return
        if lock:
            self._sched_lock(label)

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
                # 先解锁再回调：回调里可能还要发起下一个操作（登记完顺手刷新）。
                if locked:
                    self._sched_lock("")
                fn(payload)
        except queue.Empty:
            pass
        self.root.after(80, self._drain)

    def _sched_lock(self, label):
        self.sched_busy = label
        if os.name != "nt":
            return
        for b in self.sched_buttons:
            b.configure(state="disabled" if label else "normal")

    def _days(self):
        """「什么时候自动跑」里写的几号 -> [3, 12, 26]；写错了返回 None。"""
        try:
            return RC.parse_days(self.days.get())
        except ValueError:
            return None

    def _save_time(self):
        k = lambda name: RC.sched_key(name, RC.ALL)
        kw = {k("sched_hour"): self.hour.get(), k("sched_minute"): self.minute.get()}
        days = self._days()
        if days:
            kw.update(RC.days_settings(days, RC.ALL))
        E.update_settings(**kw)

    def _sched_region(self):
        return SC.region_arg(self.sched_picker.selection(), self.tree)

    def _sched_scope_changed(self, dirty=True):
        sel = self.sched_picker.selection()
        self.sched_dirty = self.sched_dirty or dirty
        text = SC.scope_text(sel, self.tree)
        if sel and not SC.is_everything(sel, self.tree):
            text += "（以后新加的平台 / 国家不会自动包含）"
        self.sched_scope_var.set("勾选：" + text + (
            "　—— 改过了，点「设为每月自动运行」才生效" if self.sched_dirty else ""))

    def refresh_schedule(self):
        if os.name != "nt":
            self.sched_head.set("计划任务只在 Windows 上能设置。")
            return
        self._work(W.all_task_states, self._show_schedule, lock=False)

    def _show_schedule(self, states):
        if isinstance(states, tuple):                    # ("err", 说明)
            self.sched_head.set("读不到计划任务的状态")
            self.sched_head_lbl.configure(foreground=S.FG["bad"])
            self.log_op(states[1], "bad")
            return
        main = states.get(RC.ALL_TASK)
        reg = RC.args_region(main.get("ARGS")) if main else None
        if main and not self.sched_dirty:
            # 勾选跟着已登记的任务走（和下面的几号几点一样），除非人刚改过还没登记。
            self.sched_picker.load(self.tree, SC.from_region_arg(reg, self.tree))
            self._sched_scope_changed(dirty=False)
        n = len(RC.lineup(reg or RC.ALL))
        ov = RC.schedule_overview(states, n)
        self.overview = ov
        self.sched_head.set(ov["headline"])
        self.sched_head_lbl.configure(foreground=S.FG[ov["kind"]])
        main = ov["main"]
        out, src = resolve_out_dir()
        # 这个任务跑哪些店：按登记的参数（ALL = 全部），从 config 来。
        lines = ["会跑：%s" % SC.scope_text(SC.from_region_arg(reg, self.tree), self.tree)
                 if main else "勾选：%s" % SC.scope_text(self.sched_picker.selection(),
                                                       self.tree)]
        if main:
            lines += ["下次：%s" % (main.get("NEXT") or "还没排上"),
                      "上次：%s —— %s" % (main.get("LAST") or "还没跑过",
                                         RC.task_verdict(main.get("RESULT")))]
        lines.append("报表写到：%s（%s）" % (os.path.abspath(out), src))
        self.sched_detail.set("\n".join(lines))
        for w in self.sched_notes.winfo_children():
            w.destroy()
        for kind, text in ov["notes"]:
            ttk.Label(self.sched_notes, text=("⚠ " if kind != "good" else "") + text,
                      foreground=S.FG[kind], wraplength=WRAP, justify="left"
                      ).pack(anchor="w", pady=1)
        legacy = dict((name, st) for _, name, st in ov["legacy"])
        for name, (var, b) in self.legacy_rows.items():
            st = legacy.get(name)
            var.set("%s　%s" % (RC.task_when(st), RC.task_verdict(st.get("RESULT")))
                    if st else "未登记")
            b.configure(state="normal" if st else "disabled")
            row = self.legacy_frames[name]
            if st:
                row.pack(fill="x", pady=1)
            else:
                row.pack_forget()
        if legacy:
            self.legacy_none.pack_forget()
        else:
            self.legacy_none.pack(anchor="w")
        # 界面上的时间跟已登记的对齐：有新任务按新任务，只有旧任务按（墨西哥）旧任务。
        ref = main or (ov["legacy"][0][2] if ov["legacy"] else None)
        if ref:
            try:
                days = ref.get("DAYS") or [int(ref.get("DAY"))]
                self.days.set(", ".join(str(d) for d in days))
                h, m = ref.get("TIME", "02:00").split(":")
                self.hour.set(int(h))
                self.minute.set(int(m))
            except (ValueError, TypeError):
                pass

    def do_install(self):
        if not os.path.exists(W.RUNNER):
            messagebox.showerror("程序文件不全", "找不到 run_monthly.cmd：\n%s" % W.RUNNER)
            return
        if self.overview is None:
            # 还没读到这台电脑上登记了哪些任务：不能在不知道有没有旧任务的
            # 情况下登记（那样就不会问要不要替换 MX月度报表）。
            messagebox.showinfo("稍等", "还在读取已登记的计划任务，几秒后再点。")
            self.refresh_schedule()
            return
        try:
            d = RC.parse_days(self.days.get())
        except ValueError as e:
            messagebox.showwarning("设为每月自动运行", "「每月几号」写得不对：%s" % e)
            return
        h, m = self.hour.get(), self.minute.get()
        region = self._sched_region()
        if not region:
            messagebox.showwarning("设为每月自动运行", "「自动跑哪些」一个都没勾。")
            return
        sel = self.sched_picker.selection()
        groups = RC.lineup(region)
        legacy = self.overview.get("legacy") or []
        what = ("以后每月 %s 号 %02d:%02d，这台电脑会自己跑一轮%s：\n"
                "下载上个月的数据并出报表 —— %s。\n%s\n"
                "那时候电脑必须开着、已经登录、没睡着、没锁屏。\n\n"
                % (RC.days_text(d), h, m, "（每次都是全部店铺）" if len(d) > 1 else "",
                   SC.scope_text(sel, self.tree),
                   "以后新加的平台 / 国家也会自动跑。\n" if region == RC.ALL else
                   "只跑勾上的；以后新加的平台 / 国家不会自动包含。\n"))
        q = RC.replace_question(legacy)
        if q:
            ans = messagebox.askyesnocancel("设为每月自动运行", what + q)
            if ans is None:
                return
            drop = [name for _, name, _ in legacy] if ans else []
        else:
            if not messagebox.askokcancel("设为每月自动运行", what + "继续？"):
                return
            drop = []
        self._save_time()
        E.update_settings(sched_scope=SC.to_settings(sel, self.tree))
        hours = RC.time_limit_hours(len(groups))
        args = RC.task_args(region)
        self.log_op("")
        self.log_op("正在登记「%s」：每月 %s 号 %02d:%02d（%s，时限 %d 小时）"
                    % (RC.ALL_TASK, RC.days_text(d), h, m, args, hours))

        def body():
            rc, out = W.install(RC.ALL_TASK, W.task_xml(
                args, day=d, hour=h, minute=m,
                time_limit=RC.time_limit_xml(hours)))
            gone = []
            if rc == 0:
                for name in drop:
                    gone.append((name,) + W.delete_task(name))
            return rc, out, gone
        self._work(body, self._after_install, label="设为每月自动运行")

    def _after_install(self, r):
        if isinstance(r, tuple) and r and r[0] == "err":
            self.log_op("✗ %s" % r[1], "bad")
            return
        rc, out, gone = r
        if rc == 0:
            self.sched_dirty = False
            self._sched_scope_changed(dirty=False)
            self.log_op("✓ 登记好了。电脑那天要是关着，开机后会补跑；用电池也照跑。", "good")
            for name, drc, dout in gone:
                self.log_op("✓ 已删除旧任务「%s」" % name if drc == 0 else
                            "✗ 旧任务「%s」没删掉：%s" % (name, (dout or "").strip()),
                            "good" if drc == 0 else "bad")
            self.log_op("建议现在点一次「试跑一次」，1 分钟就知道到时候跑不跑得起来。", "dim")
        else:
            self.log_op("✗ 没登记成（错误码 %s）" % rc, "bad")
            self.log_op((out or "").strip() or "(没有更多信息)", "bad")
        self.refresh_schedule()

    def do_rehearse(self):
        if not messagebox.askokcancel(
                "试跑一次",
                "会让 Windows 真的把程序拉起来跑一遍，但只出报表（勾选的平台和国家）——\n"
                "不下载、不开浏览器，跑完自己收拾干净。约 1–2 分钟。\n\n开始？"):
            return
        self.log_op("")
        self.log_op(">>> 试跑  %s" % datetime.datetime.now().strftime("%H:%M:%S"))
        self.view.follow_current()
        region = self._sched_region() or RC.ALL
        self._work(lambda: W.rehearse(region), self._after_rehearse, label="试跑一次")

    def _after_rehearse(self, r):
        kind, val = r
        if kind == "err":
            self.log_op("✗ %s" % val, "bad")
            return
        meaning = {"0": ("✓ 试跑通过：Windows 能把程序拉起来，报表也出来了。", "good"),
                   "1": ("✓ 试跑通过：报表出来了，有几项要人看一眼（每个月都常见）。",
                         "good"),
                   "267009": ("⚠ 还没跑完，过一会儿点「刷新」", "warn")}
        msg, tag = meaning.get(val, ("✗ 试跑没成功（代码 %s：%s）。原因看右边日志里最近"
                                     "一轮「定时」的记录。"
                                     % (val, runlog.exit_meaning(val)[0]
                                        if str(val).isdigit() and int(val) < 8
                                        else "见日志"), "bad"))
        self.log_op(msg, tag)
        self.log_op("临时任务已经删掉了。", "dim")

    def do_run_now(self):
        if not (self.overview or {}).get("main"):
            messagebox.showinfo("现在就跑一次", "「%s」还没登记，先点「设为每月自动运行」。\n"
                                "（只想马上跑一遍的话，用「立即运行」页更直接。）"
                                % RC.ALL_TASK)
            return
        if runlog.active_run():
            messagebox.showwarning("正在运行", "已经有一轮在跑了。等它结束，或先中止它。")
            return
        reg = RC.args_region(self.overview["main"].get("ARGS"))
        sel = SC.from_region_arg(reg, self.tree)
        n = SC.store_count(sel, self.tree)
        if not messagebox.askokcancel(
                "现在就跑一次",
                "立刻按正式任务完整跑一遍：%s，%s。\n\n"
                "跑的时候浏览器会自己反复开关，别去点它。\n"
                "这个窗口可以关掉，不影响后台继续跑。\n\n开始？"
                % (SC.scope_text(sel, self.tree), E.estimate("both", n))):
            return
        self.view.follow_current()
        self._work(lambda: W.run_task(RC.ALL_TASK), self._after_run_now,
                   label="现在就跑一次")

    def _after_run_now(self, r):
        rc, out = r
        if rc == 0:
            self.log_op("✓ 已经开始跑了，进度看右边的运行日志。", "good")
        else:
            self.log_op("✗ 没能启动：%s" % ((out or "").strip() or rc), "bad")

    def do_delete(self):
        if not (self.overview or {}).get("main"):
            messagebox.showinfo("取消每月自动运行", "「%s」本来就没登记。" % RC.ALL_TASK)
            return
        if not messagebox.askokcancel(
                "取消每月自动运行",
                "以后不再每月自动跑了。\n\n只是撤掉 Windows 里的登记 —— 程序、"
                "下载好的数据、已经出的报表都不动，手动跑也照常。\n\n确定？"):
            return
        self._work(lambda: W.delete_task(RC.ALL_TASK), self._after_delete,
                   label="取消每月自动运行")

    def do_delete_legacy(self, name):
        if not messagebox.askokcancel(
                "删除旧任务",
                "删除旧任务「%s」？\n\n%s"
                % (name, "全部国家的任务（%s）已登记，它会接着跑这个国家。"
                   % RC.ALL_TASK if (self.overview or {}).get("main")
                   else "注意：全部国家的任务还没登记 —— 删掉之后这个国家就不会"
                        "自动跑了。")):
            return
        self._work(lambda: W.delete_task(name), self._after_delete, label="删除旧任务")

    def _after_delete(self, r):
        rc, out = r
        self.log_op("✓ 已删除。" if rc == 0 else "✗ 没能删除：%s" % ((out or "").strip()
                                                               or rc),
                    "good" if rc == 0 else "bad")
        self.refresh_schedule()

    def do_check(self):
        self.log_op("")
        self.log_op(">>> 检查这台电脑")
        self._work(lambda: W.machine_checks(RC.ALL), self._after_check,
                   label="检查这台电脑")

    def _after_check(self, items):
        if isinstance(items, tuple):
            self.log_op("✗ %s" % items[1], "bad")
            return
        bad = 0
        for ok, title, note in items:
            mark, tag = ("✓", "good") if ok else (("?", "warn") if ok is None
                                                  else ("✗", "bad"))
            bad += ok is False
            self.log_op("  %s  %s：%s" % (mark, title, note), tag)
        if not bad:
            self.log_op("全都没问题，这台电脑可以半夜没人管地自己跑。", "good")
            return
        self.log_op("有 %d 项不合格。不处理的话，到时候可能跑不起来。" % bad, "warn")
        if messagebox.askyesno(
                "要现在设好吗",
                "可以帮你把「不睡眠 / 不休眠 / 屏幕不自动关」一次设好。\n\n"
                "会弹一个 Windows 的管理员确认框，点「是」就行。\n\n"
                "另外两项（开机自动登录、屏保锁屏）需要你手动设，操作记录里写了在哪里设。"):
            self._work(W.fix_power, self._after_fix, label="设置电源")

    def _after_fix(self, r):
        rc, out = r
        self.log_op("已设置。再点一次「检查这台电脑」确认。" if rc == 0
                    else "没能设置：%s" % ((out or "").strip() or rc),
                    "good" if rc == 0 else "bad")

    def do_logs(self):
        d = runlog.runs_dir(LOG_DIR)
        if not os.path.isdir(d):
            d = LOG_DIR
        if not os.path.isdir(d):
            messagebox.showinfo("还没有日志", "还没跑过，所以还没有日志：\n%s" % d)
            return
        open_path(d)

    # ================================================================ 期间报表

    def _build_period(self, page):
        """季度 / 全年 / 自定义的期间报表：只用各月定稿的月度结果（run_period.py）。"""
        hint(page, "期间报表 = 把各月的定稿月度报表加起来：每个店铺每个月用最新的完整版本，"
                   "不重新下载、不重新读平台数据，几秒钟。每个平台 + 国家一份（币种不同）。"
             ).pack(anchor="w", pady=(0, S.GAP))
        latest = self.months[0]
        f = section(page, "期间")
        self.period_kind = tk.StringVar(value="quarterly")
        row = ttk.Frame(f)
        row.pack(anchor="w")
        for kind, text in PC.KINDS:
            ttk.Radiobutton(row, text=text, value=kind, variable=self.period_kind,
                            command=self._period_changed).pack(side="left", padx=(0, S.PAD))
        grid = ttk.Frame(f)
        grid.pack(anchor="w", pady=(S.SMALL, 0))
        ttk.Label(grid, text="季度").grid(row=0, column=0, sticky="w")
        self.period_quarter = ttk.Combobox(grid, values=PC.quarters(latest), width=10, state="readonly")
        self.period_quarter.set(PC.quarters(latest)[0])
        self.period_quarter.grid(row=0, column=1, sticky="w", padx=(S.GAP, S.PAD))
        ttk.Label(grid, text="年份").grid(row=0, column=2, sticky="w")
        self.period_year = ttk.Combobox(grid, values=PC.years(latest), width=8, state="readonly")
        self.period_year.set(PC.years(latest)[0])
        self.period_year.grid(row=0, column=3, sticky="w", padx=(S.GAP, 0))
        ttk.Label(grid, text="从").grid(row=1, column=0, sticky="w", pady=(S.SMALL, 0))
        months = E.recent_months(36)
        self.period_from = ttk.Combobox(grid, values=months, width=10, state="readonly")
        self.period_from.set(months[min(2, len(months) - 1)])
        self.period_from.grid(row=1, column=1, sticky="w", padx=(S.GAP, S.PAD), pady=(S.SMALL, 0))
        ttk.Label(grid, text="到").grid(row=1, column=2, sticky="w", pady=(S.SMALL, 0))
        self.period_to = ttk.Combobox(grid, values=months, width=10, state="readonly")
        self.period_to.set(months[0])
        self.period_to.grid(row=1, column=3, sticky="w", padx=(S.GAP, 0), pady=(S.SMALL, 0))
        for w in (self.period_quarter, self.period_year, self.period_from, self.period_to):
            w.bind("<<ComboboxSelected>>", lambda e: self._period_changed())
        self.period_text = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.period_text, style="Section.TLabel", wraplength=WRAP,
                  justify="left").pack(anchor="w", pady=(S.SMALL, 0))
        hint(f, "季度和全年也会在月报跑完时自动生成：季度最后一个月的月报跑完以后生成那个季度，"
                "12 月的还生成全年（美客多是账单基准）。这里可以随时重新生成，或者生成任意几个月。"
             ).pack(anchor="w", pady=(2, 0))

        f = section(page, "范围")
        self.period_picker = ScopePicker(f, with_stores=True, on_change=self._period_changed,
                                         height=8)
        self.period_picker.pack(fill="x")
        self.period_picker.load(self.tree, SC.everything(self.tree), self.results)
        row = ttk.Frame(f)
        row.pack(fill="x", pady=(S.SMALL, 0))
        ttk.Button(row, text="全选", width=6,
                   command=lambda: self.period_picker.set_all(True)).pack(side="left")
        ttk.Button(row, text="全不选", width=7,
                   command=lambda: self.period_picker.set_all(False)
                   ).pack(side="left", padx=(S.SMALL, 0))
        self.period_scope_var = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.period_scope_var, style="Section.TLabel",
                  wraplength=WRAP, justify="left").pack(anchor="w", pady=(S.SMALL, 0))
        hint(f, "全都勾上 = 每个有月度结果的店铺都算（包括 config.json 里已经去掉的店）。"
                "只勾部分店铺 = 报表里只有这几家。").pack(anchor="w", pady=(2, 0))

        f = section(page, "美客多用哪个基准")
        self.period_basis = tk.StringVar(value="bill")
        for basis, text in PC.BASES:
            ttk.Radiobutton(f, text=text, value=basis, variable=self.period_basis,
                            command=self._period_changed).pack(anchor="w")
        hint(f, "只对美客多有效；Falabella、亚马逊只有一种基准。").pack(anchor="w", pady=(2, 0))

        f = section(page, "各店各月（只看，不写文件）")
        self.period_grid = ttk.Treeview(f, columns=(), show="tree headings", height=8,
                                        selectmode="browse")
        self.period_grid.heading("#0", text="平台 / 店铺", anchor="w")
        self.period_grid.column("#0", width=190, stretch=True)
        self.period_grid.tag_configure("market", font=S.FONT_BOLD)
        self.period_grid.tag_configure("bad", foreground=S.FG["bad"])
        self.period_grid.pack(fill="x")
        self.period_grid.bind("<<TreeviewSelect>>", lambda e: self._period_reason())
        self.period_reasons = {}
        self.period_blocks = []
        self.period_seq = 0
        self.period_grid_var = tk.StringVar(value="")
        hint(f, var=self.period_grid_var).pack(anchor="w", pady=(S.SMALL, 0))
        ttk.Button(f, text="刷新", width=8, command=self.refresh_period
                   ).pack(anchor="w", pady=(S.SMALL, 0))

        f = section(page, "生成")
        row = ttk.Frame(f)
        row.pack(fill="x")
        self.btn_period = ttk.Button(row, text="▶  生成期间报表", style="Accent.TButton",
                                     command=self.do_period)
        self.btn_period.pack(side="left")
        ttk.Button(row, text="打开期间报表目录", command=self.do_open_period_dir
                   ).pack(side="left", padx=(S.GAP, 0))
        self.period_status = tk.StringVar(value="")
        self.period_status_lbl = ttk.Label(f, textvariable=self.period_status, font=S.FONT_BOLD,
                                           wraplength=WRAP, justify="left")
        self.period_status_lbl.pack(anchor="w", pady=(S.SMALL, 0))
        self.period_out = tk.Text(f, height=10, wrap="word", font=S.FONT_SMALL, bd=0,
                                  background="#fbfbfc", padx=6, pady=4)
        self.period_out.pack(fill="x", pady=(S.SMALL, 0))
        self.period_out.configure(state="disabled")
        hint(f, "写到「立即运行」页「报表输出到」的目录下（季度在 quarterly、全年在 annual、自定义在 custom"
                "（旧布局在「期间报表」文件夹里））。同名文件会被替换，报表的⑩列出和上一份相比换了版本的"
                "店铺月份。没有完整版本的店铺月份不计入合计，报表①有红字提示。").pack(anchor="w", pady=(S.SMALL, 0))
        self.period_proc = None
        self._period_changed()

    def _period(self):
        """The tab's period, or None (and the reason in the period line) when the
        choice is not valid (e.g. 自定义 from > to)."""
        try:
            return PC.period_of(self.period_kind.get(), self.period_quarter.get(),
                                self.period_year.get(), self.period_from.get(), self.period_to.get())
        except ValueError as e:
            self.period_text.set("✗ %s" % e)
            return None

    def _period_scope(self):
        sel = self.period_picker.selection()
        return SC.region_arg(sel, self.tree), SC.subset(sel, self.tree)

    def _period_changed(self):
        p = self._period()
        if p:
            self.period_text.set("期间：%s" % PR.period_cn(p))
        self.period_scope_var.set("已选：" + SC.scope_text(self.period_picker.selection(), self.tree))
        self.refresh_period()

    def refresh_period(self):
        """Store × month preview: ✓ grade of the final version, ✗ 无完整版本 (the
        reason from the newest version shows under the grid when a row is picked).
        Reads results files only; writes nothing."""
        tv = self.period_grid
        tv.delete(*tv.get_children(""))
        self.period_reasons, self.period_blocks = {}, []
        p = self._period()
        region, stores = self._period_scope()
        if not p or not region:
            self.period_grid_var.set("先选期间和范围。")
            self._period_lock()
            return
        # A year of thirty stores is hundreds of files: read them off the Tk thread.
        # Only the newest request is shown (a click while one is reading wins).
        self.period_seq += 1
        seq, basis, box = self.period_seq, self.period_basis.get(), {}
        self.period_grid_var.set("正在读取月度结果 …")
        self._period_lock()

        def body():
            try:
                box["blocks"] = PC.grid(p, region, stores, basis)
            except Exception as e:
                box["error"] = e

        threading.Thread(target=body, daemon=True).start()

        def poll():
            if seq != self.period_seq:
                return
            if not box:
                self.root.after(100, poll)
                return
            if "error" in box:
                self.period_grid_var.set("读不了月度结果：%s" % box["error"])
                self._period_lock()
                return
            self._show_period_grid(p, box["blocks"])

        poll()

    def _show_period_grid(self, p, blocks):
        tv = self.period_grid
        cols = ["m%d" % i for i in range(len(p.months))]
        tv.configure(columns=cols)
        for c, m in zip(cols, p.months):
            tv.heading(c, text=m, anchor="w")
            tv.column(c, width=110, anchor="w", stretch=False)
        missing, why = 0, []
        for b in blocks:
            mid = "m:" + b["market"]
            tv.insert("", "end", iid=mid, text=b["title"], values=[""] * len(cols), open=True,
                      tags=("market",))
            if b["error"]:
                tv.insert(mid, "end", iid=mid + ":err", text="读不了：%s" % b["error"],
                          values=[""] * len(cols), tags=("bad",))
                continue
            missing += b["missing"]
            for store, cells in b["rows"]:
                iid = "s:%s:%s" % (b["market"], store)
                tv.insert(mid, "end", iid=iid, text=store, values=[t for _, _, t, _ in cells],
                          tags=("bad",) if any(not ok for _, ok, _, _ in cells) else ())
                self.period_reasons[iid] = ["%s %s：%s" % (store, m, r)
                                            for m, ok, _, r in cells if not ok]
                why += self.period_reasons[iid]
        if not blocks:
            self.period_grid_var.set("所选范围在这个期间还没有任何月度结果，生成不了期间报表。")
        elif missing:
            self.period_grid_var.set("✗ 有 %d 个店铺月份没有完整版本，不会计入合计（点一行看那家店的原因）：\n%s"
                                     % (missing, why[0] + ("…" if len(why) > 1 else "")))
        else:
            self.period_grid_var.set("✓ 每个店铺每个月都有完整版本。")
        self.period_blocks = blocks
        self._period_lock()

    def _period_reason(self):
        iid = self.period_grid.focus()
        lines = self.period_reasons.get(iid)
        if lines:
            self.period_grid_var.set("\n".join(lines))

    def _period_lock(self):
        ok = (not self.period_proc and bool(self.period_blocks)
              and self._period() is not None and bool(self._period_scope()[0]))
        self.btn_period.configure(state="normal" if ok else "disabled")

    def _period_say(self, text):
        self.period_out.configure(state="normal")
        self.period_out.insert("end", text + "\n")
        self.period_out.see("end")
        self.period_out.configure(state="disabled")

    def do_period(self):
        """Run run_period.py for the tab's choice (a child process, like the other
        steps); its output streams into the box under the button."""
        p = self._period()
        region, stores = self._period_scope()
        if not p or not region or self.period_proc:
            return
        E.update_settings(out_dir=self._out_dir())
        cmd = PC.command(E.interpreter(), ROOT, p, region, stores, self.period_basis.get(),
                         self._out_dir())
        self.period_out.configure(state="normal")
        self.period_out.delete("1.0", "end")
        self.period_out.configure(state="disabled")
        self.period_status.set("● 正在生成 %s 的期间报表 …" % p.label)
        self.period_status_lbl.configure(foreground=S.FG["run"])
        try:
            proc = subprocess.Popen(cmd, cwd=ROOT, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                                    stderr=subprocess.STDOUT, env=PC.child_env(), **PC.creation_flags())
        except Exception as e:
            self.period_status.set("✗ 启动失败：%s" % e)
            self.period_status_lbl.configure(foreground=S.FG["bad"])
            return
        self.period_proc = proc
        self._period_lock()
        lines = queue.Queue()

        def pump():
            for raw in iter(proc.stdout.readline, b""):
                lines.put(raw.decode("utf-8", "replace").rstrip("\r\n"))
            lines.put(None)

        threading.Thread(target=pump, daemon=True).start()

        def drain():
            done = False
            try:
                while True:
                    line = lines.get_nowait()
                    if line is None:
                        done = True
                        break
                    self._period_say(line)
            except queue.Empty:
                pass
            if not done:
                self.root.after(100, drain)
                return
            code = proc.wait()
            self.period_proc = None
            text, kind = PC.exit_text(code)
            self.period_status.set(text)
            self.period_status_lbl.configure(foreground=S.FG[kind])
            self.refresh_period()

        drain()

    def do_open_period_dir(self):
        p = self._period()
        region, _ = self._period_scope()
        d = PC.folder(p, region or RC.ALL, self._out_dir()) if p else None
        if not d:
            messagebox.showinfo("还没有期间报表", "还没有生成过期间报表。先点「生成期间报表」。")
            return
        open_path(d)

    # ================================================================ 状态

    def _watch(self):
        """每 0.8 秒：有没有一轮在跑（任何入口起的），刚起的那轮怎么样了。"""
        try:
            self.active = runlog.active_run()
            if self.launch:
                self._check_launch(self.active)
            self._lock(bool(self.active) or bool(self.launch))
            now = datetime.datetime.now().timestamp()
            # 清单不常变：空闲时每 30 秒重读一次（定时任务跑完也能反映出来）
            if not self.busy and now - self.human_at > 30:
                self.human_at = now
                self._refresh_human()
        except Exception:
            pass
        self.root.after(800, self._watch)

    def _on_view_tick(self, view):
        """日志查看器每读一次：刷新顶上那条横幅。"""
        if self.launch and not self.active:
            return                       # 还在"正在启动…"
        st = view.state
        records = view.follower.records
        if self.active and (not st or st.get("run_id") != self.active.get("run_id")):
            # 右边钉住了另一轮：横幅仍然说正在跑的那一轮（没有店铺 / 步骤细节）
            st = runlog.run_state(self.active["run_id"])
            records = ()
        text, kind = E.status_line(st, records)
        self._set_banner(text, kind)

    def _check_launch(self, cur):
        proc, at = self.launch["proc"], self.launch["at"]
        if proc.poll() is None:
            if cur:
                self.launch = None        # 日志已经出现，交给查看器
            return
        # 进程已经结束。找它留下的那条记录（不能按 PID 找：Windows 上 .venv 里的
        # python.exe 只是个转发器）。
        self.launch = None
        since = at.replace(microsecond=0) - datetime.timedelta(seconds=1)
        mine = [r for r in runlog.list_runs(limit=5)
                if r["entrance"] == "manual" and r["when"] >= since]
        if mine:
            return
        err = ""
        try:
            with io.open(runlog.launch_error_path("manual"), "rb") as f:
                err = f.read().decode("utf-8", "replace").strip()
        except Exception:
            pass
        self._set_banner("✗ 没能启动（退出码 %s）" % proc.returncode, "bad")
        messagebox.showerror(
            "没能启动",
            "运行器一启动就退出了（退出码 %s），没有写出日志。\n\n%s"
            % (proc.returncode, err[-1500:] or "（没有任何错误输出）"))

    def _on_run_state(self, st):
        """所看那一轮的状态变了：结束了就刷新人工处理清单和计划任务状态。"""
        if st and st.get("state") != "running":
            self.results = SC.last_results(LOG_DIR)
            self.run_picker.set_results(self.results)
            self.sched_picker.set_results(self.results)
            self.period_picker.set_results(self.results)
            if not self.period_proc:
                self.refresh_period()          # a monthly run may have added final versions
            self.human_at = datetime.datetime.now().timestamp()
            self._refresh_human()
            if not self.sched_busy:
                self.refresh_schedule()

    def _lock(self, running, force=False):
        if running == self.busy and not force:
            return
        self.busy = running
        has = bool(getattr(self, "tree", None)) and bool(
            SC.store_count(getattr(self, "sel", {}), self.tree))
        picked = bool(self._picked_human()) if hasattr(self, "human_picker") else False
        st = E.button_states(bool(running), has, bool(self.human["actions"]), picked)
        for name, b in (("run", self.btn_run), ("download", self.btn_download),
                        ("report", self.btn_report), ("collect", self.btn_collect),
                        ("preview", self.btn_preview)):
            b.configure(state="normal" if st[name] else "disabled")
        self.stop_btn.configure(state="normal" if st["stop"] else "disabled")
        if hasattr(self, "human_open"):
            self.human_open.configure(state="normal" if st["open_stores"] else "disabled")
            self.human_rerun.configure(state="normal" if st["rerun"] else "disabled")

    def _on_close(self):
        E.update_settings(out_dir=self._out_dir(),
                          run_scope=SC.to_settings(self.sel, self.tree))
        self._save_time()
        if (self.launch or runlog.active_run()) and not messagebox.askokcancel(
                "还在运行",
                "有一轮还在跑。关掉这个窗口**不会**停它，它会在后台跑完，\n"
                "日志照常写，下次打开窗口还能接着看。\n\n"
                "想停它请点「取消」，再点顶上的「中止这一轮」。\n\n关闭窗口？"):
            return
        self.root.destroy()


def open_path(p):
    if os.name == "nt":
        os.startfile(p)                   # noqa: S606
    else:
        subprocess.Popen(["xdg-open", p])


def crash_log_path():
    return os.path.join(LOG_DIR, "_window_crash.log")


def main(argv=None):
    tab = E.pick_tab(sys.argv[1:] if argv is None else argv)
    try:
        root = tk.Tk()
    except Exception:
        print("打不开窗口：%s" % traceback.format_exc().strip().splitlines()[-1])
        return 1

    def report(exc, val, tb):
        # 界面里的错不能悄悄吞掉，也不该甩一屏 traceback 给财务：写进文件，弹一句话。
        text = "".join(traceback.format_exception(exc, val, tb))
        try:
            os.makedirs(LOG_DIR, exist_ok=True)
            with io.open(crash_log_path(), "a", encoding="utf-8") as f:
                f.write("%s\n%s\n" % (runlog.stamp(), text))
        except Exception:
            pass
        try:
            messagebox.showerror("出错了", "窗口里出了一个没料到的错：\n%s\n\n"
                                 "详细信息记在 %s，发给技术的人看。"
                                 % (str(val)[:300], crash_log_path()))
        except Exception:
            pass
    root.report_callback_exception = report
    App(root, tab=tab)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
