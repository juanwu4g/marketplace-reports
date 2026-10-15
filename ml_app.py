# -*- coding: utf-8 -*-
"""「ML 报表」—— MercadoLibre 全部国家月度报表的统一窗口。

    双击 ML报表.cmd        （打开「立即运行」页）
    双击 启动.cmd          （同一个窗口，「立即运行」页；老名字，生产上还在用）
    双击 定时设置.cmd      （同一个窗口，「每月自动运行」页）
    或   .venv\\Scripts\\pythonw ml_app.py [--tab run|human|schedule]

    ┌ 顶部横幅：现在在跑什么（国家 · 月份 · 步骤 · 店铺 · 已运行多久）  [中止这一轮] ┐
    ├ 左：三页                         │ 右：运行日志（任何入口起的都能看）        │
    │   立即运行 / 需要人工处理 / 每月自动运行 │   看哪一轮 · 国家/店铺/级别/查找 · 跟到最新 │
    └ 底部：报表写到哪 · 用的哪个 Python                                          ┘

这个窗口只做三件事：把参数拼成 run_monthly.py 的命令行（或把它登记成计划
任务）、把那一轮的日志文件实时贴出来、把结果说清楚。**所有真正的逻辑都在
原来的脚本里**；窗口里能写成纯函数的规则都在 ml_entrance.py / region_choice.py
/ runlog.py，有测试钉着。这里只剩摆控件。

以前是两个窗口（启动 / 定时设置），各有一份日志查看器，按国家各登记一个计划
任务。现在一个窗口、一个推荐的计划任务（ML月度报表，--region ALL）。旧的按
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
    # ML报表.cmd 打开窗口之前的自检，必须在 import tkinter 之前：缺 tkinter 时
    # 要说一句人话，而不是甩一屏 traceback。见 ml_entrance.self_check。
    import console  # noqa: F401  中文输出编码保护
    import ml_entrance
    sys.exit(ml_entrance.self_check())

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import ml_entrance as E
import region_choice as RC
import runlog
import ui_style as S
import win_tasks as W
from log_view import LogView, fit_to_screen
# 报表输出目录的解析规则只有一份，在 run_monthly 里 —— 定时任务真正跑的是它。
from run_monthly import resolve_out_dir

ROOT = E.ROOT
LOG_DIR = os.path.join(ROOT, "logs")
TITLE = "MercadoLibre 月度报表"
LEFT_W = 560          # 左边那一栏的宽度（像素）
WRAP = LEFT_W - 60    # 左边说明文字的折行宽度
WHO = "ML报表窗口"     # 写进中止请求：是哪个窗口点的


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
        WRAP = LEFT_W - int(60 * k)


class App(object):
    def __init__(self, root, tab="run"):
        self.root = root
        root.title(TITLE)
        S.apply(root)
        _scale_widths(root)
        fit_to_screen(root, 1560, 900)
        root.minsize(1100, 640)

        s = E.load_settings()
        self.choice = RC.remembered_choice(s)   # 立即运行的范围：ALL 或某个国家
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
        self._build_run(self.tab_run)
        self._build_human(self.tab_human)
        self._build_schedule(self.tab_sched)

        self.view = LogView(right, log_dir=None, on_state=self._on_run_state,
                            on_tick=self._on_view_tick, who=WHO)
        self.view.pack(fill="both", expand=True, padx=(S.GAP, 0))

        self.show_tab(tab)
        self._fill_stores()
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
               "schedule": self.tab_sched}.get(name, self.tab_run)
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
        ttk.Label(f, text="国家").grid(row=0, column=0, sticky="w")
        self.region_box = ttk.Combobox(f, values=RC.choice_labels(), width=14,
                                       state="readonly")
        self.region_box.set(RC.label(self.choice))
        self.region_box.grid(row=0, column=1, sticky="w", padx=(S.GAP, 0))
        self.region_box.bind("<<ComboboxSelected>>", lambda e: self._pick_region())

        ttk.Label(f, text="月份").grid(row=1, column=0, sticky="w", pady=(S.GAP, 0))
        self.month = ttk.Combobox(f, values=self.months, width=14, state="readonly")
        self.month.set(self.months[0])
        self.month.grid(row=1, column=1, sticky="w", padx=(S.GAP, 0), pady=(S.GAP, 0))
        hint(f, "默认是上一个完整月份。当月还没过完，做月报没有意义。"
             ).grid(row=2, column=1, sticky="w", padx=(S.GAP, 0))

        ttk.Label(f, text="店铺").grid(row=3, column=0, sticky="nw", pady=(S.GAP, 0))
        box = ttk.Frame(f)
        box.grid(row=3, column=1, sticky="we", padx=(S.GAP, 0), pady=(S.GAP, 0))
        self.all_stores = tk.BooleanVar(value=True)
        self.all_chk = ttk.Checkbutton(box, text="全部店铺", variable=self.all_stores,
                                       command=self._toggle_stores)
        self.all_chk.pack(anchor="w")
        self.scope_var = tk.StringVar(value="")
        ttk.Label(box, textvariable=self.scope_var, style="Section.TLabel",
                  wraplength=WRAP - 60, justify="left").pack(anchor="w", pady=(2, 0))
        lb_row = ttk.Frame(box)
        lb_row.pack(fill="x", pady=(S.SMALL, 0))
        self.lb = tk.Listbox(lb_row, selectmode="extended", height=5,
                             exportselection=False, activestyle="none")
        sb = ttk.Scrollbar(lb_row, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="x", expand=True)
        sb.pack(side="left", fill="y")
        self.lb.bind("<<ListboxSelect>>", lambda e: self._update_estimate())
        self.store_hint = tk.StringVar(value="")
        hint(box, var=self.store_hint, wrap=WRAP - 60).pack(anchor="w", pady=(2, 0))
        self.stores = []

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
        hint(f, "只出报表：用已下载的数据重出，约 1 分钟，不开浏览器，可以随便重跑。\n"
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

    def _fill_stores(self):
        """按所选范围重填店铺清单。换国家时选中的店全部作废。"""
        self.groups = RC.lineup(self.choice)
        self.stores = [] if self.choice == RC.ALL else (
            self.groups[0][1] if self.groups else [])
        self.lb.configure(state="normal")
        self.lb.delete(0, "end")
        if self.choice == RC.ALL:
            for r, ss in self.groups:
                for s in ss:
                    self.lb.insert("end", "%s　（%s）" % (s, RC.label(r)))
            self.all_stores.set(True)
            self.all_chk.configure(state="disabled")
            self.store_hint.set("选了「全部国家」：一个国家跑完再跑下一个，每个国家"
                                "各出一份全店汇总表。只跑几家店请先选具体的国家。")
        else:
            for s in self.stores:
                self.lb.insert("end", s)
            self.all_chk.configure(state="normal")
            self.store_hint.set("取消勾选“全部店铺”后，可按住 Ctrl 多选。只选部分"
                                "店铺时不出全店汇总表（避免覆盖真正的那份）。")
        self.scope_var.set(RC.scope_text(self.groups))
        self._toggle_stores()
        self._update_estimate()
        self._lock(self.busy, force=True)

    def _pick_region(self):
        choice = RC.from_label(self.region_box.get())
        if choice == self.choice:
            return
        self.choice = choice
        E.update_settings(region=choice)
        self._fill_stores()

    def _toggle_stores(self):
        on = not self.all_stores.get() and self.choice != RC.ALL
        self.lb.configure(state="normal" if on else "disabled")
        self._update_estimate()

    def _update_estimate(self):
        n = sum(len(ss) for _, ss in getattr(self, "groups", []))
        if not self.all_stores.get() and self.choice != RC.ALL:
            n = len(self.lb.curselection()) or n
        self.estimate_var.set("%d 家店，%s" % (n, E.estimate("both", n))
                              if n else "所选范围里没有店铺")

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

    def _picked_stores(self):
        picked = [self.stores[i] for i in self.lb.curselection()
                  if i < len(self.stores)]
        subset, problem = E.check_stores(self.choice, self.all_stores.get(), picked)
        if problem:
            messagebox.showwarning("没选店铺", problem)
            return False, None
        return True, subset

    def do_start(self, step):
        ok, subset = self._picked_stores()
        if not ok:
            return
        month = self.month.get()
        if step != "collect":
            title, text = E.confirm_text(step, self.choice, month, subset,
                                         latest=self.months[0], groups=self.groups)
            if not messagebox.askokcancel(title, text):
                return
        self.start(step, E.STEP_TITLES[step], subset if step != "collect" else None,
                   month=month)

    def start(self, step, title, stores, month=None, reports_for=None, region=None):
        """起一轮：run_monthly.py --entrance manual ...，独立的后台进程。"""
        region = region or self.choice
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
        title = "%s：%s" % (RC.label(region), title)
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
        ok, subset = self._picked_stores()
        if not ok:
            return
        month = self.month.get()
        cmds = RC.preview_cmds(E.interpreter(), os.path.join(ROOT, "run_reports.py"),
                               month, subset, self.choice)
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
                parts.append("━━━━━━━━ %s ━━━━━━━━\n%s" % (RC.label(region), out))
            box["out"] = "\n".join(parts)

        def wait():
            if "out" not in box:
                self.root.after(150, wait)
                return
            self.btn_preview.configure(state="normal")
            self._show_text("先看用哪批数据 —— %s %s" % (RC.label(self.choice), month),
                            box["out"])

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
        cls = runlog.LineClassifier("report", [s for _, ss in self.groups for s in ss])
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
        row = ttk.Frame(page)
        row.pack(fill="x")
        self.human_lb = tk.Listbox(row, selectmode="extended", height=7,
                                   exportselection=False, activestyle="none")
        sb = ttk.Scrollbar(row, orient="vertical", command=self.human_lb.yview)
        self.human_lb.configure(yscrollcommand=sb.set)
        self.human_lb.pack(side="left", fill="x", expand=True)
        sb.pack(side="left", fill="y")
        self.human_lb.bind("<<ListboxSelect>>", lambda e: self._show_human())
        self.human_how = tk.StringVar(value="")
        # 固定高度：说明文字随选中项变长变短时，下面的按钮不能跟着上下跳 ——
        # 实测选中一家店后按钮上移，下一下就点空了。
        tk.Label(page, textvariable=self.human_how, fg=S.HINT, justify="left",
                 anchor="nw", height=5, wraplength=WRAP, font=S.FONT_SMALL
                 ).pack(anchor="w", fill="x", pady=(S.GAP, S.GAP))
        self.human_open = ttk.Button(page, text="① 打开所选店铺（去登录 / 处理）",
                                     command=self.do_open_stores)
        self.human_open.pack(anchor="w", fill="x")
        self.human_rerun = ttk.Button(page, text="② 处理好了：重跑这些店铺并重出全部报表",
                                      command=self.do_rerun_human)
        self.human_rerun.pack(anchor="w", fill="x", pady=(S.GAP, 0))
        hint(page, "重跑：只重新下载选中的店，然后给它们所在国家的全部店铺重出报表"
                   "（含全店汇总）。一次只能补跑一个国家的店。"
             ).pack(anchor="w", pady=(S.GAP, 0))

    def _refresh_human(self, announce=True):
        """重读清单。announce：有新的一轮带来了要处理的店，就弹窗提醒一次。"""
        try:
            self.human = runlog.pending_human_actions()
        except Exception as e:
            self.human = {"month": None, "actions": []}
            self.human_how.set("读不到上一轮的结果：%s" % e)
        acts = self.human["actions"]
        keep = set(self.human_lb.get(i) for i in self.human_lb.curselection())
        self.human_lb.delete(0, "end")
        for a in acts:
            self.human_lb.insert("end", RC.action_line(a))
        for i, a in enumerate(acts):
            if not keep or RC.action_line(a) in keep:
                self.human_lb.selection_set(i)
        text, kind = E.human_headline(self.human)
        self.human_head.set(("⚠ " if acts else "✓ ") + text)
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
                "在「需要人工处理」页选中店铺，点「打开所选店铺」去处理；\n"
                "处理完点「重跑这些店铺并重出全部报表」。"
                % (len(stores), "\n".join("· %s" % RC.action_line(a) for a in acts)))
        elif not announce:
            self.human_told = newest or None

    def _picked_human(self):
        acts = self.human["actions"]
        return [acts[i] for i in self.human_lb.curselection() if i < len(acts)]

    def _show_human(self):
        if self.human["actions"]:
            self.human_how.set(E.human_detail(self._picked_human()))
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
            messagebox.showinfo("打开店铺", "先在列表里选中要处理的店铺。")
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
        ttk.Label(page, text="推荐：一个任务，每月自动跑全部国家", style="Section.TLabel"
                  ).pack(anchor="w")
        hint(page, "任务名「%s」。墨西哥、智利……一个国家跑完接着跑下一个，同一份日志；"
                   "不会再有两个任务撞车、后一个被跳过的问题。" % RC.ALL_TASK
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

        f = section(page, "什么时候自动跑")
        s = E.load_settings()
        k = lambda name: RC.sched_key(name, RC.ALL)
        self.day = tk.IntVar(value=int(s.get(k("sched_day"), s.get("sched_day", 15))))
        self.hour = tk.IntVar(value=int(s.get(k("sched_hour"), s.get("sched_hour", 2))))
        self.minute = tk.IntVar(value=int(s.get(k("sched_minute"),
                                                s.get("sched_minute", 0))))
        row = ttk.Frame(f)
        row.pack(anchor="w")
        ttk.Label(row, text="每月").pack(side="left")
        ttk.Spinbox(row, from_=1, to=28, width=4, textvariable=self.day
                    ).pack(side="left", padx=S.SMALL)
        ttk.Label(row, text="号").pack(side="left")
        ttk.Spinbox(row, from_=0, to=23, width=4, format="%02.0f",
                    textvariable=self.hour).pack(side="left", padx=(S.PAD, S.SMALL))
        ttk.Label(row, text=":").pack(side="left")
        ttk.Spinbox(row, from_=0, to=59, width=4, increment=5, format="%02.0f",
                    textvariable=self.minute).pack(side="left", padx=S.SMALL)
        hint(f, "出的永远是上个月的报表（15 号跑，出上个月）。15 号上个月的账单已经"
                "结清；选凌晨是因为跑的时候会重开紫鸟浏览器。日期最大 28。"
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

        f = section(page, "按国家登记的旧任务")
        hint(f, "以前每个国家一个任务。这里只列出来，不会自动删除 —— 换成上面那个任务"
                "之后再删。").pack(anchor="w", pady=(0, S.SMALL))
        self.legacy_box = ttk.Frame(f)
        self.legacy_box.pack(fill="x")
        self.legacy_rows = {}
        for r in RC.REGIONS:
            row = ttk.Frame(self.legacy_box)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=RC.task_name(r), width=14).pack(side="left")
            var = tk.StringVar(value="…")
            ttk.Label(row, textvariable=var, style="Hint.TLabel").pack(
                side="left", fill="x", expand=True)
            b = ttk.Button(row, text="删除", width=6, state="disabled",
                           command=lambda name=RC.task_name(r): self.do_delete_legacy(name))
            b.pack(side="right")
            self.legacy_rows[r] = (var, b)

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

    def _save_time(self):
        k = lambda name: RC.sched_key(name, RC.ALL)
        E.update_settings(**{k("sched_day"): self.day.get(),
                             k("sched_hour"): self.hour.get(),
                             k("sched_minute"): self.minute.get()})

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
        n = len(RC.lineup(RC.ALL))
        ov = RC.schedule_overview(states, n)
        self.overview = ov
        self.sched_head.set(ov["headline"])
        self.sched_head_lbl.configure(foreground=S.FG[ov["kind"]])
        main = ov["main"]
        out, src = resolve_out_dir()
        lines = []
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
        legacy = dict((r, st) for r, _, st in ov["legacy"])
        for r, (var, b) in self.legacy_rows.items():
            st = legacy.get(r)
            var.set("%s　%s" % (RC.task_when(st), RC.task_verdict(st.get("RESULT")))
                    if st else "未登记")
            b.configure(state="normal" if st else "disabled")
        # 界面上的时间跟已登记的对齐：有新任务按新任务，只有旧任务按（墨西哥）旧任务。
        ref = main or (ov["legacy"][0][2] if ov["legacy"] else None)
        if ref:
            try:
                self.day.set(int(ref.get("DAY")))
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
        d, h, m = self.day.get(), self.hour.get(), self.minute.get()
        groups = RC.lineup(RC.ALL)
        legacy = self.overview.get("legacy") or []
        what = ("以后每月 %d 号 %02d:%02d，这台电脑会自己跑一轮：\n"
                "下载上个月的数据并出报表 —— %s。\n\n"
                "那时候电脑必须开着、已经登录、没睡着、没锁屏。\n\n"
                % (d, h, m, RC.scope_text(groups)))
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
        hours = RC.time_limit_hours(len(groups))
        args = RC.task_args(RC.ALL)
        self.log_op("")
        self.log_op("正在登记「%s」：每月 %d 号 %02d:%02d（%s，时限 %d 小时）"
                    % (RC.ALL_TASK, d, h, m, args, hours))

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
                "会让 Windows 真的把程序拉起来跑一遍，但只出报表（全部国家）——\n"
                "不下载、不开浏览器，跑完自己收拾干净。约 1–2 分钟。\n\n开始？"):
            return
        self.log_op("")
        self.log_op(">>> 试跑  %s" % datetime.datetime.now().strftime("%H:%M:%S"))
        self.view.follow_current()
        self._work(lambda: W.rehearse(RC.ALL), self._after_rehearse, label="试跑一次")

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
        n = sum(len(ss) for _, ss in RC.lineup(RC.ALL))
        if not messagebox.askokcancel(
                "现在就跑一次",
                "立刻按正式任务完整跑一遍：%s，%s。\n\n"
                "跑的时候浏览器会自己反复开关，别去点它。\n"
                "这个窗口可以关掉，不影响后台继续跑。\n\n开始？"
                % (RC.scope_text(RC.lineup(RC.ALL)), E.estimate("both", n))):
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
                "以后不再每月自动跑全部国家了。\n\n只是撤掉 Windows 里的登记 —— 程序、"
                "下载好的数据、已经出的报表都不动，手动跑也照常。\n\n确定？"):
            return
        self._work(lambda: W.delete_task(RC.ALL_TASK), self._after_delete,
                   label="取消每月自动运行")

    def do_delete_legacy(self, name):
        if not messagebox.askokcancel(
                "删除旧任务",
                "删除按国家登记的旧任务「%s」？\n\n%s"
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
            self.human_at = datetime.datetime.now().timestamp()
            self._refresh_human()
            if not self.sched_busy:
                self.refresh_schedule()

    def _lock(self, running, force=False):
        if running == self.busy and not force:
            return
        self.busy = running
        has = bool(getattr(self, "groups", None))
        picked = bool(self.human_lb.curselection()) if hasattr(self, "human_lb") else False
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
        E.update_settings(out_dir=self._out_dir(), region=self.choice)
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
