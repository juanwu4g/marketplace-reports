# -*- coding: utf-8 -*-
"""运行日志查看器：「ML 报表」窗口右边那一栏。

    view = LogView(parent, on_state=callback)
    view.pack(fill="both", expand=True)

它只做一件事：盯着 logs/runs/<run_id>.log 这个**文件**，有新内容就贴出来。
不从任何管道读 —— 所以不管那一轮是谁起的（本窗口点的、半夜计划任务自己跑
起来的、命令行敲的），看到的都一样；窗口关了重开，也能接着看。

这里**只有 Tk 胶水**。该看哪一轮、怎么跟读、怎么解析一行、一行分几列、
退出码是什么意思，全在 runlog.py（那边能在没有 tkinter 的机器上测）。

列靠 Text 控件的制表位对齐（而不是靠空格）：汉字和英文字宽不同，用空格
补齐在中文 Windows 上永远对不齐。正文太长时折行，折下来的行从「内容」那一
列开始，不用拖横向滚动条。

不要从无界面的代码 import 本模块：它一上来就 import tkinter。
"""
import os
import subprocess

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

from core import runlog
from core import ui_style as S

FOLLOW = "跟随当前运行（有新的一轮会自动切过去）"
ALL_STORES = "全部店铺"
ALL_REGIONS = "全部国家"
LEVEL_CHOICES = [("全部级别", None), ("警告和错误", runlog.WARN),
                 ("只看错误", runlog.ERROR)]
TICK_MS = 400
# 屏幕上最多放这么多行（筛选之后）。一轮全店下载几千行，远够；再多 Text 控件
# 重排会卡。文件里一行不少。
MAX_SHOWN = 20000


def fit_to_screen(root, want_w, want_h):
    """窗口不超过屏幕，放不下就直接最大化。

    运行日志是整个界面里唯一会伸缩的区域。窗口比屏幕大的话，被挤到屏幕外面
    的正是它 —— 在 1280x800 的机器上实测只剩一条缝。
    """
    try:
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        if want_h > sh - 90 or want_w > sw - 20:
            root.geometry("%dx%d+0+0" % (min(want_w, sw - 20), max(480, sh - 90)))
            try:
                root.state("zoomed")          # Windows；别的平台没有这个状态
            except Exception:
                pass
        else:
            # 放在任务栏上方居中。不指定位置的话系统随手一放，底边（状态栏）
            # 常常落到任务栏后面 —— 1920x1080 上实测如此。
            root.geometry("%dx%d+%d+%d" % (want_w, want_h, (sw - want_w) // 2,
                                           max(0, (sh - want_h - 90) // 2)))
    except Exception:
        root.geometry("%dx%d" % (want_w, want_h))


class LogView(ttk.Frame):
    """见模块说明。on_state(state) 在所看那一轮的状态变化时被调用，
    state 是 runlog.run_state() 的结果（没有任何运行记录时为 None）。
    on_tick(view) 每次轮询之后被调用（宿主拿 view.state / view.follower.records
    刷新顶上那条"现在在跑什么"）。"""

    def __init__(self, master, log_dir=None, on_state=None, on_tick=None,
                 who="", **kw):
        ttk.Frame.__init__(self, master, **kw)
        self.follower = runlog.Follower(log_dir)
        self.log_dir = log_dir
        self.on_state = on_state
        self.on_tick = on_tick
        self.who = who                # 写进中止请求里：是哪个窗口点的
        self.state = None
        self._labels = {}             # 下拉框里的文字 -> run_id
        self._with_region = False     # 这一轮是不是一轮多国（要不要国家那一列）
        self._widths = None
        self._font = tkfont.Font(family=S.MONO_FAMILY, size=S.FONT_LOG[1])
        self._bold = tkfont.Font(family=S.MONO_FAMILY, size=S.FONT_LOG[1],
                                 weight="bold")
        self._build()
        self.after(50, self._tick)

    # ---------------------------------------------------------- 布局

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="运行日志", style="Title.TLabel").pack(side="left")
        # 入口色块：一眼分清这是"手动"还是"定时"那一轮。
        self.badge = tk.Label(top, text=" — ", fg="white", bg=S.MUTED,
                              font=S.FONT_SMALL, padx=6, pady=1)
        self.badge.pack(side="left", padx=(S.GAP, 0))
        self.pick = ttk.Combobox(top, state="readonly", values=[FOLLOW],
                                 postcommand=self._refresh_runs)
        self.pick.set(FOLLOW)
        self.pick.pack(side="left", padx=(S.GAP, 0), fill="x", expand=True)
        self.pick.bind("<<ComboboxSelected>>", self._on_pick)
        ttk.Button(top, text="打开日志文件", command=self.open_file
                   ).pack(side="left", padx=(S.GAP, 0))

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(S.GAP, S.SMALL))
        self.region_lbl = ttk.Label(bar, text="国家")
        self.region = ttk.Combobox(bar, state="readonly", width=9,
                                   values=[ALL_REGIONS])
        self.region.set(ALL_REGIONS)
        self.region.bind("<<ComboboxSelected>>", lambda e: self._rerender())
        self.store_lbl = ttk.Label(bar, text="店铺")
        self.store_lbl.pack(side="left")
        self.store = ttk.Combobox(bar, state="readonly", width=20,
                                  values=[ALL_STORES])
        self.store.set(ALL_STORES)
        self.store.pack(side="left", padx=(S.SMALL, S.PAD))
        self.store.bind("<<ComboboxSelected>>", lambda e: self._rerender())
        ttk.Label(bar, text="级别").pack(side="left")
        self.level = ttk.Combobox(bar, state="readonly", width=10,
                                  values=[c[0] for c in LEVEL_CHOICES])
        self.level.set(LEVEL_CHOICES[0][0])
        self.level.pack(side="left", padx=(S.SMALL, S.PAD))
        self.level.bind("<<ComboboxSelected>>", lambda e: self._rerender())
        ttk.Label(bar, text="查找").pack(side="left")
        self.find = tk.StringVar()
        ttk.Entry(bar, textvariable=self.find, width=16).pack(
            side="left", padx=(S.SMALL, S.PAD))
        self._find_job = None
        self.find.trace_add("write", lambda *a: self._find_later())
        self.autoscroll = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="跟到最新一行", variable=self.autoscroll,
                        command=self._follow_toggled).pack(side="left")
        self.count_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.count_var, style="Hint.TLabel"
                  ).pack(side="right")

        # 表头和正文是两个 Text，用同一组制表位，所以列对得齐；表头不随正文滚走。
        box = tk.Frame(self, bd=1, relief="solid", background="#d0d7de")
        box.pack(fill="both", expand=True)
        self.head = tk.Text(box, height=1, wrap="none", font=self._bold,
                            bg="#f6f8fa", fg=S.HINT, bd=0, padx=8, pady=4,
                            highlightthickness=0, cursor="arrow")
        self.head.pack(fill="x")
        body = tk.Frame(box)
        body.pack(fill="both", expand=True)
        self.txt = tk.Text(body, wrap="word", bg=S.LOG_BG, fg=S.LOG_FG,
                           font=self._font, bd=0, padx=8, pady=4,
                           highlightthickness=0, insertwidth=0, spacing1=1,
                           spacing3=1)
        sy = ttk.Scrollbar(body, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sy.set)
        sy.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)
        self.txt.tag_config("time", foreground=S.LOG_TIME)
        self.txt.tag_config(runlog.ERROR, foreground=S.LOG_ERROR_FG,
                            background=S.LOG_ERROR_BG)
        self.txt.tag_config(runlog.WARN, foreground=S.LOG_WARN_FG,
                            background=S.LOG_WARN_BG)
        self.txt.tag_config("head", foreground=S.LOG_HEAD)
        self.txt.tag_config("banner", font=self._bold, background=S.LOG_BANNER_BG,
                            spacing1=6, spacing3=4)
        self.txt.tag_config("row")
        # 往上翻看的时候不要被新行拽回到底部
        self.txt.bind("<MouseWheel>", self._wheel)
        self.txt.bind("<Button-4>", self._wheel)
        self.txt.configure(state="disabled")
        self._set_columns()

    # ---------------------------------------------------------- 列

    def _headers(self):
        return (["时间"] + (["国家"] if self._with_region else [])
                + ["步骤", "店铺", "级别", "内容"])

    def _set_columns(self):
        """按现在要显示的列和认识的店名算制表位。店名长短不一，取最长的那个。"""
        f = self._font
        pad = 18
        samples = [["00:00:00", "时间"]]
        if self._with_region:
            samples.append(["墨西哥", "国家"])
        samples.append(["补收", "步骤"])
        samples.append(["店铺"] + list(self.follower.stores or [])[:200])
        samples.append(["警告", "级别"])
        stops, x = [], 0
        for col in samples:
            w = max(f.measure(t) for t in col if t)
            w = min(w, 260)                    # 极长的店名就让它挤一挤
            x += w + pad
            stops.append(x)
        widths = tuple(stops)
        if widths == self._widths:
            return
        self._widths = widths
        for t in (self.head, self.txt):
            t.configure(tabs=widths)
        # 折行从「内容」那一列开始
        self.txt.tag_config("row", lmargin2=widths[-1])
        self.head.configure(state="normal")
        self.head.delete("1.0", "end")
        self.head.insert("end", "\t".join(self._headers()))
        self.head.configure(state="disabled")

    def _show_region_filter(self, show):
        if show == self._with_region:
            return
        self._with_region = show
        if show:
            self.region_lbl.pack(side="left", before=self.store_lbl)
            self.region.pack(side="left", padx=(S.SMALL, S.PAD),
                             before=self.store_lbl)
        else:
            self.region_lbl.pack_forget()
            self.region.pack_forget()
            self.region.set(ALL_REGIONS)
        self._widths = None
        self._set_columns()

    # ---------------------------------------------------------- 筛选 / 渲染

    def _filters(self):
        s = self.store.get()
        r = self.region.get()
        code = None
        if r != ALL_REGIONS:
            for c in self.follower.regions:
                if runlog.region_name(c) == r:
                    code = c
        return {"store": None if s == ALL_STORES else s,
                "min_level": dict(LEVEL_CHOICES).get(self.level.get()),
                "region": code, "text": self.find.get().strip() or None}

    @staticmethod
    def _tags(rec):
        if rec.get("step") == runlog.STEP_START and "个国家：" in (rec.get("body") or ""):
            return ("row", "banner")
        if rec["level"] in (runlog.ERROR, runlog.WARN):
            return ("row", rec["level"])
        if rec["step"] == runlog.STEP_START:
            return ("row", "head")
        return ("row",)

    def _insert(self, rec):
        cols = runlog.view_columns(rec, self._with_region)
        tags = self._tags(rec)
        self.txt.insert("end", cols[0] + "\t", tags + ("time",),
                        "\t".join(cols[1:]) + "\n", tags)

    def _append(self, records, fresh=False):
        if not records:
            return
        f = self._filters()
        shown = [r for r in records if runlog.matches(r, **f)]
        if not shown:
            return
        self.txt.configure(state="normal")
        if fresh and len(shown) > MAX_SHOWN:
            self.txt.insert("end", "（前面还有 %d 行没显示；用上面的筛选缩小范围，"
                                   "或点「打开日志文件」看全文）\n"
                            % (len(shown) - MAX_SHOWN), ("head",))
            shown = shown[-MAX_SHOWN:]
        for rec in shown:
            self._insert(rec)
        if self.autoscroll.get():
            self.txt.see("end")
        self.txt.configure(state="disabled")

    def _clear(self):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")

    def _rerender(self):
        self._clear()
        self._append(self.follower.records, fresh=True)
        self._count()

    def _count(self):
        n = len(self.follower.records)
        f = self._filters()
        if any(f.values()):
            self.count_var.set("显示 %d / %d 行" % (len(self.follower.filtered(**f)), n))
        else:
            self.count_var.set("共 %d 行" % n if n else "")

    def _find_later(self):
        if self._find_job:
            self.after_cancel(self._find_job)
        self._find_job = self.after(300, self._rerender)

    def _wheel(self, event):
        up = getattr(event, "delta", 0) > 0 or getattr(event, "num", 0) == 4
        if up and self.autoscroll.get():
            self.autoscroll.set(False)

    def _follow_toggled(self):
        if self.autoscroll.get():
            self.txt.see("end")

    def _sync_choices(self, rerender=True):
        want = [ALL_STORES] + list(self.follower.stores)
        if list(self.store.cget("values")) != want:
            self.store.configure(values=want)
            self._set_columns()
        if self.store.get() not in want:
            self.store.set(ALL_STORES)
        regions = [ALL_REGIONS] + [runlog.region_name(r) for r in self.follower.regions]
        if list(self.region.cget("values")) != regions:
            self.region.configure(values=regions)
        need = bool(self.follower.regions)
        if need != self._with_region:
            self._show_region_filter(need)     # 多出 / 少了国家那一列
            if rerender:
                self._rerender()

    # ---------------------------------------------------------- 选哪一轮

    def _refresh_runs(self):
        self._labels = {}
        values = [FOLLOW]
        for run in runlog.list_runs(self.log_dir):
            label = runlog.run_label(run, self.log_dir)
            while label in self._labels:          # 同一分钟两轮：别互相盖掉
                label += " "
            self._labels[label] = run["run_id"]
            values.append(label)
        self.pick.configure(values=values)

    def _on_pick(self, _event=None):
        rid = self._labels.get(self.pick.get())
        if rid:
            self.follower.pin(rid)
        else:
            self.follower.follow_current()
        self._tick(reschedule=False)

    def follow_current(self):
        """回到跟随模式。起了新的一轮之后由宿主窗口调用。"""
        self.pick.set(FOLLOW)
        self.follower.follow_current()
        self.autoscroll.set(True)

    def open_file(self):
        """用系统默认程序打开正在看的这一轮的日志文件（看全文、搜索、转发）。"""
        rid = self.follower.run_id
        path = runlog.run_paths(rid, self.log_dir)["log"] if rid else ""
        if not path or not os.path.exists(path):
            path = runlog.runs_dir(self.log_dir)
        try:
            if os.name == "nt":
                os.startfile(path)                # noqa: S606
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass

    # ---------------------------------------------------------- 状态

    def _show_badge(self, st):
        if not st:
            self.badge.configure(text=" 还没有运行记录 ", bg=S.MUTED)
            return
        e = st.get("entrance") or ""
        self.badge.configure(text=" %s " % runlog.ENTRANCE_ZH.get(e, e or "?"),
                             bg=S.BADGE_BG.get(e, S.MUTED))

    # ---------------------------------------------------------- 轮询

    def _tick(self, reschedule=True):
        delay = TICK_MS
        try:
            r = self.follower.tick()
            if r["reset"]:
                self._clear()
                self.store.set(ALL_STORES)
                self.region.set(ALL_REGIONS)
                self._sync_choices(rerender=False)
                self._append(self.follower.records, fresh=True)
                self._count()
            elif r["new"]:
                self._append(r["new"])
                self._sync_choices()
                self._count()
            st = r["state"]
            self.state = st
            if r["state_changed"]:
                self._show_badge(st)
                if self.on_state:
                    self.on_state(st)
            if self.on_tick:
                self.on_tick(self)
            if r["more"]:
                delay = 30            # 还有没读完的（刚打开一份大日志）
        except Exception as e:
            # 查看器出错不该弄垮宿主窗口，但也不能悄悄停住 —— 把错摆出来。
            try:
                self.count_var.set("日志查看器出错：%s: %s"
                                   % (e.__class__.__name__, str(e)[:120]))
            except Exception:
                pass
        if reschedule:
            self.after(delay, self._tick)
