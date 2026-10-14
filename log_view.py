# -*- coding: utf-8 -*-
"""运行日志查看器：「启动」和「定时设置」两个窗口嵌的是同一个控件。

    view = LogView(parent, on_state=callback)
    view.pack(fill="both", expand=True)

它只做一件事：盯着 logs/runs/<run_id>.log 这个**文件**，有新内容就贴出来。
不从任何管道读 —— 所以不管那一轮是谁起的（本窗口点的、另一个窗口点的、
半夜计划任务自己跑起来的），看到的都一样；窗口关了重开，也能接着看。

这里**只有 Tk 胶水**。该看哪一轮、怎么跟读、怎么解析一行、退出码是什么意思，
全在 runlog.py（那边能在没有 tkinter 的机器上测）。要改行为先去那边找。

不要从无界面的代码 import 本模块：它一上来就 import tkinter。
"""
import datetime

import tkinter as tk
from tkinter import ttk, messagebox

import runlog

FOLLOW = "跟随当前运行（有新的一轮自动切过去）"
ALL_STORES = "全部店铺"
LEVEL_CHOICES = [("全部", None), ("警告和错误", runlog.WARN),
                 ("只看错误", runlog.ERROR)]

BADGE_BG = {"manual": "#0969da", "scheduled": "#8250df", "cli": "#57606a"}
KIND_FG = {"run": "#0969da", "good": "#1a7f37", "warn": "#9a6700",
           "bad": "#b42318"}
TICK_MS = 400


def fit_to_screen(root, want_w, want_h):
    """窗口不超过屏幕，放不下就直接最大化。

    运行日志在窗口最下面，是整个界面里唯一会伸缩的区域。窗口比屏幕高的话，
    被挤到屏幕外面的正是它 —— 在 1280x800 的机器上实测只剩一条缝。
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
    state 是 runlog.run_state() 的结果（没有任何运行记录时为 None）。"""

    def __init__(self, master, log_dir=None, on_state=None, height=14,
                 who="", **kw):
        ttk.Frame.__init__(self, master, **kw)
        self.follower = runlog.Follower(log_dir)
        self.log_dir = log_dir
        self.on_state = on_state
        self.who = who                # 写进中止请求里：是哪个窗口点的
        self.state = None
        self._labels = {}             # 下拉框里的文字 -> run_id
        self._stopping = None         # 已经发过中止请求的那一轮
        self._build(height)
        self.after(50, self._tick)

    # ---------------------------------------------------------- 布局

    def _build(self, height):
        top = ttk.Frame(self)
        top.pack(fill="x")
        # 入口徽标用实心色块：一眼就得分清这是"手动"还是"定时"那一轮。
        self.badge = tk.Label(top, text=" — ", fg="white", bg="#8c959f",
                              font=("", 10, "bold"), padx=8, pady=2)
        self.badge.pack(side="left")
        self.state_var = tk.StringVar(value="还没有运行记录")
        self.state_lbl = ttk.Label(top, textvariable=self.state_var,
                                   font=("", 10, "bold"))
        self.state_lbl.pack(side="left", padx=(8, 0))
        self.stop_btn = ttk.Button(top, text="中止这一轮", width=12,
                                   command=self.stop, state="disabled")
        self.stop_btn.pack(side="right")

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(4, 4))
        ttk.Label(bar, text="看哪一轮").pack(side="left")
        self.pick = ttk.Combobox(bar, state="readonly", width=30,
                                 values=[FOLLOW],
                                 postcommand=self._refresh_runs)
        self.pick.set(FOLLOW)
        self.pick.pack(side="left", padx=(4, 8), fill="x", expand=True)
        self.pick.bind("<<ComboboxSelected>>", self._on_pick)
        ttk.Label(bar, text="店铺").pack(side="left")
        self.store = ttk.Combobox(bar, state="readonly", width=14,
                                  values=[ALL_STORES])
        self.store.set(ALL_STORES)
        self.store.pack(side="left", padx=(4, 8))
        self.store.bind("<<ComboboxSelected>>", lambda e: self._rerender())
        ttk.Label(bar, text="级别").pack(side="left")
        self.level = ttk.Combobox(bar, state="readonly", width=8,
                                  values=[c[0] for c in LEVEL_CHOICES])
        self.level.set(LEVEL_CHOICES[0][0])
        self.level.pack(side="left", padx=(4, 8))
        self.level.bind("<<ComboboxSelected>>", lambda e: self._rerender())
        self.autoscroll = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="跟着滚到底", variable=self.autoscroll
                        ).pack(side="left")

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)
        self.txt = tk.Text(body, wrap="none", height=height, bg="#1e1e1e",
                           fg="#d4d4d4", insertbackground="#d4d4d4",
                           font=("Consolas", 9))
        sy = ttk.Scrollbar(body, orient="vertical", command=self.txt.yview)
        sx = ttk.Scrollbar(body, orient="horizontal", command=self.txt.xview)
        self.txt.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        self.txt.pack(side="left", fill="both", expand=True)
        self.txt.tag_config(runlog.ERROR, foreground="#f48771")
        self.txt.tag_config(runlog.WARN, foreground="#e2c08d")
        self.txt.tag_config("head", foreground="#9cdcfe")
        self.txt.tag_config("good", foreground="#89d185")
        self.txt.configure(state="disabled")

    # ---------------------------------------------------------- 筛选 / 渲染

    def _filters(self):
        s = self.store.get()
        lv = dict(LEVEL_CHOICES).get(self.level.get())
        return (None if s == ALL_STORES else s), lv

    @staticmethod
    def _tag(rec):
        if rec["level"] in (runlog.ERROR, runlog.WARN):
            return rec["level"]
        if rec["step"] == runlog.STEP_START:
            return "head"
        return ()

    @staticmethod
    def _shown(rec):
        """屏幕上显示的那一行。

        日志文件里每行都带日期和入口，那是给事后翻文件的人看的。窗口里这两样
        是多余的 —— 入口就是上面那块徽标，日期在「看哪一轮」里 —— 省下来的
        二十几列留给正文，窄屏上少拖一次横向滚动条。文件内容不变。
        """
        if not rec.get("time"):
            return rec["raw"]                 # 解析不了的行原样显示
        return " | ".join([rec["time"][-8:], rec["step"], rec["store"],
                           rec["level"], rec["message"]])

    def _append(self, records):
        if not records:
            return
        store, lv = self._filters()
        self.txt.configure(state="normal")
        for rec in records:
            if runlog.matches(rec, store, lv):
                self.txt.insert("end", self._shown(rec) + "\n", self._tag(rec))
        if self.autoscroll.get():
            self.txt.see("end")
        self.txt.configure(state="disabled")

    def _clear(self):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")

    def _rerender(self):
        self._clear()
        self._append(self.follower.records)

    def _sync_stores(self):
        want = [ALL_STORES] + list(self.follower.stores)
        if list(self.store.cget("values")) != want:
            self.store.configure(values=want)
        if self.store.get() not in want:
            self.store.set(ALL_STORES)

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

    # ---------------------------------------------------------- 状态

    def _show_state(self, st):
        self.state = st
        if not st:
            self.badge.configure(text=" — ", bg="#8c959f")
            self.state_var.set("还没有运行记录")
            self.stop_btn.configure(state="disabled")
            return
        e = st.get("entrance") or ""
        self.badge.configure(text=" %s " % runlog.ENTRANCE_ZH.get(e, e or "?"),
                             bg=BADGE_BG.get(e, "#8c959f"))
        month = st.get("month") or ""
        stopping = (st["state"] == "running"
                    and self._stopping == st.get("run_id"))
        if stopping:
            text = "● 正在中止…（几秒内生效）"
        elif st["state"] == "running" and st.get("stalled"):
            # 短说法；完整的那句在窗口底部的状态栏里，这一行还要给按钮留地方
            text = "▲ 没有响应，多半卡住了"
        elif st["state"] == "running":
            text = "● 正在跑　%s　%s" % (month, self._elapsed(st))
        elif st["state"] == "aborted":
            text = "■ 已中止　%s" % month
        elif st["state"] == "skipped":
            text = "■ %s" % st["text"]
        else:
            text = "■ 已结束　%s　%s" % (month, st["text"])
        self.state_var.set(text)
        self.state_lbl.configure(foreground=KIND_FG.get(st.get("kind"), ""))
        self.stop_btn.configure(
            state="normal" if st["state"] == "running" and not stopping
            else "disabled")

    @staticmethod
    def _elapsed(st):
        parsed = runlog.parse_run_id(st.get("run_id"))
        if not parsed:
            return ""
        secs = int((datetime.datetime.now() - parsed[0]).total_seconds())
        if secs < 0:
            return ""
        return "已跑 %d:%02d:%02d" % (secs // 3600, secs % 3600 // 60, secs % 60)

    def stop(self, confirm=True):
        """请正在看的那一轮停下来。它不是本窗口起的也一样能停。"""
        st = self.state
        if not st or st.get("state") != "running":
            return False
        e = runlog.ENTRANCE_ZH.get(st.get("entrance"), "")
        if confirm and not messagebox.askokcancel(
                "中止",
                "确定要结束正在跑的这一轮吗？（入口：%s）\n\n"
                "会连同它打开的紫鸟和浏览器一起关掉。\n"
                "已经下好的文件会保留；中止后不会接着出报表。" % e):
            return False
        try:
            runlog.request_stop(st["run_id"], self.who, self.log_dir)
        except Exception as ex:
            messagebox.showerror("中止", "中止请求写不进去：%s" % ex)
            return False
        self._stopping = st["run_id"]
        self._show_state(st)
        return True

    # ---------------------------------------------------------- 轮询

    def _tick(self, reschedule=True):
        delay = TICK_MS
        try:
            r = self.follower.tick()
            if r["reset"]:
                self._clear()
                self.store.set(ALL_STORES)
                self._sync_stores()
                self._append(self.follower.records)
            else:
                self._append(r["new"])
            if r["new"]:
                self._sync_stores()
            st = r["state"]
            # 在跑的时候每次都刷新：状态栏上有个"已跑多久"。
            if r["state_changed"] or (st and st.get("state") == "running"):
                self._show_state(st)
            if r["state_changed"] and self.on_state:
                self.on_state(st)
            if r["more"]:
                delay = 30            # 还有没读完的（刚打开一份大日志）
        except Exception as e:
            # 查看器出错不该弄垮宿主窗口，但也不能悄悄停住 —— 把错摆出来。
            try:
                self.state_var.set("日志查看器出错：%s: %s"
                                   % (e.__class__.__name__, str(e)[:120]))
            except Exception:
                pass
        if reschedule:
            self.after(delay, self._tick)
