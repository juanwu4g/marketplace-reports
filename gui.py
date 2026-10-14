# -*- coding: utf-8 -*-
"""
给不写代码的人用的小窗口：选月份、选店铺、点按钮。

    双击 启动.cmd     （推荐，它会用对解释器）
    或   .venv\\Scripts\\python gui.py

这个界面只做三件事：把参数拼成 run_monthly.py 的命令行、把那一轮的日志实时
贴出来、把结果说清楚。**所有真正的逻辑都在原来的脚本里**，这里一行业务判断
都没有 —— 界面坏了不该影响出数，出数的口径也不该藏在界面里。

和「定时设置」跑的是同一个过程
------------------------------
这个窗口**不再自己起 run_batch / run_reports**。每个按钮起的都是
run_monthly.py —— 计划任务起的也是它，命令行只差一个 `--entrance manual`。
锁、日志文件、告警文件、"下载失败就不出报表"、下载超时，都在那一个脚本里，
两个入口不可能再各走各的路。以前这里自己串两步，默认月份还不带 --month，
走的下载路线和定时任务不一样。

几个刻意的设计：

* **绝不吞掉校验结果。** 这套报表最值钱的就是那十几条硬校验。跑完只显示
  一个绿色"完成"等于把它们全扔了，所以状态栏按退出码把话说全，未过的行
  在日志里标色。

* **看的是日志文件，不是管道。** 起的那一轮是个独立的后台进程，窗口关了它
  照样跑；重新打开窗口，下面的「运行日志」会自己接上。定时任务半夜跑起来的
  那一轮，在这里也看得到、也能中止。

* **有一轮在跑就锁住按钮 —— 不管是谁起的。** 下载脚本启动时会 taskkill 掉
  已运行的紫鸟客户端，再起一轮会把正在跑的那轮杀掉。判断依据是
  logs/current.json 加进程是否还活着，而不是本窗口手里有没有句柄。
"""
import datetime
import io
import json
import os
import subprocess
import sys
import threading

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import region_choice
import runlog
import site_time
from log_view import LogView, fit_to_screen

ROOT = os.path.dirname(os.path.abspath(__file__))
DEFAULT_REPORT_DIR = os.path.join(ROOT, "downloads", "data", "reports", "financial")
SETTINGS = os.path.join(ROOT, "gui_settings.json")
# 左列说明文字的折行宽度（像素）。左列宽度由它决定，右边全部留给日志。
WRAP = 470
RUNNER = os.path.join(ROOT, "run_monthly.py")

KIND_FG = {"run": "#0969da", "good": "#1a7f37", "warn": "#9a6700",
           "bad": "#b42318"}
KIND_MARK = {"good": "✓", "warn": "⚠", "bad": "✗"}


def load_settings():
    """界面自己的偏好（目前只有输出目录）。

    刻意不写进 config.json —— 那个文件装凭据和店铺白名单，是照着
    config.example.json 填出来的；把界面偏好混进去，新人对着范本填的时候
    会多出不认识的键。这份坏了删掉即可，不影响任何业务。
    """
    try:
        with io.open(SETTINGS, encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def save_settings(d):
    try:
        with io.open(SETTINGS, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass                      # 存不下偏好不该妨碍干活


def interpreter():
    """永远优先用项目自己的 .venv。

    不是洁癖：requirements.txt 锁的是 pandas 3，而机器上常见的 conda 环境
    是 pandas 2，两者对 NaN 转字符串的处理不同，报表会崩在 ⑨ 库存页。
    """
    p = os.path.join(ROOT, ".venv", "Scripts",
                     "python.exe" if os.name == "nt" else "python")
    return p if os.path.exists(p) else sys.executable


def run_hours(n):
    """n 家店大概要跑多久，给确认框用。每家约 15 分钟。"""
    m = n * 15
    return "%d 分钟" % m if m < 60 else "%.1f 小时" % (m / 60.0)


def recent_months(n=15):
    d = site_time.site_today().replace(day=1)   # 和 run_reports.prev_month 同一个"今天"
    out = []
    for _ in range(n):
        d -= datetime.timedelta(days=1)
        out.append("%04d-%02d" % (d.year, d.month))
        d = d.replace(day=1)
    return out


def store_list(region=None):
    """店铺白名单。读不到就返回空 —— 界面照常能开，按钮会提示配置有问题。

    给了 region 就是那个国家的"全部店铺"（store_config.stores_for_region）。
    """
    if region:
        return region_choice.stores(region)
    try:
        sys.path.insert(0, ROOT)
        import store_config
        return list(store_config.batch_stores())
    except Exception:
        return []


class App(object):
    def __init__(self, root):
        self.root = root
        # 刚起、日志还没出现的那一轮：{"proc", "at", "title"}。日志一出现就
        # 交给下面的查看器，这里不再管它。
        self.launch = None
        self.busy = None
        self.human = {"month": None, "actions": []}
        self.human_told = None        # 已经为哪一轮弹过"需要人工处理"
        self.human_at = 0.0
        # 选的是哪个国家。店铺清单、"全部店铺"、预览、起的每一轮都跟着它走。
        self.region = region_choice.remembered(load_settings())
        root.title(region_choice.window_title(self.region))
        fit_to_screen(root, 1680, 900)
        root.minsize(1000, 600)

        # 左边是参数和按钮（宽度固定），右边整列都给运行日志。
        # 以前是上下排：参数、按钮占掉 600 像素，日志挤在最底下 —— 在
        # 1280x800 的屏幕上只剩一行高，而日志恰恰是跑起来以后唯一要看的东西。
        self._build_status()          # 先占住底边，再分左右
        self.left = ttk.Frame(root)
        self.left.pack(side="left", fill="y")
        self.right = ttk.Frame(root)
        self.right.pack(side="left", fill="both", expand=True)
        self._build_top()
        self._build_buttons()
        self._build_human()
        self._build_log()
        self._refresh_human(announce=False)
        self.root.after(200, self._watch)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if interpreter() == sys.executable and not os.path.exists(
                os.path.join(ROOT, ".venv")):
            self._say("⚠ 没找到 .venv，正在用 %s —— 依赖版本可能不对"
                      % sys.executable, "warn")

    # ---------------------------------------------------------------- 布局

    def _build_top(self):
        f = ttk.LabelFrame(self.left, text=" 参数 ", padding=10)
        f.pack(fill="x", padx=(12, 6), pady=(12, 6))

        # 国家放在最上面：下面的店铺清单、"全部店铺"指哪几家都由它决定。
        ttk.Label(f, text="国家").grid(row=0, column=0, sticky="w")
        self.region_box = ttk.Combobox(f, values=region_choice.labels(), width=12,
                                       state="readonly")
        self.region_box.set(region_choice.label(self.region))
        self.region_box.grid(row=0, column=1, sticky="w", padx=(6, 18))
        self.region_box.bind("<<ComboboxSelected>>", lambda e: self._pick_region())
        ttk.Label(f, text="一次只跑一个国家；两国共用一个紫鸟，不能同时跑。",
                  foreground="#666", wraplength=WRAP - 130, justify="left"
                  ).grid(row=0, column=2, columnspan=3, sticky="w")

        ttk.Label(f, text="会计月份").grid(row=1, column=0, sticky="w",
                                       pady=(10, 0))
        self.month = ttk.Combobox(f, values=recent_months(), width=12,
                                  state="readonly")
        self.month.set(recent_months()[0])
        self.month.grid(row=1, column=1, sticky="w", padx=(6, 18), pady=(10, 0))
        ttk.Label(f, text="默认是上一个完整月份。当月还没过完，做月报没有意义。",
                  foreground="#666", wraplength=WRAP - 130, justify="left"
                  ).grid(row=1, column=2, columnspan=3, sticky="w", pady=(10, 0))

        ttk.Label(f, text="店铺").grid(row=2, column=0, sticky="nw", pady=(10, 0))
        box = ttk.Frame(f)
        box.grid(row=2, column=1, columnspan=4, sticky="w", pady=(10, 0))
        self.all_stores = tk.BooleanVar(value=True)
        ttk.Checkbutton(box, text="全部店铺", variable=self.all_stores,
                        command=self._toggle_stores).pack(anchor="w")
        self.stores = []
        self.lb = tk.Listbox(box, selectmode="extended", height=4,
                             exportselection=False, width=24)
        self.lb.pack(side="left", pady=(4, 0))
        self.store_hint = tk.StringVar(value="")
        ttk.Label(box, textvariable=self.store_hint, foreground="#666",
                  justify="left", wraplength=WRAP - 190).pack(
            side="left", anchor="n", padx=12, pady=(4, 0))
        self._fill_stores()

        ttk.Label(f, text="报表输出到").grid(row=3, column=0, sticky="w",
                                        pady=(12, 0))
        self.outdir = tk.StringVar(
            value=load_settings().get("out_dir") or DEFAULT_REPORT_DIR)
        ttk.Entry(f, textvariable=self.outdir, width=34).grid(
            row=3, column=1, columnspan=2, sticky="w", padx=(6, 6), pady=(12, 0))
        ttk.Button(f, text="浏览…", command=self._pick_outdir, width=8).grid(
            row=3, column=3, sticky="w", pady=(12, 0))
        ttk.Button(f, text="恢复默认", command=self._reset_outdir, width=10).grid(
            row=3, column=4, sticky="w", padx=(6, 0), pady=(12, 0))
        ttk.Label(f, text="可以直接指向共享盘，财务就不用等人拷贝。目录不存在会自动建。"
                          "定时任务也写到这里。",
                  foreground="#666", wraplength=WRAP - 60, justify="left"
                  ).grid(row=4, column=1, columnspan=4, sticky="w", pady=(2, 0))

    def _fill_stores(self):
        """按所选国家重填店铺清单。换国家时选中的店全部作废 —— 不能把墨西哥
        的店带进智利那一轮。"""
        self.stores = store_list(self.region)
        self.lb.configure(state="normal")
        self.lb.delete(0, "end")
        for s in self.stores:
            self.lb.insert("end", s)
        self._toggle_stores()
        self.store_hint.set(
            "取消勾选“全部店铺”后，可按住 Ctrl 多选。\n"
            "%s的店铺来自 config.json，共 %d 家。\n"
            "只选部分店铺时不会出全店汇总表（避免覆盖真正的那份）。"
            % (region_choice.label(self.region), len(self.stores)))

    def _pick_region(self):
        region = region_choice.from_label(self.region_box.get())
        if region == self.region:
            return
        self.region = region
        save_settings(dict(load_settings(), region=region))
        self.root.title(region_choice.window_title(region))
        self._fill_stores()
        self._say("已切换到%s：%d 家店铺" % (region_choice.label(region),
                                        len(self.stores)))

    def _pick_outdir(self):
        d = filedialog.askdirectory(title="选择报表输出目录",
                                    initialdir=self.outdir.get() or ROOT)
        if d:
            self.outdir.set(os.path.normpath(d))
            save_settings(dict(load_settings(), out_dir=self.outdir.get()))

    def _reset_outdir(self):
        self.outdir.set(DEFAULT_REPORT_DIR)
        save_settings(dict(load_settings(), out_dir=DEFAULT_REPORT_DIR))

    def _out_dir(self):
        return (self.outdir.get() or "").strip() or DEFAULT_REPORT_DIR

    def _toggle_stores(self):
        self.lb.configure(state="disabled" if self.all_stores.get() else "normal")

    def _build_buttons(self):
        f = ttk.LabelFrame(self.left, text=" 操作 ", padding=10)
        f.pack(fill="x", padx=(12, 6), pady=6)
        # (按钮文字, 回调, 说明, 是否会起一轮运行)
        # 最后那个标志决定有一轮在跑时要不要禁用它。按标志而不是按下标：以前
        # 是 self.buttons[:4]，加到第五个按钮（补取）时没人改那个 4，于是任务
        # 跑着的时候"补取欠的报表"还能点，一点就把正在跑的紫鸟杀掉。
        rows = [
            ("① + ② 一次跑完", self.do_all,
             "先下载【上面所选店铺】的数据，下完接着出报表，中间不用管。"
             "每家店约 15 分钟（墨西哥 12 家约 3 小时）。平时就点这一个。", True),
            ("① 下载数据", self.do_download,
             "打开紫鸟，下载【上面所选店铺】的报表并清洗。每家约 15 分钟。", True),
            ("② 出报表", self.do_reports,
             "用已下载的数据，生成【上面所选店铺】的财务报表。约 1 分钟，"
             "不联网、不开浏览器，可以随便重跑。", True),
            ("补取欠的报表", self.do_collect,
             "MercadoPago 有些报表要生成几十分钟。这个只打开【上面所选国家】里"
             "确实欠着报表的店，不重新下载，与上面勾的店铺无关。", True),
            ("先看用哪批数据", self.do_dry_run,
             "不生成任何文件，只列出每家店会用哪个下载目录、缺不缺该月账单。"
             "秒出，任何时候都能点。", False),
            ("打开报表目录", self.do_open_dir,
             "在资源管理器里打开生成好的 .xlsx 所在目录。", False),
        ]
        self.buttons = []
        for i, (label, cmd, desc, spawns) in enumerate(rows):
            b = ttk.Button(f, text=label, command=cmd, width=16)
            b.grid(row=i, column=0, sticky="w", pady=2)
            ttk.Label(f, text=desc, foreground="#555", justify="left",
                      wraplength=WRAP - 150).grid(
                row=i, column=1, sticky="w", padx=12)
            if spawns:
                self.buttons.append(b)
            elif cmd == self.do_dry_run:
                self.preview_btn = b
        self.stop_btn = ttk.Button(f, text="中止", command=self.do_stop,
                                   width=16, state="disabled")
        self.stop_btn.grid(row=len(rows), column=0, sticky="w", pady=(8, 0))
        ttk.Label(f, text="结束正在跑的那一轮（定时任务起的也能停）。已下好的文件会保留。",
                  foreground="#555", wraplength=WRAP - 150, justify="left"
                  ).grid(row=len(rows), column=1, sticky="w", padx=12, pady=(8, 0))
        ttk.Label(f, foreground="#555",
                  text="想让它每月自动跑、不用人来点？双击 定时设置.cmd。"
                       "那边跑的和这里是同一个过程，只是多了个定时。",
                  wraplength=WRAP - 20, justify="left"
                  ).grid(row=len(rows) + 1, column=0, columnspan=2, sticky="w",
                         pady=(8, 0))

    def _build_human(self):
        """「需要人工处理」：脚本自己解决不了、要人去店铺里点一下的事。

        登录 / 验证、账号语言、店铺打不开、缺报表权限。这些放着不管就会每个
        月原样失败一次，而它们以前只埋在几千行日志中间。这里常驻显示，
        一轮跑完新出现时还会弹窗；处理完一键补跑。
        """
        self.human_frame = f = tk.LabelFrame(
            self.left, text=" 需要人工处理 ", padx=10, pady=8,
            font=("", 10, "bold"))
        f.pack(fill="x", padx=(12, 6), pady=6)
        self.human_lb = tk.Listbox(f, selectmode="extended", height=4,
                                   exportselection=False, width=60)
        self.human_lb.pack(fill="x")
        self.human_lb.bind("<<ListboxSelect>>", lambda e: self._show_human())
        self.human_how = tk.StringVar(value="")
        # 固定 4 行高：说明文字随选中项变长变短时，下面两个按钮不能跟着
        # 上下跳 —— 实测选中一家店后按钮上移，下一下就点空了。
        tk.Label(f, textvariable=self.human_how, foreground="#555",
                 justify="left", anchor="nw", height=4,
                 wraplength=WRAP).pack(anchor="w", fill="x", pady=(4, 4))
        row = ttk.Frame(f)
        row.pack(anchor="w")
        self.human_open = ttk.Button(row, text="打开所选店铺（去登录 / 处理）",
                                     command=self.do_open_stores)
        self.human_open.pack(side="left")
        self.human_rerun = ttk.Button(row, text="处理好了：重跑这些店铺并重出全部报表",
                                      command=self.do_rerun_human)
        self.human_rerun.pack(side="left", padx=(8, 0))

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
            self.human_lb.insert("end", region_choice.action_line(a))
        for i, a in enumerate(acts):
            if not keep or region_choice.action_line(a) in keep:
                self.human_lb.selection_set(i)
        stores = sorted(set(a["store"] for a in acts))
        if acts:
            self.human_frame.configure(
                text=" ⚠ 需要人工处理：%d 家店铺（%s） " % (len(stores),
                                                    self.human["month"] or ""),
                foreground=KIND_FG["bad"])
            self.human_lb.configure(height=min(6, max(2, len(acts))))
        else:
            self.human_frame.configure(text=" 需要人工处理：没有 ",
                                       foreground=KIND_FG["good"])
            self.human_lb.configure(height=1)
            self.human_how.set("最近一轮下载里，没有需要人去店铺里处理的事。")
        self._show_human()
        self._lock_human()
        newest = max((a.get("run_id") or "" for a in acts), default="")
        if announce and acts and newest and newest != self.human_told:
            self.human_told = newest
            messagebox.showwarning(
                "需要人工处理",
                "这一轮有 %d 家店铺需要人去处理，脚本自己解决不了：\n\n%s\n\n"
                "在左边「需要人工处理」里选中店铺，点「打开所选店铺」去处理；\n"
                "处理完点「重跑这些店铺并重出全部报表」。"
                % (len(stores), "\n".join(
                    "· %s —— %s" % (a["store"], a["title"]) for a in acts)))
        elif not announce:
            self.human_told = newest or None

    def _picked_human(self):
        acts = self.human["actions"]
        return [acts[i] for i in self.human_lb.curselection() if i < len(acts)]

    def _show_human(self):
        picked = self._picked_human()
        if not picked:
            if self.human["actions"]:
                self.human_how.set("在上面选中要处理的店铺（按住 Ctrl 可多选）。")
            return
        lines = []
        for a in picked:
            lines.append("%s：%s%s" % (a["store"], a["how"],
                                     "" if a["rerun"] else "（处理之前重跑没用）"))
        self.human_how.set("\n".join(lines))

    def _lock_human(self):
        busy = bool(self.busy)
        has = bool(self.human["actions"])
        st = "normal" if has and not busy else "disabled"
        self.human_open.configure(state=st)
        self.human_rerun.configure(state=st)

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
        self._say("正在打开 %s …（紫鸟窗口出来后，在里面完成登录 / 验证）"
                  % "、".join(stores), "run")
        self.human_open.configure(state="disabled")

        def body():
            kw = {}
            if os.name == "nt":
                kw["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW",
                                              0x08000000)
            failed = []
            for s in stores:
                try:
                    p = subprocess.run(
                        [interpreter(), "-u", os.path.join(ROOT, "open_store.py"),
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
        self._lock_human()
        if failed:
            self._say("✗ 没能打开：%s" % "、".join(failed), "bad")
        else:
            self._say("已打开 %s。在紫鸟窗口里处理好以后，回来点「重跑这些店铺并"
                      "重出全部报表」。" % "、".join(stores), "warn")

    def do_rerun_human(self):
        """只重新下载这几家店，然后给**全部**店铺重出报表（含全店汇总）。

        "全部店铺"是这几家店所在国家的全部店铺：区域按店铺算，和上面选的
        国家无关 —— 清单里两国的店都会出现。
        """
        plan = region_choice.rerun_plan(self._picked_human())
        stores, useless = plan["stores"], plan["useless"]
        if stores and not plan["region"]:
            messagebox.showinfo(
                "重跑", "所选的店铺分属不同的国家，一轮只能跑一个国家：\n\n%s\n\n"
                        "请只选同一个国家的店铺，分两次点。" % "\n".join(
                            "%s：%s" % (region_choice.label(r), "、".join(ss))
                            for r, ss in plan["groups"]))
            return
        if not stores:
            messagebox.showinfo(
                "重跑", "所选的店铺都是重跑解决不了的问题：\n%s"
                        % "\n".join(useless) if useless
                else "先在列表里选中要重跑的店铺。")
            return
        month = self.human["month"] or self.month.get()
        note = ("\n\n这几项重跑也解决不了，先跳过：\n%s" % "\n".join(useless)
                if useless else "")
        if not messagebox.askokcancel(
                "重跑并重出报表",
                "会重新下载 %s 的 %s 数据（每家约 15 分钟），\n"
                "下完给%s全部店铺重出报表，并重出全店汇总表。%s\n\n开始？"
                % ("、".join(stores), month,
                   region_choice.label(plan["region"]), note)):
            return
        self.start("both", "重跑 %s 并重出全部报表" % "、".join(stores), stores,
                   month=month, reports_for="all", region=plan["region"])

    def _build_log(self):
        f = ttk.LabelFrame(self.right, text=" 运行日志（手动、定时共用） ", padding=6)
        f.pack(fill="both", expand=True, padx=(6, 12), pady=(12, 6))
        self.view = LogView(f, on_state=self._on_run_state, who="启动窗口",
                            height=12)
        self.view.pack(fill="both", expand=True)

    def _build_status(self):
        self.status = tk.StringVar(value="就绪")
        bar = ttk.Frame(self.root)
        bar.pack(side="bottom", fill="x", padx=12, pady=(0, 12))
        self.status_lbl = ttk.Label(bar, textvariable=self.status,
                                    font=("", 10, "bold"))
        self.status_lbl.pack(side="left")
        ttk.Label(bar, text="解释器：%s" % interpreter(),
                  foreground="#888").pack(side="right")

    def _say(self, text, kind=None):
        self.status.set(text)
        self.status_lbl.configure(foreground=KIND_FG.get(kind, ""))

    # ---------------------------------------------------------------- 参数

    def _picked_stores(self):
        """返回 (能不能继续, 店铺子集或 None)。None = 全部店铺。"""
        if self.all_stores.get():
            return True, None
        picked = [self.stores[i] for i in self.lb.curselection()]
        if not picked:
            # 以前这种情况会悄悄按"全部店铺"跑，同时又不出汇总表 —— 两头不靠。
            messagebox.showwarning(
                "没选店铺", "取消了“全部店铺”，但列表里一家都没选。\n"
                            "按住 Ctrl 点选要跑的店，或者重新勾上“全部店铺”。")
            return False, None
        return True, picked

    def _confirm_month(self):
        if self.month.get() == recent_months()[0]:
            return True
        return messagebox.askokcancel(
            "补做历史月份",
            "你选的是 %s，不是最近一个月。\n\n"
            "脚本会去那个月的账单明细页取数，下载目录带 _m%s 后缀。\n\n继续？"
            % (self.month.get(), self.month.get().replace("-", "")))

    # ---------------------------------------------------------------- 动作

    def do_download(self):
        ok, stores = self._picked_stores()
        if ok and self._confirm_month():
            self.start("download", "下载数据", stores)

    def do_reports(self):
        ok, stores = self._picked_stores()
        if ok:
            self.start("report", "出报表", stores)

    def do_all(self):
        """下载 + 出报表，一次跑完，中间不需要人点。

        两步怎么衔接（下载硬失败就不出报表）由 run_monthly.py 决定，这里
        只是告诉它 --step both。
        """
        ok, stores = self._picked_stores()
        if not ok or not self._confirm_month():
            return
        who = ("%s全部 %d 家店铺" % (region_choice.label(self.region),
                                    len(self.stores)) if stores is None
               else "所选的 %d 家店铺" % len(stores))
        n_run = len(self.stores) if stores is None else len(stores)
        if not messagebox.askokcancel(
                "一次跑完",
                "会先下载 %s 的 %s 数据，下完接着出报表。\n\n"
                "每家店约 15 分钟，%d 家大约 %s，中间不需要你操作。\n"
                "期间这台机器会反复开关紫鸟浏览器 —— 请不要手动去点紫鸟，\n"
                "那会打断下载。\n\n"
                "这个窗口可以关掉，不影响后台继续跑。\n\n开始？"
                % (who, self.month.get(), n_run, run_hours(n_run))):
            return
        self.start("both", "下载 + 出报表", stores)

    def do_collect(self):
        self.start("collect", "补取欠的报表", None)

    def start(self, step, title, stores, month=None, reports_for=None,
              region=None):
        """起一轮：run_monthly.py --entrance manual ...，独立的后台进程。

        region 不给就是上面选的国家；补跑「需要人工处理」的店时按店铺的国家传。
        """
        region = region or self.region
        cur = runlog.active_run()
        if cur or self.launch:
            messagebox.showwarning(
                "正在运行",
                "已经有一轮在跑了（入口：%s）。\n请等它结束，或先点“中止”。"
                % runlog.ENTRANCE_ZH.get((cur or {}).get("entrance"), "手动"))
            return
        # 手输（而非"浏览…"选）的目录也先记下来：定时任务读的是这份设置。
        save_settings(dict(load_settings(), out_dir=self._out_dir()))
        argv = [interpreter(), "-u", RUNNER] + runlog.runner_argv(
            "manual", step, month or self.month.get(), stores, self._out_dir(),
            reports_for=reports_for, region=region)
        try:
            proc = runlog.spawn_detached(
                argv, cwd=ROOT, stderr_path=runlog.launch_error_path("manual"))
        except Exception as e:
            self._say("✗ 启动失败：%s" % e, "bad")
            return
        if region != region_choice.DEFAULT:
            title = "%s：%s" % (region_choice.label(region), title)
        self.launch = {"proc": proc, "at": datetime.datetime.now(),
                       "title": title}
        self.view.follow_current()
        self._say("正在启动：%s …（几秒后下面开始出日志）" % title, "run")
        self._lock(True)

    def do_stop(self):
        cur = runlog.active_run()
        if not cur:
            return
        if not messagebox.askokcancel(
                "中止",
                "确定要结束正在跑的这一轮吗？（入口：%s）\n\n"
                "会连同它打开的紫鸟和浏览器一起关掉。\n"
                "已经下好的文件会保留；中止后不会接着出报表。"
                % runlog.ENTRANCE_ZH.get(cur.get("entrance"), "?")):
            return
        try:
            runlog.request_stop(cur["run_id"], "启动窗口")
            self._say("正在中止…（几秒内生效）", "warn")
        except Exception as e:
            self._say("✗ 中止请求写不进去：%s" % e, "bad")

    def do_dry_run(self):
        """只读预览，所以**直接**调 run_reports.py --dry-run，不经过运行器。

        它不下载、不写任何文件，不需要锁，也不该在 logs/runs 里留一条"运行
        记录"；有一轮正在跑的时候也能点。结果显示在一个单独的小窗口里。
        """
        ok, stores = self._picked_stores()
        if not ok:
            return
        cmd = region_choice.preview_cmd(
            interpreter(), os.path.join(ROOT, "run_reports.py"),
            self.month.get(), stores, self.region)
        box = {}

        def body():
            kw = {}
            if os.name == "nt":
                # 本窗口是 pythonw 起的，没有控制台；不加这个会闪一个黑窗口。
                kw["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW",
                                              0x08000000)
            try:
                p = subprocess.run(
                    cmd, cwd=ROOT, stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    timeout=300, env=dict(os.environ, PYTHONIOENCODING="utf-8",
                                          PYTHONUTF8="1"), **kw)
                box["out"] = (p.stdout or b"").decode("utf-8", "replace")
            except Exception as e:
                box["out"] = "没跑成：%s: %s" % (e.__class__.__name__, e)

        def wait():
            if "out" not in box:
                self.root.after(150, wait)
                return
            self.preview_btn.configure(state="normal")
            self._show_preview("先看用哪批数据 —— %s %s"
                               % (region_choice.label(self.region),
                                  self.month.get()),
                               box["out"])

        self.preview_btn.configure(state="disabled")
        threading.Thread(target=body, daemon=True).start()
        wait()

    def _show_preview(self, title, text):
        w = tk.Toplevel(self.root)
        w.title(title)
        w.geometry("900x520")
        t = tk.Text(w, wrap="none", font=("Consolas", 9))
        sb = ttk.Scrollbar(w, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        t.pack(side="left", fill="both", expand=True)
        t.tag_config(runlog.WARN, foreground="#9a6700")
        t.tag_config(runlog.ERROR, foreground="#b42318")
        cls = runlog.LineClassifier("report", self.stores)
        for line in (text or "").splitlines():
            got = cls.feed(line)
            lv = got[2] if got else runlog.INFO
            t.insert("end", line + "\n", lv if lv != runlog.INFO else ())
        t.configure(state="disabled")

    def do_open_dir(self):
        d = self._out_dir()
        if not os.path.isdir(d):
            messagebox.showinfo("还没有报表",
                                "目录还不存在：\n%s\n\n先点“出报表”。" % d)
            return
        if os.name == "nt":
            os.startfile(d)                   # noqa: S606
        else:
            subprocess.Popen(["xdg-open", d])

    # ---------------------------------------------------------------- 状态

    def _watch(self):
        """每秒看一眼：有没有一轮在跑（任何入口起的），刚起的那轮怎么样了。"""
        try:
            cur = runlog.active_run()
            if self.launch:
                self._check_launch(cur)
            self._lock(bool(cur) or bool(self.launch))
            # 清单不常变：空闲时每 30 秒重读一次（别的窗口、定时任务跑完也能反映出来）
            now = datetime.datetime.now().timestamp()
            if not self.busy and now - self.human_at > 30:
                self.human_at = now
                self._refresh_human()
        except Exception:
            pass
        self.root.after(800, self._watch)

    def _check_launch(self, cur):
        proc, at = self.launch["proc"], self.launch["at"]
        if proc.poll() is None:
            if cur:
                self.launch = None        # 日志已经出现，交给查看器
            return
        # 进程已经结束。找它留下的那条记录：入口是手动、开始时间不早于我们
        # 起它的时刻。不能按 PID 找 —— Windows 上 .venv 里的 python.exe 只是
        # 个转发器，Popen 拿到的 PID 和脚本里 os.getpid() 不是同一个。
        self.launch = None
        since = at.replace(microsecond=0) - datetime.timedelta(seconds=1)
        mine = [r for r in runlog.list_runs(limit=5)
                if r["entrance"] == "manual" and r["when"] >= since]
        if mine:
            self._on_run_state(runlog.run_state(mine[0]["run_id"]))
            return
        # 一行日志都没写就退出了：解释器或依赖有问题。stderr 是唯一的线索。
        err = ""
        try:
            with io.open(runlog.launch_error_path("manual"), "rb") as f:
                err = f.read().decode("utf-8", "replace").strip()
        except Exception:
            pass
        self._say("✗ 没能启动（退出码 %s）" % proc.returncode, "bad")
        messagebox.showerror(
            "没能启动",
            "运行器一启动就退出了（退出码 %s），没有写出日志。\n\n%s"
            % (proc.returncode, err[-1500:] or "（没有任何错误输出）"))

    def _on_run_state(self, st):
        """查看器所看的那一轮状态变了。状态栏必须把结果说清楚。

        退出码含义以 runlog.EXIT_MEANING 为准（和 README、契约是同一份）。
        1 不是"失败"，是"有事要看" —— 报表已经写出来了，只是有校验没过，
        这恰恰是最该让人看见的情况。
        """
        if not st:
            return
        who = "%s，%s%s" % (runlog.ENTRANCE_ZH.get(st.get("entrance"), "?"),
                           runlog.region_tag(st), st.get("month") or "")
        if st["state"] == "running" and st.get("stalled"):
            self._say("▲ 这一轮（%s）%s" % (who, st["text"]), "bad")
        elif st["state"] == "running":
            self._say("正在运行（%s）…　可以关掉这个窗口，不影响它继续跑" % who,
                      "run")
        else:
            self._say("%s 最近一轮（%s）：%s"
                      % (KIND_MARK.get(st.get("kind"), "■"), who, st["text"]),
                      st.get("kind"))
            self.human_at = datetime.datetime.now().timestamp()
            self._refresh_human()

    def _lock(self, running):
        if running == self.busy:
            return
        self.busy = running
        for b in self.buttons:
            b.configure(state="disabled" if running else "normal")
        self.stop_btn.configure(state="normal" if running else "disabled")
        self._lock_human()

    def _on_close(self):
        # 手输（而非"浏览…"选）的目录也记下来
        save_settings(dict(load_settings(), out_dir=self._out_dir(),
                           region=self.region))
        if (self.launch or runlog.active_run()) and not messagebox.askokcancel(
                "还在运行",
                "有一轮还在跑。关掉这个窗口**不会**停它，它会在后台跑完，\n"
                "日志照常写，下次打开窗口还能接着看。\n\n"
                "想停它请点“取消”，再点“中止”。\n\n关闭窗口？"):
            return
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if os.name == "nt" else "clam")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
