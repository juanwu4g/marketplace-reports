# -*- coding: utf-8 -*-
"""范围选择：平台 → 国家 → 店铺 的勾选树（「立即运行」和「每月自动运行」共用）。

只管摆和勾；树长什么样、勾了之后跑什么都在 core/scope.py。ttk.Treeview 没有
自带的勾选框：行首画一个小方框（全勾 / 部分 / 没勾三种图），点一行就切换，点左边的
展开标记只展开 / 收起。方框是画出来的图，不用 ☑ ☐ 这类字符 —— Windows 上这几个
字符来自不同的字体，大小差很多，没勾的那个小得几乎看不见。
"""
import tkinter as tk
from tkinter import ttk

from core import scope as SC
from core import ui_style as S

STATES = ("all", "some", "none")
BOX = 14                                  # 方框的边长（像素）
BLUE, GREY, WHITE = "#0969da", "#8c959f", "#ffffff"


def _box_images():
    """三张 14×14 的小图：全勾（蓝底白勾）、部分（白底蓝块）、没勾（白底灰框）。"""
    imgs = {}
    for st in STATES:
        img = tk.PhotoImage(width=BOX, height=BOX)
        edge, inside = (BLUE, BLUE) if st == "all" else (GREY, WHITE)
        img.put(edge, to=(0, 0, BOX, BOX))
        img.put(inside, to=(1, 1, BOX - 1, BOX - 1))
        if st == "some":
            img.put(BLUE, to=(4, 4, BOX - 4, BOX - 4))
        elif st == "all":
            for x, y in ((3, 7), (4, 8), (5, 9), (6, 8), (7, 7), (8, 6), (9, 5), (10, 4)):
                img.put(WHITE, to=(x, y, x + 1, y + 2))
        imgs[st] = img
    return imgs


class ScopePicker(object):
    """with_stores=False：只勾到国家一层（计划任务按整个国家跑，不挑店）。
    fit=True：树有几行就多高（最多 height 行），不用滚 —— 国家一层的行数少。"""

    def __init__(self, parent, with_stores=True, on_change=None, height=9, fit=False):
        self.with_stores = with_stores
        self.max_height, self.fit = height, fit
        self.on_change = on_change
        self.t = []
        self.on = {}                 # 区域 -> 勾着的店（set）；空 = 这个区域没勾
        self.results = {}
        self.marks = {}              # iid -> all / some / none（测试和读屏用）
        self.images = _box_images()  # 要留着引用，不然 Tk 会把图回收掉
        self.frame = ttk.Frame(parent)
        self.tv = ttk.Treeview(self.frame, columns=("n", "last"), show="tree headings",
                               height=height, selectmode="none")
        self.tv.heading("#0", text="平台 / 国家" + (" / 店铺" if with_stores else ""),
                        anchor="w")
        self.tv.heading("n", text="店铺", anchor="e")
        self.tv.heading("last", text="上次运行", anchor="w")
        self.tv.column("#0", width=232, stretch=True)
        self.tv.column("n", width=56, anchor="e", stretch=False)
        self.tv.column("last", width=178, anchor="w", stretch=False)
        sb = ttk.Scrollbar(self.frame, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tv.tag_configure("platform", font=S.FONT_BOLD)
        self.tv.tag_configure("off", foreground=S.HINT)
        self.tv.bind("<Button-1>", self._click)
        self.tv.bind("<space>", lambda e: self.toggle(self.tv.focus()))

    def pack(self, **kw):
        self.frame.pack(**kw)

    def grid(self, **kw):
        self.frame.grid(**kw)

    # ---------------------------------------------------------------- 内容

    def load(self, t, sel, results=None):
        """t：scope.tree()；sel：scope 的勾选。重建整棵树。"""
        self.t = t
        sel = SC.normalize(sel, t)
        self.on = {}
        for u in SC.units(t):
            picked = sel.get(u.unit, False)
            self.on[u.unit] = (set(u.stores) if picked is None else
                               set(picked) if picked else set())
        if results is not None:
            self.results = results
        for iid in self.tv.get_children(""):
            self.tv.delete(iid)
        for p in t:
            self.tv.insert("", "end", iid="p:" + p.code, open=True, tags=("platform",))
            for u in p.units:
                part = 0 < len(self.on[u.unit]) < len(u.stores)
                self.tv.insert("p:" + p.code, "end", iid="u:" + u.unit,
                               open=part and self.with_stores)
                if self.with_stores:
                    for s in u.stores:
                        self.tv.insert("u:" + u.unit, "end", iid="s:%s:%s" % (u.unit, s))
        if self.fit:
            rows = len(t) + len(SC.units(t))
            self.tv.configure(height=max(3, min(self.max_height, rows)))
        self._paint()

    def set_results(self, results):
        self.results = results or {}
        self._paint()

    def selection(self):
        """scope 的勾选：{区域: None（全部）或 [店]}。"""
        out = {}
        for u in SC.units(self.t):
            got = self.on.get(u.unit) or set()
            if got:
                out[u.unit] = None if len(got) == len(u.stores) else [
                    s for s in u.stores if s in got]
        return SC.normalize(out, self.t)

    def set_selection(self, sel):
        self.load(self.t, sel)

    # ---------------------------------------------------------------- 勾

    def _state(self, have, total):
        return "none" if not have else "all" if have >= total else "some"

    def toggle(self, iid):
        """点一行：没全勾 -> 全勾；全勾了 -> 全不勾。平台、国家一层管下面的全部。"""
        if not iid:
            return
        kind, _, rest = iid.partition(":")
        if kind == "s":
            unit, _, store = rest.partition(":")
            got = self.on[unit]
            got.symmetric_difference_update({store})
        else:
            us = ([SC.find(self.t, rest)] if kind == "u" else
                  [u for p in self.t if p.code == rest for u in p.units])
            have = sum(len(self.on[u.unit]) for u in us)
            total = sum(len(u.stores) for u in us)
            turn_on = have < total
            for u in us:
                self.on[u.unit] = set(u.stores) if turn_on else set()
        self._paint()
        if self.on_change:
            self.on_change()

    def set_all(self, on=True):
        for u in SC.units(self.t):
            self.on[u.unit] = set(u.stores) if on else set()
        self._paint()
        if self.on_change:
            self.on_change()

    def _click(self, e):
        iid = self.tv.identify_row(e.y)
        if not iid or "indicator" in str(self.tv.identify_element(e.x, e.y)):
            return None                  # 三角：交给 Treeview 自己展开 / 收起
        self.tv.focus(iid)
        self.toggle(iid)
        return "break"

    # ---------------------------------------------------------------- 画

    def _paint(self):
        for p in self.t:
            have = sum(len(self.on[u.unit]) for u in p.units)
            total = sum(len(u.stores) for u in p.units)
            st = self._state(have, total)
            self._mark("p:" + p.code, st, p.name, values=(self._count(have, total), ""),
                       tags=("platform",) if st != "none" else ("platform", "off"))
            for u in p.units:
                got = self.on[u.unit]
                st = self._state(len(got), len(u.stores))
                self._mark("u:" + u.unit, st, u.name,
                           values=(self._count(len(got), len(u.stores)),
                                   SC.unit_status(self.results, u.unit)),
                           tags=() if st != "none" else ("off",))
                if self.with_stores:
                    for s in u.stores:
                        on = s in got
                        self._mark("s:%s:%s" % (u.unit, s), "all" if on else "none", s,
                                   values=("", ""), tags=() if on else ("off",))

    def _mark(self, iid, st, text, values, tags):
        self.marks[iid] = st
        self.tv.item(iid, text=" " + text, image=self.images[st], values=values, tags=tags)

    @staticmethod
    def _count(have, total):
        return "%d 家" % total if have in (0, total) else "%d/%d 家" % (have, total)
