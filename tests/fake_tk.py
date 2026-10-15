# -*- coding: utf-8 -*-
"""一个够用的假 tkinter：让 ml_app.py / log_view.py 在没有 tkinter 的机器上
把窗口"搭"起来，按钮回调能点 —— 看的是 Python 层面有没有写错（拼错的属性、
传错的参数、少了的分支），不是长得好不好看。长相要在 Windows 上看截图。

install() 把它塞进 sys.modules（tkinter / tkinter.ttk / .font / .messagebox /
.filedialog），uninstall() 拿掉。
"""
import sys
import types


class TclError(Exception):
    pass


class _Var(object):
    _default = ""

    def __init__(self, master=None, value=None):
        self._v = self._default if value is None else value
        self._traces = []

    def get(self):
        return self._v

    def set(self, v):
        self._v = v
        for cb in self._traces:
            cb()

    def trace_add(self, mode, cb):
        self._traces.append(cb)


class StringVar(_Var):
    _default = ""


class IntVar(_Var):
    _default = 0

    def get(self):
        return int(self._v)


class BooleanVar(_Var):
    _default = False

    def get(self):
        return bool(self._v)


ROOTS = []


class Widget(object):
    """所有控件的底子：记住配置，未知方法一律什么都不做。"""

    def __init__(self, master=None, *args, **kw):
        self.master = master
        self.cfg = dict(kw)
        self.children_ = []
        self.bindings = {}
        self.mapped = False
        if master is not None and hasattr(master, "children_"):
            master.children_.append(self)
        tv = kw.get("textvariable")
        self.textvariable = tv

    # 配置
    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        self.cfg.update(kw)
        return None
    config = configure

    def cget(self, key):
        return self.cfg.get(key, "")

    def __setitem__(self, k, v):
        self.cfg[k] = v

    def __getitem__(self, k):
        return self.cfg.get(k)

    # 布局
    def pack(self, **kw):
        self.mapped = True

    def grid(self, **kw):
        self.mapped = True

    def pack_forget(self):
        self.mapped = False

    def winfo_ismapped(self):
        return self.mapped

    def winfo_children(self):
        return list(self.children_)

    def destroy(self):
        if self.master is not None and self in getattr(self.master, "children_", []):
            self.master.children_.remove(self)

    def columnconfigure(self, *a, **k):
        pass

    def winfo_width(self):
        return 800

    def winfo_height(self):
        return 600

    def winfo_reqheight(self):
        return 400

    def winfo_screenwidth(self):
        return 1920

    def winfo_screenheight(self):
        return 1080

    # 事件
    def bind(self, seq, fn=None, add=None):
        self.bindings[seq] = fn

    def bind_all(self, seq, fn=None, add=None):
        pass

    def unbind_all(self, seq):
        pass

    def after(self, ms, fn=None, *args):
        root = _root_of(self)
        root.pending.append((fn, args))
        return "after#%d" % len(root.pending)

    def after_cancel(self, job):
        pass

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        return lambda *a, **k: None


def _root_of(w):
    while getattr(w, "master", None) is not None:
        w = w.master
    return w


class Tk(Widget):
    def __init__(self, *a, **k):
        Widget.__init__(self, None)
        self.pending = []
        self.titles = []
        self.protocols = {}
        self.destroyed = False
        self.tk = types.SimpleNamespace(call=lambda *a: None)
        ROOTS.append(self)

    def title(self, t=None):
        if t is not None:
            self.titles.append(t)
        return self.titles[-1] if self.titles else ""

    def protocol(self, name, fn):
        self.protocols[name] = fn

    def destroy(self):
        self.destroyed = True

    def pump(self, rounds=3):
        """把排队的 after 回调跑几轮（每个回调可能再排下一次）。"""
        for _ in range(rounds):
            todo, self.pending = self.pending, []
            for fn, args in todo:
                if fn:
                    fn(*args)


class Toplevel(Widget):
    def title(self, t=None):
        pass


class Frame(Widget):
    pass


class Label(Widget):
    pass


class Canvas(Widget):
    def create_window(self, *a, **k):
        return 1

    def bbox(self, *a):
        return (0, 0, 10, 10)


class Scrollbar(Widget):
    pass


class Text(Widget):
    def __init__(self, master=None, **kw):
        Widget.__init__(self, master, **kw)
        self.content = []

    def insert(self, index, *chunks):
        # chars, tags, chars, tags ...
        self.content.extend(c for i, c in enumerate(chunks) if i % 2 == 0)

    def delete(self, a, b=None):
        self.content = []

    def text(self):
        return "".join(self.content)


class Listbox(Widget):
    def __init__(self, master=None, **kw):
        Widget.__init__(self, master, **kw)
        self.items = []
        self.sel = set()

    def insert(self, index, *items):
        self.items.extend(items)

    def delete(self, a, b=None):
        self.items, self.sel = [], set()

    def get(self, i):
        return self.items[i]

    def size(self):
        return len(self.items)

    def curselection(self):
        return tuple(sorted(self.sel))

    def selection_set(self, i):
        self.sel.add(i)

    def selection_clear(self, a, b=None):
        self.sel = set()


tk = types.ModuleType("tkinter")
for _n, _v in list(globals().items()):
    if _n in ("TclError", "StringVar", "IntVar", "BooleanVar", "Tk", "Toplevel",
              "Frame", "Label", "Canvas", "Scrollbar", "Text", "Listbox"):
        setattr(tk, _n, _v)
tk.Widget = Widget


# ---------------------------------------------------------------- ttk

class Combobox(Widget):
    def __init__(self, master=None, **kw):
        Widget.__init__(self, master, **kw)
        self.value = ""
        if callable(kw.get("postcommand")):
            self.post = kw["postcommand"]

    def set(self, v):
        self.value = v

    def get(self):
        return self.value

    def cget(self, key):
        v = self.cfg.get(key, "")
        return tuple(v) if key == "values" else v


class Entry(Widget):
    pass


class Button(Widget):
    def invoke(self):
        if self.cfg.get("state") != "disabled" and self.cfg.get("command"):
            return self.cfg["command"]()


class Notebook(Widget):
    def __init__(self, master=None, **kw):
        Widget.__init__(self, master, **kw)
        self.tabs_ = []
        self.selected = None

    def add(self, child, **kw):
        self.tabs_.append([child, dict(kw)])

    def select(self, child=None):
        if child is None:
            return self.selected
        self.selected = child

    def tab(self, child, **kw):
        for t in self.tabs_:
            if t[0] is child:
                t[1].update(kw)
                return t[1]


class PanedWindow(Widget):
    def add(self, child, **kw):
        pass


class Style(object):
    def __init__(self, *a):
        pass

    def theme_use(self, *a):
        pass

    def configure(self, *a, **k):
        pass

    def lookup(self, *a):
        return "#f0f0f0"


ttk = types.ModuleType("tkinter.ttk")
for _n, _c in (("Frame", Frame), ("Label", Label), ("Button", Button),
               ("Checkbutton", Button), ("Combobox", Combobox), ("Entry", Entry),
               ("Spinbox", Entry), ("LabelFrame", Frame), ("Notebook", Notebook),
               ("PanedWindow", PanedWindow), ("Scrollbar", Scrollbar),
               ("Style", Style)):
    setattr(ttk, _n, _c)


class Font(object):
    def __init__(self, *a, **k):
        pass

    def measure(self, text):
        return 7 * len(text or "")


font = types.ModuleType("tkinter.font")
font.Font = Font

# messagebox：answers 里排好要回答什么（没排的：ask* 回 True，show* 记下来）
ANSWERS = []
SHOWN = []


def _ask(kind):
    def f(title, text, **k):
        SHOWN.append((kind, title, text))
        return ANSWERS.pop(0) if ANSWERS else True
    return f


def _show(kind):
    def f(title, text, **k):
        SHOWN.append((kind, title, text))
    return f


messagebox = types.ModuleType("tkinter.messagebox")
for _k in ("askokcancel", "askyesno", "askyesnocancel"):
    setattr(messagebox, _k, _ask(_k))
for _k in ("showinfo", "showwarning", "showerror"):
    setattr(messagebox, _k, _show(_k))

filedialog = types.ModuleType("tkinter.filedialog")
filedialog.askdirectory = lambda **k: ""

tk.ttk, tk.font, tk.messagebox, tk.filedialog = ttk, font, messagebox, filedialog

MODULES = {"tkinter": tk, "tkinter.ttk": ttk, "tkinter.font": font,
           "tkinter.messagebox": messagebox, "tkinter.filedialog": filedialog}


def install():
    saved = dict((k, sys.modules.get(k)) for k in MODULES)
    sys.modules.update(MODULES)
    for k in ("core.ml_app", "core.log_view"):
        sys.modules.pop(k, None)
    return saved


def uninstall(saved):
    for k, v in saved.items():
        if v is None:
            sys.modules.pop(k, None)
        else:
            sys.modules[k] = v
    for k in ("core.ml_app", "core.log_view"):
        sys.modules.pop(k, None)
