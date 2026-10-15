# -*- coding: utf-8 -*-
"""窗口的字体、颜色、间距 —— 统一窗口和日志查看器用同一套。

只有常量，不 import tkinter（apply() 里才用到 Tk 对象），所以别的模块随便引。
"""
import os

WIN = os.name == "nt"

# 中文界面字体：Windows 上用微软雅黑 UI，别的平台交给 Tk 默认。
UI_FAMILY = "Microsoft YaHei UI" if WIN else "TkDefaultFont"
# 日志用等宽英文字体；汉字 Tk 会自动退到有字形的字体。列靠制表位对齐，
# 所以混排也不会错位。
MONO_FAMILY = "Consolas" if WIN else "TkFixedFont"

FONT = (UI_FAMILY, 10)
FONT_BOLD = (UI_FAMILY, 10, "bold")
FONT_SMALL = (UI_FAMILY, 9)
FONT_TITLE = (UI_FAMILY, 12, "bold")
FONT_BANNER = (UI_FAMILY, 11, "bold")
FONT_LOG = (MONO_FAMILY, 10)

# 间距：外边距 / 分组之间 / 控件之间
PAD = 12
GAP = 8
SMALL = 4

# 颜色。kind：run 在跑 / good 正常 / warn 要看一眼 / bad 出事
FG = {"run": "#0969da", "good": "#1a7f37", "warn": "#9a6700", "bad": "#b42318",
      None: "#1f2328", "": "#1f2328"}
BANNER_BG = {"run": "#ddf4ff", "good": "#dafbe1", "warn": "#fff8c5",
             "bad": "#ffebe9", None: "#f6f8fa", "": "#f6f8fa"}
HINT = "#57606a"           # 说明文字
MUTED = "#8c959f"

# 入口色块：一眼分清这是"手动"还是"定时"那一轮。
BADGE_BG = {"manual": "#0969da", "scheduled": "#8250df", "cli": "#57606a"}

# 日志颜色（浅底）
LOG_BG = "#ffffff"
LOG_FG = "#1f2328"
LOG_TIME = "#6e7781"
LOG_ERROR_FG, LOG_ERROR_BG = "#b42318", "#ffebe9"
LOG_WARN_FG, LOG_WARN_BG = "#7d4e00", "#fff8c5"
LOG_HEAD = "#0550ae"
LOG_BANNER_BG = "#eaeef2"


def apply(root):
    """给整个窗口设默认字体和 ttk 主题。失败不要紧，用系统默认。"""
    try:
        from tkinter import ttk
        style = ttk.Style(root)
        style.theme_use("vista" if WIN else "clam")
        style.configure(".", font=FONT)
        style.configure("TNotebook.Tab", padding=(14, 6), font=FONT)
        style.configure("Hint.TLabel", foreground=HINT, font=FONT_SMALL)
        style.configure("Title.TLabel", font=FONT_TITLE)
        style.configure("Section.TLabel", font=FONT_BOLD)
        style.configure("Accent.TButton", font=FONT_BOLD, padding=(12, 6))
        style.configure("TButton", padding=(8, 4))
        style.configure("TLabelframe.Label", font=FONT_BOLD)
        for kind, color in FG.items():
            if kind:
                style.configure("%s.TLabel" % kind.capitalize(), foreground=color)
    except Exception:
        pass
    try:
        root.option_add("*Font", FONT)
        root.option_add("*Listbox.font", FONT)
    except Exception:
        pass
