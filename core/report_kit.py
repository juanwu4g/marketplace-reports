# -*- coding: utf-8 -*-
"""
The report family's frame, shared by every platform's workbook
(docs/REPORT_FAMILY.md): look, sheet scaffolding, check records, names.

What lives here is what makes two platforms' workbooks recognisable as sisters;
what they *say* stays in each platform's report_build. MercadoLibre's
report_build imports these instead of defining its own — the values were moved,
not changed, and the golden tests prove it.

Only openpyxl; no pandas, no project imports: a report builder that runs as a
plain script can import it after putting the repository root on sys.path.
"""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ sheets

# Ten sheets, always, in this order. A sheet that does not apply to a platform
# says so ("不适用：…") instead of disappearing: people learn where things are.
SHEET_NAMES = [
    "① 阅读说明", "② 合并损益表", "③ 平账勾稽表", "④ 平台费用明细",
    "⑤ 退货退款分析", "⑥ 退货质检明细", "⑦ SKU盈利分析", "⑧ 月度趋势",
    "⑨ 库存与动销分析", "⑩ 数据缺口与待办",
]

# ------------------------------------------------------------------ look

FONT_NAME = "Arial"
MNY = '#,##0.00;(#,##0.00);-'
MNY0 = '#,##0;(#,##0);-'
INT = '#,##0;(#,##0);-'
PCT = '0.0%;(0.0%);-'
PCT2 = '0.00%'
PCT3 = '0.000%'

C_NAVY, C_BAND, C_SUB, C_CALC = "1F3864", "4472C4", "D9E2F3", "EDEDED"
C_TOTAL, C_INPUT, C_OK, C_BAD, C_WARN = "FCE4D6", "FFFF00", "E2EFDA", "FBE5E5", "FFF2CC"
C_RED, C_GREEN, C_AMBER, C_GREY = "C00000", "375623", "BF8F00", "595959"

_THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def money_format(decimals):
    """Number format for money in a currency with this many decimals (CLP: 0)."""
    return MNY0 if not decimals else MNY


def F(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, italic=italic)


def fill(hexcolor):
    return PatternFill("solid", fgColor=hexcolor)


def wrap(vertical="center"):
    return Alignment(wrap_text=True, vertical=vertical)


def est_height(text, per_line=60, line_h=15, minimum=15):
    """按字符数粗估行高。中文宽字符，60 字/行是保守值。"""
    if not text:
        return minimum
    return max(minimum, line_h * (len(str(text)) // per_line + 1))


class Sheet(object):
    """薄封装：多店铺时列数随店铺数变化，这里统一算列位。"""

    def __init__(self, wb, name, first_store_col, nstores, has_total=True):
        self.ws = wb.create_sheet(name)
        self.c0 = first_store_col
        self.n = nstores
        self.has_total = has_total

    def scol(self, i):
        return get_column_letter(self.c0 + i)

    @property
    def tcol_idx(self):
        return self.c0 + self.n

    @property
    def tcol(self):
        return get_column_letter(self.tcol_idx)

    def title(self, text, sub="", sub2=""):
        title_rows(self.ws, text, sub, sub2)

    def header(self, row, cols, widths, height=30, color=C_NAVY):
        ws = self.ws
        header_cells(ws, row, cols, color=color)
        ws.row_dimensions[row].height = height
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def title_rows(ws, text, sub="", sub2=""):
    """A sheet's head: A1 the title (navy, bold, row 24 high), A2 / A3 grey notes
    when given (Sheet.title on any worksheet)."""
    ws["A1"] = text
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F(9, False, C_GREY)
    if sub2:
        ws["A3"] = sub2
        ws["A3"].font = F(9, False, C_GREY)


def header_cells(ws, row, cols, first_col=1, color=C_NAVY, wrap_text=True):
    """The cells of a table header: white bold text on `color`, centred, boxed —
    one per entry of cols, from column first_col. Row height and column widths
    are left to the caller (Sheet.header sets both). wrap_text=False leaves the
    text on one line (the alignment then carries no wrap flag at all)."""
    if wrap_text:
        align = dict(horizontal="center", vertical="center", wrap_text=True)
    else:
        align = dict(horizontal="center", vertical="center")
    for i, t in enumerate(cols, first_col):
        c = ws.cell(row=row, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(color)
        c.alignment = Alignment(**align)
        c.border = BOX


# ------------------------------------------------------------------ checks

def check(key, store, name, ok, detail, severity="error", skipped=False):
    """One check record (REPORT_FAMILY §4): error = numbers not trustworthy until
    fixed; warn = look at this; skipped = not applicable here, with the reason."""
    return {"key": key, "store": store, "name": name, "ok": bool(ok),
            "detail": detail, "severity": severity, "skipped": skipped}


# ------------------------------------------------------------------ names

def store_workbook_name(store, month):
    return "%s_%s.xlsx" % (store, month)


def combined_workbook_name(country, platform_code, month):
    """CL_FA_全店汇总_2026-09.xlsx (MercadoLibre keeps its historical CL_ML_ / MX_ML_)."""
    return "%s_%s_全店汇总_%s.xlsx" % (str(country).upper(), str(platform_code).upper(), month)


def alert_file_name(country, platform_code, month):
    """！报表异常_CL_FA_202609.txt"""
    return "！报表异常_%s_%s_%s.txt" % (str(country).upper(), str(platform_code).upper(),
                                     str(month).replace("-", ""))


# ------------------------------------------------------------------ recalculation

def recalc(path, timeout=180):
    """用 LibreOffice 重算并回填公式缓存值。openpyxl 只写公式不写结果，
    不重算的话 pandas / data_only=True 读到的全是 None。"""
    import os
    import subprocess
    out_dir = os.path.join(os.path.dirname(os.path.abspath(path)), "_recalc")
    cmd = ["soffice", "--headless", "--norestore", "--convert-to",
           "xlsx:Calc MS Excel 2007 XML", "--outdir", out_dir, path]
    try:
        subprocess.call(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except OSError:
        return False, "找不到 soffice，跳过重算（公式仍然正确，Excel 打开时会自动算）"
    produced = os.path.join(out_dir, os.path.basename(path))
    if not os.path.exists(produced):
        return False, "LibreOffice 未产出文件，跳过重算"
    try:
        os.remove(path)
    except OSError:
        pass
    os.rename(produced, path)
    try:
        os.rmdir(out_dir)
    except OSError:
        pass
    return True, "已重算"


def scan_formula_errors(path):
    """重算后扫一遍，确保没有 #REF! / #NAME? 之类烂在文件里。"""
    ERR = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "Err:")
    from openpyxl import load_workbook
    wbv = load_workbook(path, data_only=True)
    wbf = load_workbook(path)
    n_formula = n_err = n_none = 0
    bad = []
    for sn in wbf.sheetnames:
        sf, sv = wbf[sn], wbv[sn]
        for row in sf.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n_formula += 1
                    v = sv[c.coordinate].value
                    if isinstance(v, str) and any(e in v for e in ERR):
                        n_err += 1
                        bad.append("%s!%s = %s" % (sn, c.coordinate, v))
                    elif v is None:
                        n_none += 1
    return n_formula, n_err, n_none, bad[:20]
