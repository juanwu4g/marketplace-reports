# -*- coding: utf-8 -*-
"""
Amazon monthly financial workbook — a sister of the other platforms'.

Ten sheets, the same look and the same check levels: the layout is written by
core/report_family.py (docs/REPORT_FAMILY.md §6); this module supplies the
content. Accounting rules: README.md ("Accounting rules") — agreed with finance
BEFORE they are coded (docs/NEW_PLATFORM_SOP.md §4).

    load_store(store, folder, month)   read the files of one download folder
    derive(S)                          S["m"]: the numbers, S["m"]["pl"][key] per P&L row
    validate(stores)                   check records (core/report_kit.check)
    build_workbook(stores, ctx, out)   the sheets, from core/report_family
    build(pairs, month, out)           what run_reports calls
"""
import datetime
import io
import json
import os

from openpyxl import Workbook

from core import report_family as RF
from core import report_kit as K
from platforms.amazon import periods
from platforms.amazon import site as platform_site

SHEETS = K.SHEET_NAMES
# False until load_store / derive / validate are written for the real files.
IMPLEMENTED = False

# ② fixed rows: key, 中文, the platform's own term, kind (d 取数 / in 录入 / sum 小计 /
# f 计算), note, formula (rows added, "-key" subtracted; RF.Span(a, b) sums a block).
# Replace the placeholder lines with the agreed ones; keep the order stable.
_R = RF.PLRow
PL_ROWS = [
    _R("sales", "商品销售收入", "TODO", "d", ""),
    _R("refund", "减：退货退款", "TODO", "d", ""),
    _R("net_sales", "净销售收入", "", "f", "＝ 商品销售收入 − 退货退款", ("sales", "refund")),
    _R("commission", "减：销售佣金", "TODO", "d", ""),
    _R("shipping", "减：物流费用", "TODO", "d", ""),
    _R("ads", "减：广告推广", "TODO", "d", ""),
    _R("other_fees", "减：其他费用", "TODO", "d", ""),
    _R("fee_total", "平台费用小计", "", "sum", "", RF.Span("commission", "other_fees")),
    _R("contrib", "经营贡献毛利（未扣商品成本）", "", "f", "＝ 净销售收入 − 平台费用小计",
       ("net_sales", "fee_total")),
    _R("cogs", "减：商品采购成本", "", "in", "⚠ 平台数据不含此项：在此填总额"),
    _R("profit", "税前净利润", "", "f", "＝ 经营贡献毛利 − 商品采购成本", ("contrib", "-cogs")),
]
DATA_KEYS = [x.key for x in PL_ROWS if x.kind == "d"]


def read_meta(folder):
    try:
        with io.open(os.path.join(folder, "run_meta.json"), encoding="utf-8") as fh:
            return json.load(fh) or {}
    except Exception:
        return {}


def load_store(store, folder, month):
    """{"store", "folder", "month", "meta", "files": {form: path}} plus the
    loaded tables. Drop customer personal data on load."""
    return {"store": store, "folder": folder, "month": month, "meta": read_meta(folder),
            "files": {}}


def derive(S):
    """Fill S["m"]: at least {"pl": {key: amount for every "d" row}}."""
    if not IMPLEMENTED:
        raise NotImplementedError("亚马逊 的报表还没实现（platforms/amazon/report_build.py："
                                  "load_store / derive / validate）")
    return S


def validate(stores):
    """One report_kit.check per thing finance must be able to trust (REPORT_FAMILY §4)."""
    return []


# ======================================================================= workbook

def build_workbook(stores, ctx, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    country = ctx.get("country") or "MX"
    fam = RF.Family(stores, money=K.money_format(platform_site.get(country)["decimals"]))
    bad = RF.failed_errors(ctx["checks"])
    fam.readme_sheet(
        wb, SHEETS[0], "亚马逊 月度财务报表 · %s · %s" % (ctx["period_cn"], ctx["who"]),
        "生成于 %s。" % ctx["built"], [
            ("一、这份报表回答什么", [
                ("本月赚了多少", "【%s】从商品销售收入推到经营贡献毛利。" % SHEETS[1]),
                ("钱对不对得上", "【%s】" % SHEETS[2]),
                ("还缺什么", "【%s】" % SHEETS[9])]),
            ("二、数据来源", [(S["store"], "下载于 %s" % (S["meta"].get("downloaded_at") or "?"))
                             for S in stores]),
        ],
        warning=("本期有 %d 项 error 级校验未通过，请先看【%s】。" % (len(bad), SHEETS[9])) if bad else None)
    sh, r, _rows_at, note_col = fam.pl_sheet(
        wb, SHEETS[1], PL_ROWS, value=lambda S, key: S["m"]["pl"].get(key, 0.0),
        title="合并损益表（%s）" % ctx["period_cn"], subtitle="口径：TODO")
    sales = lambda S: S["m"]["pl"].get("sales", 0.0)
    fam.ratio_block(sh, r + 1, "费用率（占商品销售收入）", [
        ("平台费用合计", lambda S: -sum(S["m"]["pl"].get(k, 0.0) for k in
                                   ("commission", "shipping", "ads", "other_fees")), sales)],
        note_col)
    fam.bridge_sheet(wb, SHEETS[2], "平账勾稽表（%s）" % ctx["period_cn"],
                     "每座桥最后一行是残差，应为 0。", [
                         ("桥 A　TODO", [RF.Line("TODO", lambda S: 0.0),
                                        RF.Line("残差（应为 0）", lambda S: 0.0, style="residual")])])
    for i in range(3, 9):
        RF.not_applicable_sheet(wb, SHEETS[i], SHEETS[i][2:],
                                "不适用：亚马逊 的这一页还没做（TODO）。")
    fam.gaps_sheet(wb, SHEETS[9], "数据缺口与待办（%s）" % ctx["period_cn"], ctx["checks"])
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    return out_path


# ======================================================================= entry

def build_context(stores, month, build_date=None):
    y, mo = periods.parse_month(month)
    return {"month": month, "period_cn": "%d 年 %d 月" % (y, mo),
            "who": stores[0]["store"] if len(stores) == 1 else "%d 家店铺汇总" % len(stores),
            "built": build_date or datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}


def build(pairs, month, out_path, strict=False, build_date=None):
    """pairs: [(store, folder)]. Returns a result dict (outcome, checks, out)."""
    stores = [derive(load_store(s, f, month)) for s, f in pairs]
    ctx = build_context(stores, month, build_date)
    ctx["checks"] = validate(stores)
    bad = RF.failed_errors(ctx["checks"])
    if strict and bad:
        return {"outcome": "strict_aborted", "checks": ctx["checks"], "out": out_path}
    build_workbook(stores, ctx, out_path)
    return {"outcome": "checks_failed" if bad else "ok", "checks": ctx["checks"], "out": out_path,
            "metrics": {S["store"]: S["m"] for S in stores}}
