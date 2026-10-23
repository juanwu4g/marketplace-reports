# -*- coding: utf-8 -*-
"""Generate period workbooks: names, folders, the previous file, alerts, and
the automatic generation after a monthly run.

    generate(period, market, …)   one market's workbook -> a result dict
    auto_periods(month)           which periods a monthly run of `month` closes
    alert_text / write_alert      the 「！期间报表异常」 file of an automatic run
"""
import datetime
import io
import os
import traceback

from core import layout
from core import markets
from core import results as R
from core.period_report import model as MD
from core.period_report import workbook as WB

COMPLETE, INCOMPLETE, EMPTY, ERROR = "complete", "incomplete", "empty", "error"
STATUS_CN = {COMPLETE: "完整", INCOMPLETE: "已生成，但有店铺月份没有完整版本", EMPTY: "没有月度结果，没生成",
             ERROR: "出错，没生成"}
QUARTER_ENDS = ("03", "06", "09", "12")


# ------------------------------------------------------------------ names

def suffix(basis):
    """'_账单基准' / '_订单基准' for MercadoLibre; '' for a platform's only basis."""
    return "_" + R.BASIS_CN[basis] if basis in ("bill", "order") else ""


def workbook_name(market, period, basis):
    """MX_ML_2026-Q3_账单基准.xlsx, CL_FA_2026.xlsx, MX_AZ_2026-07_2026-11.xlsx."""
    pf, cc = markets.split(markets.normalize(market))
    return "%s_%s_%s%s.xlsx" % (cc, pf, period.label, suffix(basis))


def alert_name(market, period, basis):
    """！期间报表异常_MX_ML_2026-Q3_账单基准.txt"""
    pf, cc = markets.split(markets.normalize(market))
    return "！期间报表异常_%s_%s_%s%s.txt" % (cc, pf, period.label, suffix(basis))


def target(period, market, basis, out_dir=None, config_path=None):
    """Where the workbook goes: layout.period_dir(kind, period, market)/<name>."""
    d = layout.period_dir(period.kind, period.label, market, override=out_dir, config_path=config_path)
    return os.path.join(d, workbook_name(market, period, basis))


def markets_of(region):
    """--region of run_monthly ('MX', 'FA-CL', 'MX,FA-CL', 'ALL') -> markets."""
    region = str(region or "").upper().replace(" ", "")
    if region == "ALL":
        return markets.all_markets()
    out = []
    for u in [x for x in region.split(",") if x]:
        m = markets.unit_market(u)
        if m not in out:
            out.append(m)
    return out


# ------------------------------------------------------------------ one workbook

def generate(period, market, basis=None, stores=None, out_dir=None, config_path=None, now=None):
    """Build one market's period workbook. Never raises: the result says what
    happened.

    Result: {market, basis, period, path, status, stores, missing [Cell],
    changed [(store, month, old, new, what)], first, carried, message}.
    status: COMPLETE (written, every store-month final), INCOMPLETE (written,
    some store-month without a final version), EMPTY (no results at all for the
    period: nothing written), ERROR (nothing written; message says why)."""
    market = markets.normalize(market)
    pf, _ = markets.split(market)
    basis = MD.basis_of(pf, basis)
    res = {"market": market, "basis": basis, "period": period, "path": None, "status": ERROR,
           "stores": [], "missing": [], "changed": [], "first": True, "carried": 0, "message": "",
           "data": None}
    try:
        D = MD.gather(market, period, basis, stores, config_path, built=now)
    except Exception as e:                               # noqa: BLE001 - reported, never raised
        res["message"] = "读月度结果时出错：%s: %s" % (e.__class__.__name__, e)
        res["trace"] = traceback.format_exc()
        return res
    res.update(data=D, stores=list(D.stores), missing=D.missing())
    if not any(c.final or c.newest for c in D.cells.values()):
        res["status"] = EMPTY
        res["message"] = ("%s %s：期间内没有任何月度结果%s，没生成。"
                          % (markets.label(market), period.label,
                             "（所列店铺）" if stores else ""))
        return res
    path = target(period, market, basis, out_dir, config_path)
    res["path"] = path
    tmp = path + ".part"
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        previous = WB.versions_of(path)
        changes = WB.build(D, tmp, previous)
        os.replace(tmp, path)
    except PermissionError as e:
        _drop(tmp)
        res["path"] = None
        res["message"] = ("写不进去：%s。这个文件多半正被 Excel 打开着 —— 关掉以后再生成一次。（%s）"
                          % (path, e))
        return res
    except Exception as e:                               # noqa: BLE001
        _drop(tmp)
        res["path"] = None
        res["message"] = "生成时出错：%s: %s" % (e.__class__.__name__, e)
        res["trace"] = traceback.format_exc()
        return res
    res["changed"] = changes["rows"]
    res["first"] = changes["first"]
    res["carried"] = len(D.carried_checks())
    res["status"] = INCOMPLETE if res["missing"] else COMPLETE
    res["message"] = "已写出：%s" % path
    return res


def _drop(p):
    try:
        if os.path.exists(p):
            os.remove(p)
    except OSError:
        pass


def summary_lines(res):
    """A result as Chinese lines (console, run log, alert file)."""
    p = res["period"]
    head = "%s · %s%s：%s" % (markets.label(res["market"]), p.label,
                             "（%s）" % MD.basis_cn(res["basis"]) if MD.basis_cn(res["basis"]) else "",
                             STATUS_CN[res["status"]])
    out = [head]
    if res["message"]:
        out.append("  " + res["message"])
    if res["status"] in (COMPLETE, INCOMPLETE):
        D = res["data"]
        out.append("  店铺 %d 家，%d 个月；用了 %d 个店铺月份的完整版本" % (
            len(res["stores"]), len(p.months), len(D.finals())))
        for c in res["missing"]:
            out.append("  ✗ %s %s 无完整版本：%s" % (c.store, c.month, c.reason()))
        if res["first"]:
            out.append("  第一次生成这份报表")
        elif res["changed"]:
            out.append("  和上一份相比换了版本的店铺月份 %d 个：%s" % (len(res["changed"]), "、".join(
                "%s %s" % (s, m) for s, m, _, _, _ in res["changed"][:10])))
        else:
            out.append("  和上一份相比没有换版本")
        if res["carried"]:
            out.append("  月度版本带过来的提醒 %d 条（见⑩）" % res["carried"])
    return out


# ------------------------------------------------------------------ automatic

def auto_periods(month):
    """The periods a monthly run of `month` (re)generates: the quarter when the
    month ends one; in December also the year."""
    out = []
    if month[5:7] in QUARTER_ENDS:
        out.append(MD.quarter(R.quarter_of(month)))
    if month[5:7] == "12":
        out.append(MD.year(month[:4]))
    return out


def alert_text(res, month, entrance_cn=""):
    lines = ["%s 期间报表异常" % markets.label(res["market"]),
             "时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             "起因：%s 月报跑完后自动生成 %s 的期间报表%s" % (month, res["period"].label,
                                                 "（%s）" % entrance_cn if entrance_cn else ""), ""]
    lines += summary_lines(res)
    lines += ["", "期间报表只用每个店铺每个月的完整版本：没有完整版本的月份没有计入合计。",
              "把那个月的报表重出（「立即运行」→ 只出报表），有了完整版本以后，在窗口的「期间报表」页再生成一次，",
              "或者等下一次月报自动运行，它会重新生成。"]
    return "\n".join(lines)


def write_alert(folder, res, month, entrance_cn=""):
    """Write the alert file next to the month's monthly alerts. Returns its path."""
    os.makedirs(folder, exist_ok=True)
    p = os.path.join(folder, alert_name(res["market"], res["period"], res["basis"]))
    with io.open(p, "w", encoding="utf-8") as f:
        f.write(alert_text(res, month, entrance_cn))
    return p


def clear_alert(folder, res):
    """A complete regeneration removes the alert an earlier one left. Returns
    the path removed, or None."""
    p = os.path.join(folder, alert_name(res["market"], res["period"], res["basis"]))
    if os.path.exists(p):
        try:
            os.remove(p)
            return p
        except OSError:
            return None
    return None
