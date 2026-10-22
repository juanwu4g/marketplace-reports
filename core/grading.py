# -*- coding: utf-8 -*-
"""The grade of one monthly report version (docs/RESULTS_AND_PERIODS.md §3).

    grade(platform, basis, wrote, checks, missing, data_incomplete, provisional, formula_errors)
        -> {"level": one of core/results.GRADES, "reasons": [short Chinese sentences]}

Levels, first match wins:

  failed         the workbook was not written
  gaps           a form the report reads was not downloaded (cause not NO_DATA /
                 NOT_AVAILABLE, see FORM_OK_CAUSES); or a blocking check failed; or
                 the platform said a source export is incomplete (ML Ventas "No
                 pudimos mostrar algunos datos", order basis); or the recalculated
                 workbook holds formula errors
  provisional    none of the above, but a bill the month uses has not closed
  complete_warn  none of the above, but a non-blocking check failed
  complete       nothing failed

Blocking = an error-level check that failed, except the ones in NON_BLOCKING:
the error-level checks whose failure does not make the money figures wrong, each
with the reason. That list is policy — add to it only with a written reason and
the evidence. Warn-level checks never change the grade (they are in the results
file's "checks" and in the workbook's ⑩).
"""
from core import results

# (platform, basis, check key) -> why its failure leaves the money figures right.
# Evidence: the September 2026 reports of the real folders on the dev box
# (~/report-golden and ~/report-experiments, 2026-10-22).
NON_BLOCKING = {
    ("ML", "bill", "map_cover"):
        "SKU 映射覆盖率只影响⑦按 SKU 的归属；②的收入与费用直接来自账单行，合计不变"
        "（Tecnohub_CL 2026-09 覆盖 24.2%，②与账本残差为 0）",
    ("ML", "order", "cn_orphan"):
        "订单基准的收入与订单级费用按下单日取自销售报表；账单对销售报表里没有的订单计的费，"
        "按定义不属于本基准的本月订单（多为上月订单），金额在⑩列出供核实。"
        "销售报表真被截断时另有 ventas_incomplete / sales_window 判为有缺口"
        "（墨西哥、智利几乎每家店每月都有：TANKE_EE、EWTTO_SM、DAB_CL、Tecnohub_CL 2026-09）",
    ("ML", "order", "cn_unreversed"):
        "已取消订单被计费且未冲销：钱确实被平台扣了，②如实列为费用；这是向平台申诉的待办，"
        "不是报表数字错（DAB_CL 2026-09：1 单 9,750 CLP）",
}

# Causes of a missing form that are not a gap: there was nothing to download
# (core/download_diag.NO_ATTENTION: NO_DATA = the period has no data,
# NOT_AVAILABLE = the platform does not offer it — a billing period not begun,
# a store without Full).
FORM_OK_CAUSES = ("NO_DATA", "NOT_AVAILABLE")

CAUSE_CN = {
    "NO_PERMISSION": "账号没有权限", "UI_NOT_FOUND": "页面控件找不到", "PAGE_BLOCKED": "页面被拦",
    "DOWNLOAD_NOT_LANDED": "文件没有落盘", "STILL_GENERATING": "平台还在生成",
    "PLATFORM_DELAYED": "平台延迟", "PLATFORM_ERROR": "平台出错", "STORE_OPEN_FAILED": "店铺没打开",
    "REQUEST_NOT_REGISTERED": "平台没有登记申请", "FILE_REJECTED": "文件不能用",
    "LANGUAGE_WRONG": "页面语言不对", "INTERNAL_ERROR": "脚本出错", "EXPIRED": "超期未取回",
    "UNKNOWN": "原因不明",
}


def non_blocking_reason(platform, basis, key):
    return NON_BLOCKING.get((platform, basis, key))


def form_gaps(missing):
    """The missing-form records that make a gap: the report reads the form and
    the cause is not one of FORM_OK_CAUSES. A record without "in_report"
    (MercadoLibre's downloader does not write it) is judged by
    core/download_diag.affects_report, as the order report's ⑩ does."""
    from core import download_diag
    out = []
    for rec in missing or ():
        if not rec or not rec.get("form"):
            continue
        if rec.get("cause") in FORM_OK_CAUSES:
            continue
        if download_diag.affects_report(rec):
            out.append(rec)
    return out


def _failed(checks):
    return [c for c in checks or () if not c.get("ok") and not c.get("skipped")
            and c.get("severity", "error") == "error"]


def grade(platform, basis, wrote, checks=(), missing=(), data_incomplete=False, provisional=(),
          formula_errors=False, why_not_written=None):
    """See the module docstring. `checks` are this store's report_kit.check records,
    `missing` the download folder's run_meta.json "missing" records, `provisional`
    the names of bills the month uses that have not closed."""
    if not wrote:
        return {"level": results.FAILED,
                "reasons": ["报表没有生成%s" % ("：%s" % why_not_written if why_not_written else "")]}
    failed = _failed(checks)
    blocking = [c for c in failed if not non_blocking_reason(platform, basis, c.get("key"))]
    soft = [c for c in failed if non_blocking_reason(platform, basis, c.get("key"))]
    gaps = []
    for rec in form_gaps(missing):
        gaps.append("缺下载：%s（%s）" % (rec["form"], CAUSE_CN.get(rec.get("cause"), rec.get("cause") or "原因不明")))
    for c in blocking:
        gaps.append("校验未过：%s" % (c.get("name") or c.get("key")))
    if data_incomplete:
        gaps.append("平台提示销售报表导出不完整（No pudimos mostrar algunos datos），收入可能偏低")
    if formula_errors:
        gaps.append("重算后工作簿里有公式错误")
    if gaps:
        return {"level": results.GAPS, "reasons": gaps}
    if provisional:
        return {"level": results.PROVISIONAL,
                "reasons": ["账单尚未关账：%s（关账后重出）" % "、".join(provisional)]}
    if soft:
        return {"level": results.COMPLETE_WARN,
                "reasons": ["提醒（不影响金额）：%s" % (c.get("name") or c.get("key")) for c in soft]}
    return {"level": results.COMPLETE, "reasons": []}
