# -*- coding: utf-8 -*-
"""MercadoLibre report — helpers shared by several sheets (store columns, per-store
sentences, the weighted tax rate, the storage-report comparison text, the "no
Full in this country" sentence)."""

import pandas as pd

from core import report_kit

from platforms.mercadolibre.report import site
from platforms.mercadolibre.report.catalog import F_DETALLE, fmt
from platforms.mercadolibre.report.loading import STORAGE_BILL_LAG_DAYS


# ════════════════════════════════════════════════════════════════════════
# 出表层 —— 通用骨架
# ════════════════════════════════════════════════════════════════════════

Sheet = report_kit.Sheet


def store_cols(stores):
    return [S["store"] for S in stores]


def fee_rows(S, es, reversal=False):
    """(笔数, 金额) —— 当期账单里某个科目的原始行或其冲销行。"""
    FA = S["billing"]
    f = FA[FA[F_DETALLE].astype(str) == es] if not reversal else \
        FA[FA["is_rev"] & (FA["parent_detalle"].astype(str) == es)]
    return int(len(f)), float(f["amount"].sum())


def per_store(stores, fn, sep="；"):
    """'EWTTO_SM 130,027.31；UNIT_PW01 2,921.26'"""
    return sep.join("%s %s" % (S["store"], fn(S)) for S in stores)


def main_store(stores):
    return max(stores, key=lambda S: S["m"]["ing"])


def agg_tax_rate(stores):
    """全部店铺合计的代扣税率（按 GMV 加权）。

    凡是句子里引用的金额是"全部店铺合计"，税率就必须同口径。早先这类句子用的是
    main_store 的税率 —— 单店报表看不出问题，12 家店的汇总报表却会把 GMV 最大
    那家的税率说成全体的。
    """
    tax = sum(S["m"]["tax_ret"] for S in stores)
    gmv = sum(S["m"]["ing"] for S in stores)
    return (tax / gmv) if gmv else 0.0


def store_count_cn(stores):
    """店铺数量的中文表述，供正文使用。"""
    n = len(stores)
    return "本店" if n == 1 else "%d 家店铺合计" % n


def storage_compare_text(m):
    """智利：仓储费报表（按占用日）与账单（按计费日）同一段日子的对比说明。"""
    rng = m.get("storage_range")
    if not rng:
        return "本期没有覆盖本月任何一天的仓储费报表，无法交叉核对。"
    lag = STORAGE_BILL_LAG_DAYS
    full = m.get("storage_days", 0) >= m.get("storage_month_days", 0)
    return ("另有独立的仓储费报表（按占用日逐日计费），它和账单一样按店铺的切期日切期，"
            "这次下载的%s覆盖本月 %s–%s 共 %d 天（本月 %d 天）。这几天报表合计 %s；"
            "账单按计费日比占用日晚 %d 天，取账单 %s–%s 的仓储费 %s 来比，差 %s（逐行取整造成）。"
            "本报表统一采用账单口径，因为那才是真正扣钱的金额。"
            % ("两期合起来完整" if full else "只", rng[0].strftime("%m/%d"),
               rng[1].strftime("%m/%d"), m.get("storage_days", 0),
               m.get("storage_month_days", 0), fmt(m["storage_accrual"]), lag,
               (rng[0] + pd.Timedelta(days=lag)).strftime("%m/%d"),
               (rng[1] + pd.Timedelta(days=lag)).strftime("%m/%d"),
               fmt(m["storage_bill_cmp"]), fmt(m["storage_bill_cmp"] - m["storage_accrual"])))


def no_full_text():
    """没有 Full 仓的国家（秘鲁），⑥ ⑨ 页和 ① 页用的一句话。"""
    return ("%s没有 Full 仓（平台在本国不提供 Full 履约，库存报表页 404）：没有仓储费、Full 揽收/"
            "退仓/罚金等费用，也没有 Full 退货质检和库存报表。本页不适用，不是数据缺失。"
            % site.COUNTRY_CN)
