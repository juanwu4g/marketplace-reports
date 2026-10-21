# -*- coding: utf-8 -*-
"""
Mercado Libre —— 月度财务报表生成器（墨西哥 / 智利 / 秘鲁）

    python -m platforms.mercadolibre.report_build --month 2026-08 \
        --store "EWTTO_SM=D:/Projects/mx_sales_data/downloads/EWTTO_SM/20260909_021624" \
        --store "UNIT_PW01(catalog)=D:/.../UNIT_PW01(catalog)/20260909_021624" \
        --out reports/MercadoLibre_月度财务报表_2026年8月.xlsx

本文件是入口：命令行、组装工作簿、重算、导出指标。报表本身分在
platforms/mercadolibre/report/ 下面几个模块里，自下而上：

    site.py          国别口径（币种、容差、税务、Ventas 列名…），set_site() 切换
    catalog.py       原始文件结构、费用科目字典 FEE_CATALOG、状态词典、数字小工具
    loading.py       读一家店的下载目录（Ventas、账单、MercadoPago、库存…）
    derived.py       derive()：一家店 → 数字字典 S（会计口径都在这里）
    checks.py        validate()：校验记录
    sheet_*.py       十页工作簿，一页（或两页）一个模块

以前所有代码都在本文件里，别处的代码和测试按 report_build.<名字> 取用；这些
名字照旧都能从这里取到（见下面"对外名字"一段）。

设计原则
--------
1) 确定性。同样的输入文件，任意次运行产出相同的行位置、相同的公式、相同的数字。
   ② 损益表的 27 行是写死的常量表（sheet_pl.PL_LINES），某个费用科目本月为 0 也
   保留该行，这样跨月对比时行号永远对齐 —— 这是"格式不变、只有趋势变"的前提。

2) 费用科目必须登记。FEE_CATALOG（catalog.py）是唯一的科目字典。出现字典里没有的
   Detalle 时，**绝不静默归入"其他"**：金额照常进入报表（否则勾稽会断），但会被单独
   归到"未登记科目"行、在 ⑨ 页列为高优先级待办、并让 check_fee_catalog 校验失败。
   （旧实现把 comprar más espacio / sobrepasar espacio / stock antiguo /
   incumplimiento 四个科目静默丢进 other，EWTTO 2026-08 合计 6,506.85 因此成了糊涂账。）

3) 冲销按"被冲销的那张单"归类，不靠字符串猜。账单行有 `Número del cargo`，
   冲销行有 `Cargo que bonifica` 指向它，用这个 ID 关联即可把
   "Anulación del cargo por campaña de publicidad de Products Ads"
   （注意平台自己把 Product 写成了 Products）正确地冲减到广告费上。
   ID 关联不上时才退回按名称匹配。

4) 一切"发现"都固化成公式与校验：
   - 代扣代缴税金  = Ventas 的 Cargo por venta e impuestos − 账单纯佣金
   - 卖家优惠券    = 由 Total 列倒推（平台把 Descuentos 列导成了空白）
   - 结算通道拆分  = Descontado de la operación ∈ {Si, No, No aplica}
   校验见 validate()。

依赖：pandas, openpyxl。重算公式缓存值需要 LibreOffice（soffice），可用 --no-recalc 跳过。
"""
from __future__ import print_function

import argparse
import datetime
import io
import json
import os
import sys
import traceback
import warnings

import pandas as pd
from openpyxl import Workbook

# 本文件按路径当脚本跑（financial_report.py），也被测试当模块 import。两种情况
# 下都要能找到仓库里的 core/ 和 platforms/：把仓库根目录放进 sys.path。
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)
from core import report_kit  # noqa: E402
from core import sku_costs as SC  # noqa: E402
from platforms.mercadolibre.report import (  # noqa: E402
    site as _site, catalog, loading, derived, checks, sheet_common, sheet_readme as _m_readme,
    sheet_pl as _m_pl, sheet_bridge as _m_bridge, sheet_fees as _m_fees,
    sheet_returns as _m_returns, sheet_sku as _m_sku, sheet_trend as _m_trend,
    sheet_inventory as _m_inventory, sheet_gaps as _m_gaps)

warnings.filterwarnings("ignore")

# ════════════════════════════════════════════════════════════════════════
# 对外名字
# ════════════════════════════════════════════════════════════════════════
# 拆分以前，所有函数和常量都是本模块的名字（rb.derive、rb.FEE_CATALOG、
# rb._no_withholding_check …），测试和 pl_billing_basis.py 都这样取。这里把各
# 模块的名字原样搬进来，一个不少。
#
# 国别口径（rb.MNY、rb.CUR、rb.V_REVENUE …）不搬：它们随 set_site() 变，搬进来
# 就是一份会过期的副本。模块级 __getattr__ 每次都去 site.py 取当前值。
_MODULES = (_site, catalog, loading, derived, checks, sheet_common, _m_readme, _m_pl,
            _m_bridge, _m_fees, _m_returns, _m_sku, _m_trend, _m_inventory, _m_gaps)
for _m in _MODULES:
    for _k, _v in vars(_m).items():
        if not _k.startswith("__") and _k not in _site.STATE and _k != "site":
            globals().setdefault(_k, _v)
del _m, _k, _v

# 下面用到的名字（上面的循环已经都搬进来了；这里写明，便于查找）
from platforms.mercadolibre.report.site import SITE_PROFILES, localize, localize_workbook, set_site  # noqa: E402,E501
from platforms.mercadolibre.report.catalog import (  # noqa: E402
    MONTH_FULL_ES, VENTAS_HEADER_ROW, V_DATE, fmt, month_bounds)
from platforms.mercadolibre.report.loading import load_run_meta  # noqa: E402
from platforms.mercadolibre.report.derived import derive  # noqa: E402
from platforms.mercadolibre.report.checks import _chk, check_hints, validate  # noqa: E402
from platforms.mercadolibre.report.sheet_readme import sheet_readme  # noqa: E402
from platforms.mercadolibre.report.sheet_pl import sheet_pl  # noqa: E402
from platforms.mercadolibre.report.sheet_bridge import sheet_bridge  # noqa: E402
from platforms.mercadolibre.report.sheet_fees import sheet_fees  # noqa: E402
from platforms.mercadolibre.report.sheet_returns import sheet_returns_money, sheet_triage  # noqa: E402
from platforms.mercadolibre.report.sheet_sku import (  # noqa: E402
    COST_UNITS_BASIS, cost_policy_closing, sheet_sku, sku_items)
from platforms.mercadolibre.report.sheet_trend import sheet_trend  # noqa: E402
from platforms.mercadolibre.report.sheet_inventory import sheet_inventory  # noqa: E402
from platforms.mercadolibre.report.sheet_gaps import sheet_gaps  # noqa: E402


def __getattr__(name):
    """rb.MNY / rb.CUR / rb.V_REVENUE …：当前站点的值（见 site.py）。"""
    if name in _site.STATE:
        return getattr(_site, name)
    raise AttributeError("module %r has no attribute %r" % (__name__, name))


# ════════════════════════════════════════════════════════════════════════
# 退出码
# ════════════════════════════════════════════════════════════════════════
# 调用方（financial_report.build）要分得清"报表写出来了但有校验没过"和"报表
# 根本没写出来"，所以每一种没写出文件的结局都有自己的码，不和 1 混用。以前
# --strict 中止、公式错误、校验未过全是 1，Python 报错也是 1（解释器对未捕获
# 异常的默认退出码），上层只能靠"文件在不在"去猜。
# financial_report.py 里照抄了一份（它刻意不 import 本文件），改这里要同步改那边。
EXIT_OK = 0             # 报表已写出，error 级校验全部通过
EXIT_CHECKS = 1         # 报表已写出，但有 error 级校验未过，或文件里有公式错误
EXIT_NO_STORE = 2       # 没有一家店铺的数据可用，没写文件
EXIT_STRICT_ABORT = 10  # --strict 且有校验未过：按要求不写文件
EXIT_CRASH = 70         # 未处理的异常（traceback 已打印）。70 = sysexits 的 EX_SOFTWARE


# ════════════════════════════════════════════════════════════════════════
# 商品成本表（core/sku_costs，config.json → reports.sku_costs）
# ════════════════════════════════════════════════════════════════════════
# 没配成本表时这里什么也不做，工作簿与以前一字不差。配了的话：每家店的 ⑦ 表五
# 各 SKU 按成本表取单位成本，件数 ＝ 表五“销售件数”（口径见 sheet_sku.COST_UNITS_BASIS），
# ② 商品采购成本 ＝ 有成本的 SKU 之和（仍是黄色可改的录入格），⑩ 多几条成本表校验。
PLATFORM = "ML"         # core/markets.py 的平台代码（成本表的“平台”列）


def apply_costs(stores, month, checks, ctx=None, table=None):
    """有成本表时：每家店得到 S["cost"]（core/sku_costs.Result），校验追加进 checks，
    ctx 记下成本表和件数口径。返回成本表；没配返回 None。国家、币种取当前站点。"""
    table = table or SC.load_configured()
    if table is None:
        return None
    country = _site.SITE.upper()
    for S in stores:
        S["cost"] = SC.apply(table, PLATFORM, country, _site.CUR, month, sku_items(S))
    checks.extend(SC.checks(table, [(S["store"], S["cost"]) for S in stores], PLATFORM, country))
    if ctx is not None:
        ctx["sku_costs"] = table
        ctx["cost_units_basis"] = COST_UNITS_BASIS
        ctx["cost_policy_closing"] = cost_policy_closing()
    return table


def sku_list(store, folder, month):
    """(平台, 国家, 币种, [core/sku_costs.Item]) —— 一家店一个月 ⑦ 表五的 SKU，
    给 tools/sku_cost_template 用。只读下载目录；站点取目录 run_meta.json 的 site
    （和 main() 一样，没有就是 mx），读完把站点换回原来的。没有 SKU 的行（换货单
    第一行等，SKU 显示为 nan）不列：成本表里填不了它。"""
    before = _site.SITE
    try:
        set_site(load_run_meta(folder).get("site") or "mx")
        p0, p1 = month_bounds(month)
        items = [it for it in sku_items(derive(store, folder, p0, p1))
                 if str(it.sku).strip().lower() not in ("", "nan", "none")]
        return PLATFORM, _site.SITE.upper(), _site.CUR, items
    finally:
        set_site(before)


# ════════════════════════════════════════════════════════════════════════
# 组装 / 重算 / CLI
# ════════════════════════════════════════════════════════════════════════

def build_context(stores, month, build_date=None):
    p0, p1 = month_bounds(month)
    cutoffs = [S["ventas_all"]["fecha"].max() for S in stores if len(S["ventas_all"])]
    # 守卫看的必须是**过滤之后**的列表。原来判断的是过滤前的 cutoffs 非空，
    # 可一家店的销售日期要是全为 NaT（导出缺日期列），过滤后就空了，max() 直接炸。
    cutoffs = [c for c in cutoffs if pd.notna(c)]
    cutoff = max(cutoffs) if cutoffs else p1
    # 不完整月份：销售报表窗口两端（最早月与最晚月），且不等于报表月本身时才提示
    months = sorted(set(m for S in stores for m in S["ventas_all"]["ym"].unique() if m != "NaT"))
    partial = set()
    if months:
        partial.add(months[0])
        partial.add(months[-1])
    partial.discard(month)
    contrib = {}
    for S in stores:
        m = S["m"]
        bill = sum(m["bill_fee_net"].values()) + m["unc_bill"]
        # 必须与 ② 合并损益表的同名行逐项一致。改损益表的行结构时这里也要改 ——
        # 它是第三处独立算同一个数的地方（另两处是 ② 的公式和 ③ 桥D），
        # 漏掉任何一处，页面之间就会对不上而没有任何提示。
        contrib[S["store"]] = (m["ing"] + m["env_ing"] - m["coupon"] - m["refund"]
                               + S["cn"]["partial_total"]
                               - m["com_gross"] - m["env_cost"] - m["dev_net"]
                               - m["unc_order"] - bill)
    return {
        "period": month,
        "period_cn": "%d年%d月" % (p0.year, p0.month),
        "period_short": "%d月" % p0.month,
        "period_es": "%s %d" % (MONTH_FULL_ES[p0.month], p0.year),
        "p0": p0, "p1": p1,
        "build_date": build_date or datetime.date.today().strftime("%Y-%m-%d"),
        "cutoff": cutoff.strftime("%Y-%m-%d") if pd.notna(cutoff) else "-",
        "partial_months": sorted(partial),
        "contrib": contrib,
    }


def build_workbook(stores, ctx, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    sheet_readme(wb, stores, ctx)
    sheet_pl(wb, stores, ctx)
    sheet_bridge(wb, stores, ctx)
    sheet_fees(wb, stores, ctx)
    sheet_returns_money(wb, stores, ctx)
    sheet_triage(wb, stores, ctx)
    sheet_sku(wb, stores, ctx)
    sheet_trend(wb, stores, ctx)
    sheet_inventory(wb, stores, ctx)
    sheet_gaps(wb, stores, ctx)
    localize_workbook(wb)
    d = os.path.dirname(os.path.abspath(out_path))
    if d and not os.path.isdir(d):
        os.makedirs(d)
    wb.save(out_path)
    return out_path


# 重算与公式错误扫描：报表家族共用（core/report_kit.py），代码原样搬过去。
recalc = report_kit.recalc
scan_formula_errors = report_kit.scan_formula_errors


def metrics_record(S, ctx):
    """一行月度指标，用于累积到指标库（趋势的真正来源）。"""
    m = S["m"]
    rec = {"period": ctx["period"], "store": S["store"], "built_at": ctx["build_date"]}
    rec["ventas_incomplete"] = bool(S.get("ventas_warning"))
    rec["ventas_warning"] = S.get("ventas_warning", "")
    for k in ("rows", "orders_unique", "units", "ing", "env_ing", "coupon", "coupon_n",
              "refund", "com_imp", "com_gross", "env_cost", "env_gross", "dev_net",
              "tax_ret", "tax_rate_gross", "total", "resid", "bill_total", "payable",
              "pay_auto", "pay_nc", "other_total", "ads_net", "full_bill_total",
              "subscription_net", "storage_bill", "storage_accrual", "unc_order",
              "unc_bill", "abnormal_orders", "n_back_on_sale", "n_stock_lost",
              "n_seller_keeps", "amt_stock_lost", "iva_in_gmv"):
        v = m.get(k)
        rec[k] = round(v, 6) if isinstance(v, float) else v
    rec["contribution"] = round(ctx["contrib"][S["store"]], 2)
    if S.get("cost") is not None:          # 配了商品成本表才有
        rec["cogs"] = round(S["cost"].cogs, 2)
        rec["profit"] = round(ctx["contrib"][S["store"]] - S["cost"].cogs, 2)
    for es, v in m["bill_fee_net"].items():
        rec["fee:" + es] = round(v, 2)
    return rec


def parse_store_arg(v):
    if "=" not in v:
        raise argparse.ArgumentTypeError("--store 需要 NAME=PATH 形式，收到 %r" % v)
    name, path = v.split("=", 1)
    name, path = name.strip(), os.path.expanduser(path.strip())
    if not os.path.isdir(path):
        raise argparse.ArgumentTypeError("目录不存在：%s" % path)
    return (name, path)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Mercado Libre 月度财务报表生成器（墨西哥 / 智利 / 秘鲁）",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", required=True, help="会计期间，如 2026-08")
    ap.add_argument("--store", action="append", required=True, type=parse_store_arg,
                    metavar="NAME=PATH", help="店铺名=下载目录，可重复")
    ap.add_argument("--out", default=None, help="输出 xlsx 路径")
    ap.add_argument("--metrics-out", default=None, help="同时导出月度指标 JSON（用于累积指标库）")
    ap.add_argument("--strict", action="store_true",
                    help="任一 error 级校验未通过、或平台声明 Ventas 导出不完整，"
                         "就不写出文件，退出码 %d" % EXIT_STRICT_ABORT)
    ap.add_argument("--no-recalc", action="store_true", help="跳过 LibreOffice 重算")
    ap.add_argument("--tol", type=float, default=None,
                    help="桥A 残差容差，默认按站点：墨西哥/秘鲁 0.05，智利 1 比索")
    ap.add_argument("--build-date", default=None, help="覆盖编制日期（用于可复现测试）")
    ap.add_argument("--site", type=str.lower, choices=sorted(SITE_PROFILES), default=None,
                    help="国家站点 mx / cl / pe。不给时取第一家店下载目录 run_meta.json 里的 site"
                         "（老目录没有这个字段，按 mx）")
    a = ap.parse_args(argv)

    # ---- 站点：一份工作簿只属于一个国家 ----
    # 每家店的下载目录在 run_meta.json 里记着 "site"。与本次站点不一致的店铺
    # 不能混进来（币种、列名、账单周期、税务口径都不一样），按"店铺数据可用"
    # 未通过剔除，和坏数据的店铺同一个出口，汇总里看得见。
    site_of = {name: str(load_run_meta(path).get("site") or "mx").lower()
               for name, path in a.store}
    site = set_site(a.site or site_of[a.store[0][0]])
    if site != "mx":
        print("[站点] %s（%s，%s）" % (site, _site.COUNTRY_CN, _site.CUR))
    if a.tol is None:
        a.tol = _site.BRIDGE_A_TOL

    p0, p1 = month_bounds(a.month)
    stores, unusable = [], []
    for name, path in a.store:
        if site_of[name] != site:
            unusable.append((name, "下载目录记的站点是 %s（run_meta.json），本报表是 %s 的。"
                                   "两国的币种、账单周期和税务口径不同，不能放进同一份报表；"
                                   "请按国家分开出（run_reports.py --region）。"
                             % (site_of[name].upper(), site.upper())))
            continue
        print("[读取] %s ← %s" % (name, path))
        S = derive(name, path, p0, p1)
        V = S["ventas_all"]
        # Ventas 导出偶尔整列缺销售日期（平台换了导出模板，或者文件本身是坏的）。
        # 没有日期就没有会计期间，这家店的所有口径都无从谈起。以前这种数据要一路
        # 走到 build_context 才炸 —— 而且是在**汇总**阶段炸，一家店的坏文件把另外
        # 11 家的汇总表一起带走。就地剔除，并在 ⑩ 页记一条未通过的校验。
        if len(V) and not V["fecha"].notna().any():
            # 把"为什么没有日期"直接说出来。两种成因的处理办法完全不同，而光说
            # "没有日期"两种都像，只能去人工开文件看。TOOL_TA03（9/15）和
            # EWTTO_SM（9/24）先后中过同一个症状，两次都得重新翻文件才知道。
            cols = [str(c) for c in V.columns if not str(c).startswith("Unnamed")]
            if V_DATE in V.columns:
                why = ("列「%s」在，但没有一个值能解析成日期 —— 多半是平台改了"
                       "日期格式。" % V_DATE)
            else:
                why = ("导出里根本没有「%s」这一列 —— 多半是导出时账号界面语言"
                       "不是西班牙语（列名会跟着变），或者表头不在第 %d 行。"
                       % (V_DATE, VENTAS_HEADER_ROW + 1))
            unusable.append((name, "Ventas 报表没有一行带销售日期。%s"
                                   "文件里的列名（前 8 个）：%s。"
                                   "本店已从本报表剔除，请重新下载这家店的 "
                                   "Ventas 再出一次。"
                                   % (why, "、".join(cols[:8]) or "（一个都没有）")))
            continue
        stores.append(S)

    for name, why in unusable:
        print("[剔除] %s —— %s" % (name, why))
    if not stores:
        print("\n[中止] 没有一家店铺的数据可用。")
        return EXIT_NO_STORE

    ctx = build_context(stores, a.month, a.build_date)
    checks = validate(stores, tol=a.tol)
    # 排在最前面：被剔除的店铺不会出现在下面任何一条校验里，不显式记一笔的话，
    # 汇总表只会写“11 家店铺”，没人会注意到少了一家。
    checks = [_chk("store_usable", name, "店铺数据可用", False, why)
              for name, why in unusable] + checks
    ctx["checks"] = checks
    apply_costs(stores, a.month, checks, ctx)        # 商品成本表；没配什么也不做

    # 每条校验只在这里打印一次。未通过的那几条，排查提示就跟在它下面 —— 以前
    # 另有一个 [诊断] 段把未通过的校验原样再打一遍，同一句话在一份报表的输出
    # 里出现两次，汇总里还有第三次。
    by_store = {S["store"]: S for S in stores}
    print("\n[校验]")
    for c in checks:
        mark = "跳过" if c["skipped"] else ("  ✓" if c["ok"] else "  ✗")
        print("  %s  %-12s %-22s %s" % (mark, c["store"][:12], c["name"], c["detail"]))
        if not c["ok"] and not c["skipped"]:
            for line in check_hints(c, by_store.get(c["store"])):
                print("        %s" % localize(line))
    failed = [c for c in checks if not c["ok"] and c["severity"] == "error"]
    print("  —— error 级未通过 %d 项；warn 级未通过 %d 项" % (
        len(failed), sum(not c["ok"] and c["severity"] == "warn" for c in checks)))
    # 稳定的机器可读记录供调用方汇总；不能只靠退出码 1 判断是哪一项失败。
    print("[REPORT_CHECKS_JSON] " + json.dumps(checks, ensure_ascii=False))

    print("\n[结果]")
    for S in stores:
        m = S["m"]
        print("  %-16s GMV %14s  到账 %14s  经营贡献毛利 %14s"
              % (S["store"], fmt(m["ing"]), fmt(m["total"]), fmt(ctx["contrib"][S["store"]])))
    print("  %-16s %18s %19s %18s"
          % ("合计", fmt(sum(S["m"]["ing"] for S in stores)),
             fmt(sum(S["m"]["total"] for S in stores)), fmt(sum(ctx["contrib"].values()))))
    if ctx.get("sku_costs") is not None:
        for S in stores:
            print("  %-16s 商品采购成本 %s（%s）  税前净利润 %s"
                  % (S["store"], fmt(S["cost"].cogs), S["cost"].coverage_text(),
                     fmt(ctx["contrib"][S["store"]] - S["cost"].cogs)))

    # --strict 的承诺是"数据有问题就不出文件"。ventas_incomplete 是 warn 级的
    # （不 strict 时照常出报表、只做标记，这一点不动），但它的意思是**平台自己
    # 承认这份导出缺数据** —— strict 下放它过去，就等于出了一份明知偏低的报表。
    # 所以只在 strict 这一处把它和 error 级校验同等对待，别处的级别不变。
    blocking = failed + [c for c in checks
                         if c["key"] == "ventas_incomplete" and not c["ok"]]
    if blocking and a.strict:
        print("\n[中止] %d 项校验未通过，--strict 模式下不生成报表：%s"
              % (len(blocking), "、".join(sorted(set(c["key"] for c in blocking)))))
        return EXIT_STRICT_ABORT

    out = a.out or os.path.join("reports", "MercadoLibre_月度财务报表_%s.xlsx" % ctx["period_cn"])
    build_workbook(stores, ctx, out)
    print("\n[输出] %s" % out)

    if not a.no_recalc:
        ok, msg = recalc(out)
        print("[重算] %s" % msg)
        if ok:
            nf, ne, nn, bad = scan_formula_errors(out)
            print("[公式] 共 %d 条，错误 %d 条，未求值 %d 条" % (nf, ne, nn))
            for b in bad:
                print("       %s" % b)
            if ne:
                print("[警告] 文件里存在公式错误，请勿直接发出。")
                return EXIT_CHECKS

    if a.metrics_out:
        recs = [metrics_record(S, ctx) for S in stores]
        d = os.path.dirname(os.path.abspath(a.metrics_out))
        if d and not os.path.isdir(d):
            os.makedirs(d)
        with io.open(a.metrics_out, "w", encoding="utf-8") as f:
            f.write(json.dumps(recs, ensure_ascii=False, indent=1))
        print("[指标] %s" % a.metrics_out)

    return EXIT_CHECKS if failed else EXIT_OK


def run(argv=None):
    """命令行入口：main() 之外再包一层，把未处理的异常变成一个专用退出码。

    不包的话，任何一处 KeyError / PermissionError 都以解释器默认的退出码 1 结束，
    和"有校验未过"是同一个数，调用方只能把"报表炸了"当成"报表出了、有点小问
    题"。traceback 照常打印 —— 这里换的是退出码，不是把错误藏起来。
    argparse 的参数错误（SystemExit 2）和 Ctrl+C 原样放行。
    """
    try:
        return main(argv)
    except Exception:
        # 打到 stdout：调用方把两路合并读，但直接在终端跑并把 stdout 重定向到
        # 文件时，traceback 也应该跟着其余输出一起留下来。
        print("\n[崩溃] 报表生成中途出错，文件未写出或不完整：")
        traceback.print_exc(file=sys.stdout)
        sys.stdout.flush()
        return EXIT_CRASH


if __name__ == "__main__":
    sys.exit(run())
