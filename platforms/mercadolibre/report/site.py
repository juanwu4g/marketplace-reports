# -*- coding: utf-8 -*-
"""MercadoLibre report — the country profile (Mexico / Chile / Peru).

One workbook belongs to one country. set_site(code) writes that country's values
(currency, money format, tolerances, VAT rate, bill cycle day, withholding or not,
the Ventas sheet name and the currency-suffixed Ventas column names …) into this
module's globals. Every other module of the report reads them as `site.NAME`, so
the value is always the one of the current run; nothing copies them.

Also here: localize() / localize_workbook(), which turn the Chile wording of the
"no withholding" texts into Peru's (SUNAT, IGV, 索尔 …) for Peru only."""

import re


# ════════════════════════════════════════════════════════════════════════
# 站点（国家）—— 墨西哥 / 智利 两份报表是"姐妹"：同样 10 页、同样的勾稽，
# 但币种、Ventas 页名、列名后缀、金额精度、容差和税务口径不同。
# ════════════════════════════════════════════════════════════════════════
# 一份工作簿只属于一个国家（币种都不一样，混在一起没有意义），而本文件每次
# 只出一份工作簿，所以和 meli_forms.set_site() 一样，用 set_site() 在开始时
# 把下面这些模块级常量换成该国的值，其余模块按 site.<名字> 读它们（见 STATE）。
#
# 墨西哥的值就是本文件原来写死的那些常量，一个字不改 —— MX 报表必须逐格不变
# （tests/test_cl_mx_unchanged.py 用一份真实 MX 下载目录逐格比对）。
#
# 本模块不 import meli_site，这里只放**出报表**要用的
# 国别事实；主机名、locale 这类下载用的事实在 meli_site.py。
SITE = "mx"
SITE_PROFILES = {
    "mx": {
        "VENTAS_SHEET": "Ventas MX",
        "VENTAS_GLOB": "*Ventas_MX*.xlsx",
        "CUR": "MXN", "CUR_CN": "墨西哥比索", "COUNTRY_CN": "墨西哥",
        "MNY": '#,##0.00;(#,##0.00);-',
        "MONEY_DEC": 2,          # fmt() 默认小数位
        "TOL_CENT": 0.01,        # 桥B / 账单自平 / Full 核对：两边"相等"的容差
        "CN_TOL": 0.05,          # 逐单比对（贷记明细、MercadoPago）的容差
        "BRIDGE_A_TOL": 0.05,    # 桥A 残差容差（--tol 的默认值）
        "COUPON_TOL": 0.02,      # 倒推券额大于它才算一笔优惠券
        "ZERO_TOL": 0.005,       # "这一行是不是 0"（决定标色、是否列出）
        "IVA_RATE": 0.16,
        # 账单周期从每月几号开始。墨西哥是自然月（1 号）。
        "BILL_CYCLE_DAY": 1,
        # 平台是否代扣税金。墨西哥是（IVA/ISR retenciones，约 9.05% 含税价），
        # 有一整套"代扣代缴税金"的行、校验和敏感性分析。
        "WITHHOLDING": True,
        # ⑦ ⑧ 页"Ventas 佣金列 − 账单纯佣金"那一列的叫法
        "TAXCOL_CN": "平台代扣代缴税金", "TAX_SHORT": "代扣税",
    },
    "cl": {
        "VENTAS_SHEET": "Ventas CL",
        "VENTAS_GLOB": "*Ventas_CL*.xlsx",
        "CUR": "CLP", "CUR_CN": "智利比索", "COUNTRY_CN": "智利",
        # 比索没有分：金额一律不显示小数。单元格里的值仍保留原始精度（Ventas
        # 的配送费会出现 .5 比索），只是显示取整 —— 不在写入时舍入，免得合计
        # 和原始文件差出几个比索。
        "MNY": '#,##0;(#,##0);-',
        "MONEY_DEC": 0,
        # 容差按"1 比索"给。Ventas 的 Total 带 .4/.5 这样的零头，而
        # MercadoPago 流水是整数，0.05 的容差会让逐单比对凭空失败。
        "TOL_CENT": 1.0,
        "CN_TOL": 1.0,
        "BRIDGE_A_TOL": 1.0,
        "COUPON_TOL": 1.0,
        "ZERO_TOL": 0.5,
        "IVA_RATE": 0.19,
        # 智利的账单不是自然月，而且**每家店的切期日不同**：DAB_CL 是 8 号到
        # 次月 7 号（Sept 文件计费日 08-08..09-07，Oct 文件 09-08..10-07），
        # Tecnohub_CL 是 24 号到次月 23 号（08-24..09-23，09-24..）。仓储费报表
        # 跟着各店的切期日走。所以切期日按店从数据里推（infer_cycle_day），
        # 这里的 8 只是一份数据都推不出来时的兜底。见 bill_cycles()。
        "BILL_CYCLE_DAY": 8,
        # 智利平台**不代扣**。DAB_CL 2026-09 实测：MercadoPago 结算流水的
        # TAXES_AMOUNT 全部为 0；Ventas 的 Cargo por venta e impuestos 与账单
        # 纯佣金 Cargo por venta 逐单相等（2,406 单里 2,389 单分毫不差，比值
        # 1.000，不是 1.19），差额是被并进这一列的退货处理费。所以智利报表
        # 没有"代扣代缴税金"这一行，见 sheet_pl / 桥 E 的智利版说明。
        "WITHHOLDING": False,
        "TAXCOL_CN": "佣金列中的退货处理费等", "TAX_SHORT": "退货处理费等",
        "VAT_NAME": "IVA", "TAX_AUTH": "SII",
    },
    "pe": {
        "VENTAS_SHEET": "Ventas PE",
        "VENTAS_GLOB": "*Ventas_PE*.xlsx",
        "CUR": "PEN", "CUR_CN": "秘鲁索尔", "COUNTRY_CN": "秘鲁",
        # 索尔有分（céntimos）：金额格式、精度和容差都和墨西哥一样。IDG_PE
        # 2026-09 实测 Ventas、账单、MercadoPago 都是两位小数（118.00、-14.16）。
        "MNY": '#,##0.00;(#,##0.00);-',
        "MONEY_DEC": 2,
        "TOL_CENT": 0.01,
        "CN_TOL": 0.05,
        "BRIDGE_A_TOL": 0.05,
        "COUPON_TOL": 0.02,
        "ZERO_TOL": 0.005,
        # 秘鲁的增值税叫 IGV（Impuesto General a las Ventas），18%，向 SUNAT 申报。
        "IVA_RATE": 0.18,
        # 账单按自然月：账单页写"Cierre 30/set/2026"；IDG_PE 实测 Sept 账单计费日
        # 09-01..09-30、Oct 账单 10-01 起。和墨西哥一样一个会计月一张账单。
        "BILL_CYCLE_DAY": 1,
        # 秘鲁平台**不代扣、也不 percibe**。IDG_PE 2026-09 实测：MercadoPago 流水
        # （settlement / reserve-release）TAXES_AMOUNT 全部为 0，没有 retención /
        # percepción 的交易类型或列；Ventas 的 Cargo por venta e impuestos 与账单
        # 纯佣金逐单相等（= 售价 × 类目费率，118 × 12% = 14.16，不另加 IGV）。
        # 所以和智利一样走"无代扣"那一套（桥 E 换成"无代扣"核对）。
        "WITHHOLDING": False,
        "TAXCOL_CN": "佣金列中的退货处理费等", "TAX_SHORT": "退货处理费等",
        "VAT_NAME": "IGV", "TAX_AUTH": "SUNAT",
        # 秘鲁没有 Full 仓：库存、仓储费、Full 费用、退货质检这些报表平台不提供
        # （库存页 404）。相关页写"本国没有 Full"，不当成缺文件。
        "HAS_FULL": False,
        # 秘鲁 Ventas 的 "Descuentos y bonificaciones" 列**有值**（墨西哥/智利导出
        # 为空列）：IDG_PE 2026-09 是 69 单 Mercado Envíos Flex 的配送补贴，合计
        # 427.37，与 MercadoPago 流水里 9 月的 69 笔 CASHBACK 分毫不差。这是平台
        # 付给卖家的钱，不是卖家承担的优惠券 —— 见 load_ventas。
        "BONIF_COL": True,
    },
}
# 后来才有的键：没写的站点取这里的值（墨西哥的值 = 原来的行为）。
SITE_DEFAULTS = {
    "VAT_NAME": "IVA", "TAX_AUTH": "SAT",
    "HAS_FULL": True,
    "BONIF_COL": False,
}


# set_site() 写进本模块的全部名字。别的模块一律按 site.<名字> 读，不要
# from site import 它们 —— 那样拿到的是 import 那一刻的值，换了站点也不会变。
# Ventas 的金额列名带币种后缀（V_REVENUE = "Ingresos por productos (MXN)" …），
# 所以也在这里；不带币种的列名（V_ORDER_ID、V_SKU …）在 catalog.py。
STATE = tuple(sorted(
    {"SITE", "V_REVENUE", "V_COMMISSION", "V_SHIP_INCOME", "V_SHIP_COST", "V_REFUND", "V_TOTAL",
     "V_UNIT_PRICE"} | set(SITE_DEFAULTS) | set().union(*SITE_PROFILES.values())))


def _ventas_col(stem):
    return "%s (%s)" % (stem, CUR)


def set_site(code):
    """把模块级常量换成某个国家的值。返回规范化后的站点代码。

    出报表前必须调一次（main() 会调）。未知代码直接报错 —— 拿墨西哥的列名
    去读智利的文件只会得到一张全是 0 的报表，看起来还像是对的。
    """
    code = str(code or "mx").lower()
    if code not in SITE_PROFILES:
        raise ValueError("未知站点 %r，可用：%s" % (code, ", ".join(sorted(SITE_PROFILES))))
    g = globals()
    g["SITE"] = code
    g.update(SITE_DEFAULTS)
    g.update(SITE_PROFILES[code])
    # Ventas 的金额列名带币种后缀：Ingresos por productos (MXN) / (CLP)
    g["V_REVENUE"] = _ventas_col("Ingresos por productos")
    g["V_COMMISSION"] = _ventas_col("Cargo por venta e impuestos")
    g["V_SHIP_INCOME"] = _ventas_col("Ingresos por envío")
    g["V_SHIP_COST"] = _ventas_col("Costos de envío")
    g["V_REFUND"] = _ventas_col("Anulaciones y reembolsos")
    g["V_TOTAL"] = _ventas_col("Total")
    g["V_UNIT_PRICE"] = _ventas_col("Precio unitario de venta de la publicación")
    return code


def is_cl():
    return SITE == "cl"


def is_pe():
    return SITE == "pe"


# ── 秘鲁沿用智利"不代扣"那一套说明文字 ──────────────────────────────────
# 无代扣分支（桥 E/F、② 备查项、① 注释、⑩ 待办、no_withholding 校验……）的
# 文字是为智利写的，里面直接写着"智利 / IVA / SII / DTE / 19/119"。秘鲁的税务
# 结构经实测和智利相同（见 SITE_PROFILES["pe"]），只是名字不同，所以不复制一份
# 秘鲁版文字，而是在写出工作簿和校验结果时把这几个专名换掉（localize）。
# 只对秘鲁生效：墨西哥、智利的输出一个字不变（golden 测试逐格比对）。
# 新写智利专属文字时注意：这里换不掉的说法（例如"截账日"那一套）要自己按
# 站点分支，别指望这张表。
LOCALIZE_FROM_CL = {
    "pe": [
        ("智利", "秘鲁"),
        ("19/119", "18/118"),
        ("而不是 1.19", "而不是 1.18"),
        ("比索", "索尔"),
        ("SII", "SUNAT"),
        ("（DTE）", "（CPE，comprobantes de pago electrónicos）"),
    ],
}
_IVA_RE = re.compile(r"(?<![A-Za-z])IVA(?![A-Za-z])")


def localize(text):
    """把为智利写的说明换成本站点的专名（秘鲁：智利→秘鲁、IVA→IGV、SII→SUNAT）。
    墨西哥/智利原样返回；非字符串原样返回。"""
    pairs = LOCALIZE_FROM_CL.get(SITE)
    if not pairs or not isinstance(text, str):
        return text
    for a, b in pairs:
        text = text.replace(a, b)
    return _IVA_RE.sub(VAT_NAME, text)


def localize_workbook(wb):
    """写盘前对整本工作簿做 localize（只改文字单元格，公式原样）。"""
    if not LOCALIZE_FROM_CL.get(SITE):
        return
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith("="):
                    nv = localize(v)
                    if nv != v:
                        c.value = nv


set_site(SITE)      # 模块默认：墨西哥。main() 会按店铺的 run_meta / --site 重设
