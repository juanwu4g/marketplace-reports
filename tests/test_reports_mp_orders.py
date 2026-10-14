# -*- coding: utf-8 -*-
"""MercadoPago 逐单验证：一次买几件时，钱记在包裹行，子行是 0 —— 子行不能当成差异。"""
import pandas as pd

import report_build as rb


def _settlement(tmp_path, rows):
    pd.DataFrame(rows, columns=["SOURCE_ID", "TRANSACTION_TYPE", "TRANSACTION_AMOUNT",
                                "TRANSACTION_DATE", "REAL_AMOUNT"]).to_csv(
        tmp_path / "settlement_v2-1-manual.csv", sep=";", index=False)


def test_package_child_rows_are_not_counted_as_differences(tmp_path):
    # 2026-10-14 BOCINA_TA02 真实形态：包裹行 336.00，两件子行 0，平台按件收款 206 + 130
    _settlement(tmp_path, [["P1", "SETTLEMENT", 206, "2026-09-05T14:25:00Z", 206.0],
                           ["P2", "SETTLEMENT", 130, "2026-09-05T14:25:00Z", 130.0],
                           ["P3", "SETTLEMENT", 99, "2026-09-06T10:00:00Z", 99.0]])
    A = pd.DataFrame({
        "oid": ["2000014884065455", "2000018301890626", "2000018301887012", "2000018300000001"],
        "pid": ["", "", "", ""],
        rb.V_TOTAL: [336.0, 0.0, 0.0, 120.0],
        rb.V_REVENUE: [400.0, 0.0, 0.0, 150.0],
        rb.V_PACK_FLAG: [" ", "Sí", "Sí", " "]})
    B = pd.DataFrame({  # 账单把收款号挂在子单号上，包裹号放在 Número de paquete
        "k_sale": ["2000018301890626", "2000018301887012", "2000018300000001"],
        "k_pack": ["2000014884065455", "2000014884065455", ""],
        "Pago": ["P1", "P2", "P3"]})
    mo = rb.mp_by_order(A, B, str(tmp_path))
    assert mo["pack_children"] == 2
    assert mo["linked"] == 2 and mo["agree"] == 1, "包裹行 336 = 206 + 130，对得上"
    assert mo["differ"] == [("2000018300000001", 120.0, 99.0)], "真正的差异照样报"
