"""A platform warning must not discard otherwise readable Ventas rows."""

import pytest
from openpyxl import Workbook

import report_build
import run_reports


@pytest.mark.parametrize("notice,header_row", [
    ("", 6),
    ("No pudimos mostrar algunos datos. Intenta generando el Excel de ventas nuevamente.", 7),
])
def test_load_ventas_keeps_rows_and_warning(tmp_path, notice, header_row):
    path = tmp_path / "Ventas_MX_example.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = report_build.VENTAS_SHEET
    if notice:
        ws["A2"] = notice
    for col, value in enumerate((report_build.V_ORDER_ID, report_build.V_DATE,
                                 report_build.V_STATUS, report_build.V_TOTAL), 1):
        ws.cell(header_row, col, value)
    for col, value in enumerate(("123", "1 de agosto de 2026 10:00 hs.",
                                 "Entregado", 100), 1):
        ws.cell(header_row + 1, col, value)
    wb.save(path)

    selected, ventas = report_build.load_ventas(tmp_path)

    assert selected == str(path)
    assert len(ventas) == 1
    assert ventas.iloc[0]["ym"] == "2026-08"
    assert ventas.attrs["incomplete_notice"] == notice


def test_run_summary_marks_generated_report_as_incomplete(monkeypatch, capsys, tmp_path):
    check = {"key": "ventas_incomplete", "store": "TEST", "name": "Ventas 导出数据不完整",
             "ok": False, "severity": "warn", "detail": "平台提示数据不完整"}
    monkeypatch.setattr(run_reports.store_config, "batch_stores", lambda: ["TEST"])
    monkeypatch.setattr(run_reports, "resolve", lambda names, month: ([("TEST", str(tmp_path))], []))
    monkeypatch.setattr(run_reports.financial_report, "has_libreoffice", lambda: False)
    monkeypatch.setattr(run_reports.financial_report, "build", lambda *args, **kwargs: {
        "ok": True, "wrote": True, "out": str(tmp_path / "report.xlsx"), "seconds": 1,
        "data_incomplete": True, "checks": [check], "returncode": 0})
    monkeypatch.setattr("sys.argv", ["run_reports.py", "--month", "2026-08",
                                    "--no-combined", "--out-dir", str(tmp_path)])

    assert run_reports.main() == 1
    output = capsys.readouterr().out
    assert "数据不全" in output
    assert "Ventas 源数据不完整" in output
