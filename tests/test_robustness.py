# -*- coding: utf-8 -*-
"""日期用平台那边的"今天"、MercadoPago 区间盖住目标月、旧欠账在补收前撤掉。"""
import datetime
import json
import os

import download_diag
import mercadopago
import pending_store
import run_batch
import run_downloads
import site_time


def _utc(*a):
    return datetime.datetime(*a, tzinfo=datetime.timezone.utc)


def test_site_today_is_mexico_city_not_this_machine():
    # 中国时间 2026-10-14 02:00 = UTC 10-13 18:00 = 墨西哥城 10-13 12:00
    assert site_time.site_today(_utc(2026, 10, 13, 18, 0)) == datetime.date(2026, 10, 13)
    # 墨西哥城午夜过后才算第二天：UTC 06:00 = 墨西哥城 00:00
    assert site_time.site_today(_utc(2026, 10, 14, 5, 59)) == datetime.date(2026, 10, 13)
    assert site_time.site_today(_utc(2026, 10, 14, 6, 0)) == datetime.date(2026, 10, 14)


def test_mercadopago_range_covers_the_target_month():
    today = datetime.date(2026, 10, 15)
    # 15 号做上个月：默认 60 天已经盖住 9 月，不动
    assert mercadopago.range_start(today, 60, datetime.date(2026, 9, 1)) == datetime.date(2026, 8, 16)
    # 补做 7 月：起点放到 7 月 1 号
    assert mercadopago.range_start(today, 60, datetime.date(2026, 7, 1)) == datetime.date(2026, 7, 1)
    assert mercadopago.range_start(today, 30) == datetime.date(2026, 9, 15)


def test_cover_from_follows_the_month_argument():
    class A(object):
        month = "2026-07"
    assert run_downloads.mp_cover_from(A()) == datetime.date(2026, 7, 1)
    A.month = None
    assert run_downloads.mp_cover_from(A()) is None


def test_stale_owed_entries_are_dropped_before_collecting(tmp_path, monkeypatch, capsys):
    store = tmp_path / "BOCINA_TA02"
    old, new = store / "20261013_161937_m202609", store / "20261013_165113_m202609"
    for d in (old, new):
        d.mkdir(parents=True)
    expected = ["mercadopago/after_collection", "mercadopago/settlement"]
    (old / "run_meta.json").write_text(json.dumps({"expected": expected, "missing": []}))
    (new / "run_meta.json").write_text(json.dumps(
        {"expected": expected,
         "missing": [{"form": "mercadopago/settlement", "cause": "PLATFORM_DELAYED"}]}))
    moot = {"kind": "mercadopago", "report": "after_collection", "store": "BOCINA_TA02",
            "out_dir": str(old)}
    still = {"kind": "mercadopago", "report": "settlement", "store": "BOCINA_TA02",
             "out_dir": str(old)}
    own = {"kind": "mercadopago", "report": "settlement", "store": "BOCINA_TA02",
           "out_dir": str(new)}
    removed = []
    monkeypatch.setattr(pending_store, "remove", lambda es: removed.extend(es))
    left = run_batch.drop_superseded([moot, still, own])
    assert removed == [moot], "最新一趟下到了 after_collection：旧的那条作废"
    assert left == [still, own], "最新一趟也没下到的、和最新一趟自己的，都留着"
    assert "撤销旧欠账 mercadopago/after_collection" in capsys.readouterr().out


def test_stock_reports_that_start_no_download_get_one_retry():
    import meli_forms
    assert "general" in meli_forms.STOCK_RETRY_ON_NO_START
    assert "consolidado" not in meli_forms.STOCK_RETRY_ON_NO_START, "要填弹窗的不在此列"


def test_mercado_pago_billing_report_does_not_count_as_missing_report_data():
    assert not download_diag.affects_report({"form": "billing/2026-09/BILL_MP_V3"})
    assert download_diag.affects_report({"form": "billing/2026-09"})
