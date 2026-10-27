# -*- coding: utf-8 -*-
"""Falabella download layer: the parts that need no browser."""
import datetime
import json

import pytest

from core import download_diag
from core import store_config
from platforms.falabella import forms
from platforms.falabella import periods
from platforms.falabella import run_batch
from platforms.falabella import run_downloads

D = datetime.date


# ------------------------------------------------------------------ periods

def test_order_chunks_cover_the_month_in_14_day_pieces():
    assert periods.order_chunks("2026-09") == [
        (D(2026, 9, 1), D(2026, 9, 14)), (D(2026, 9, 15), D(2026, 9, 28)),
        (D(2026, 9, 29), D(2026, 9, 30))]
    feb = periods.order_chunks("2026-02")
    assert feb[-1] == (D(2026, 2, 15), D(2026, 2, 28))
    assert len(periods.order_chunks("2026-10")) == 3            # 14 + 14 + 3


def test_billing_periods_21st_to_20th():
    assert periods.billing_periods("2026-09") == [
        (D(2026, 8, 21), D(2026, 9, 20)), (D(2026, 9, 21), D(2026, 10, 20))]
    assert periods.billing_periods("2027-01") == [
        (D(2026, 12, 21), D(2027, 1, 20)), (D(2027, 1, 21), D(2027, 2, 20))]
    assert periods.billing_periods("2026-12")[1] == (D(2026, 12, 21), D(2027, 1, 20))


def test_spanish_row_labels_match_the_pages():
    a, b = periods.billing_periods("2026-09")[0]
    assert periods.billing_row_label(a, b) == "21 agosto 2026 - 20 septiembre 2026"
    assert periods.liquidation_row_label("2026-09") == "01 septiembre 2026 - 30 septiembre 2026"


def test_period_state():
    a, b = D(2026, 9, 21), D(2026, 10, 20)
    assert periods.billing_period_state(a, b, D(2026, 10, 16)) == "open"
    assert periods.billing_period_state(a, b, D(2026, 10, 21)) == "open"   # day of grace
    assert periods.billing_period_state(a, b, D(2026, 10, 22)) == "closed"


def test_form_ids_and_report_relevance():
    ids = periods.form_ids("2026-09")
    assert ids[:3] == ["order/2026-09-01_2026-09-14", "order/2026-09-15_2026-09-28",
                       "order/2026-09-29_2026-09-30"]
    assert "billing" in ids and "stock" in ids and "billing2" in ids and "invoice2" in ids
    assert "billing1/2026-09-21_2026-10-20" in ids
    assert not periods.in_report("invoice1/2026-08-21_2026-09-20")
    assert periods.in_report("billing1/2026-08-21_2026-09-20")
    assert periods.previous_month(D(2026, 1, 3)) == "2025-12"


def test_expected_only():
    e = run_downloads.expected("2026-09", ["billing1"])
    assert e == ["billing1/2026-08-21_2026-09-20", "billing1/2026-09-21_2026-10-20",
                 "invoice1/2026-08-21_2026-09-20", "invoice1/2026-09-21_2026-10-20"]
    assert run_downloads.expected("2026-09", ["stock"]) == ["stock"]


# ------------------------------------------------------------------ page facts

def test_order_chip():
    assert forms.chip_matches("Imprimir … Exportar\nsep 1, 2026 - sep 14, 2026",
                              D(2026, 9, 1), D(2026, 9, 14))
    assert not forms.chip_matches("sep 1, 2026 - sep 15, 2026", D(2026, 9, 1), D(2026, 9, 14))
    assert forms.chip_matches("dic 21, 2026 - ene 3, 2027", D(2026, 12, 21), D(2027, 1, 3))
    assert forms.chip_matches("Dec 21, 2026 - Jan 3, 2027", D(2026, 12, 21), D(2027, 1, 3))
    assert not forms.chip_matches("", D(2026, 9, 1), D(2026, 9, 14))


# ------------------------------------------------------------------ diagnosis

def test_collector_uses_page_locale_not_html_lang(tmp_path):
    c = run_downloads.Collector(str(tmp_path), "S")
    c.locale = "es-CL"

    class Drv:
        current_url = "https://sellercenter.falabella.com/order"
        title = "Seller Center"

        def execute_script(self, *a):
            return {"lang": "en", "body": "", "alerts": []}

        def save_screenshot(self, p):
            return False
    c.give_up(Drv(), "stock", "ui-missing", step="x", shot=False)
    rec = download_diag.diagnose_store(["stock"], collector=c)[0]
    assert rec["cause"] == "UI_NOT_FOUND"                     # not LANGUAGE_WRONG


def test_no_data_and_not_offered_need_no_attention(tmp_path):
    c = run_downloads.Collector(str(tmp_path), "S")
    c.locale = "es-CL"
    c.give_up(None, "order/2026-09-01_2026-09-14", "no-data", row_state="empty",
              status_text="no orders", shot=False)
    c.give_up(None, "billing1/2026-09-21_2026-10-20", "not-offered", step="open", shot=False)
    recs = download_diag.diagnose_store(
        ["order/2026-09-01_2026-09-14", "billing1/2026-09-21_2026-10-20"], collector=c)
    assert [r["cause"] for r in recs] == ["NO_DATA", "NOT_AVAILABLE"]
    assert not any(download_diag.needs_attention(r) for r in recs)


def test_language_wrong_is_classified(tmp_path):
    e = forms.LanguageWrong("界面语言是 en-US，不是西班牙语")
    ev = download_diag.exception_evidence(e, step="language")
    rec = download_diag.classify(ev, {"form": "billing"})
    assert rec["cause"] == "LANGUAGE_WRONG"


def test_pdfs_do_not_affect_the_report():
    assert not download_diag.affects_report({"form": "invoice2", "in_report": False})
    assert download_diag.affects_report({"form": "billing"})
    assert download_diag.form_of_entry({"form": "stock", "kind": "x"}) == "stock"


def test_owed_entry_is_still_generating():
    rec = download_diag.record_from_entry({"form": "billing", "last_state": "generating",
                                           "requested_at": 0, "probes": 1})
    assert rec["cause"] == "STILL_GENERATING"


# ------------------------------------------------------------------ batch

CFG = {"default": "BOCINA_SM",
       "allowed": ["BOCINA_SM", "DAB_CL-Fala (Coco)", "AIS_CL-Fala"],
       "catalog": {"BOCINA_SM": {"platform": "ML", "country": "MX"},
                   "DAB_CL-Fala (Coco)": {"platform": "FA", "country": "CL"},
                   "AIS_CL-Fala": {"platform": "FA", "country": "CL"}}}


@pytest.fixture
def cfg(monkeypatch):
    d = json.loads(json.dumps(CFG))
    monkeypatch.setattr(store_config, "_stores", lambda: d)
    return d


class Args:
    def __init__(self, stores=None, market="FA-CL"):
        self.stores, self.market = stores, market


def test_resolve_stores(cfg):
    assert run_batch.resolve_stores(Args()) == ["DAB_CL-Fala (Coco)", "AIS_CL-Fala"]
    assert run_batch.resolve_stores(Args(["AIS_CL-Fala"])) == ["AIS_CL-Fala"]
    with pytest.raises(store_config.StoreConfigError):
        run_batch.resolve_stores(Args(["BOCINA_SM"]))           # a MercadoLibre store
    with pytest.raises(store_config.StoreNotAllowed):
        run_batch.resolve_stores(Args(["FRIZ_CL-Fala"]))


def test_batch_cli_refuses_wrong_market(cfg, capsys):
    assert run_batch.main(["--market", "ML-CL", "--dry-run"]) == 4
    assert run_batch.main(["--stores", "BOCINA_SM", "--dry-run"]) == 4
    assert run_batch.main(["--dry-run", "--month", "2026-09"]) == 0
    assert "invoice2" in capsys.readouterr().out


def _rec(cause, form="billing"):
    return {"form": form, "cause": cause}


def test_batch_verdict():
    ok = {"status": "ok", "missing": [_rec("NO_DATA")], "errors": []}
    att = {"status": "partial", "missing": [_rec("UI_NOT_FOUND")], "errors": []}
    bad = {"status": "failed", "missing": [], "errors": ["x"]}
    assert run_batch.batch_verdict([ok])["exit_code"] == 0
    assert run_batch.batch_verdict([ok, att])["exit_code"] == 1
    assert run_batch.batch_verdict([ok, bad])["exit_code"] == 1
    assert run_batch.batch_verdict([bad])["exit_code"] == 2
    assert run_batch.batch_verdict([ok, att])["missing_by_cause"] == {"NO_DATA": 1,
                                                                     "UI_NOT_FOUND": 1}


# ------------------------------------------------------------------ language gate

class _Ctx(forms.Ctx):
    def __init__(self):
        from platforms.falabella import site
        forms.Ctx.__init__(self, driver=None, out_dir=".", diag=None, month="2026-09",
                           site=site.get("CL"), log=lambda *a: None)

    def go(self, path, settle=0):
        pass


def test_language_gate_switches_english_back_to_spanish(monkeypatch):
    from platforms.falabella import page_kit as pk
    seen = iter([{"locale": "en-CL", "seller": "S", "name": "N"},
                 {"locale": "es-CL", "seller": "S", "name": "N"}])
    monkeypatch.setattr(pk, "page_locale", lambda d: next(seen))
    calls = []
    monkeypatch.setattr(pk, "switch_language",
                        lambda d, opt: calls.append(opt) or {"locale": "es-CL"})
    ctx = _Ctx()
    forms.check_language(ctx)
    assert calls == ["Español"]
    assert ctx.flags["language_switched"] == {"from": "en-CL", "to": "es-CL"}
    assert ctx.flags["locale"] == "es-CL"


def test_language_gate_gives_up_when_the_switch_fails(monkeypatch):
    from platforms.falabella import page_kit as pk
    monkeypatch.setattr(pk, "page_locale", lambda d: {"locale": "en-CL"})
    monkeypatch.setattr(pk, "switch_language", lambda d, opt: None)
    with pytest.raises(forms.LanguageWrong) as e:
        forms.check_language(_Ctx())
    rec = download_diag.classify(download_diag.exception_evidence(e.value, step="language"),
                                 {"form": "billing"})
    assert rec["cause"] == "LANGUAGE_WRONG"


def test_language_gate_leaves_spanish_alone(monkeypatch):
    from platforms.falabella import page_kit as pk
    monkeypatch.setattr(pk, "page_locale", lambda d: {"locale": "es-CL"})
    monkeypatch.setattr(pk, "switch_language", lambda d, opt: pytest.fail("must not switch"))
    ctx = _Ctx()
    forms.check_language(ctx)
    assert "language_switched" not in ctx.flags


def test_direct_stock_fetch_retries_then_reports_a_landing_problem(tmp_path, monkeypatch):
    import urllib.request
    href = forms.STOCK_HOSTS[0] + "fby-exports/Product_details_1.xlsx"
    calls = []

    class Resp(object):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def read(self):
            return b"xlsx-bytes"

    def flaky(req, timeout=None):
        calls.append(1)
        if len(calls) < 3:
            raise urllib.error.URLError("[Errno 11001] getaddrinfo failed")
        return Resp()
    monkeypatch.setattr(urllib.request, "urlopen", flaky)
    monkeypatch.setattr(forms.time, "sleep", lambda s: None)
    ctx = _Ctx()
    ctx.out_dir = str(tmp_path)
    assert forms.fetch_direct(ctx, href) == ["Product_details_1.xlsx"]
    assert len(calls) == 3 and ctx.flags["stock_fetched_direct"]

    def dead(req, timeout=None):
        raise urllib.error.URLError("[Errno 11001] getaddrinfo failed")
    monkeypatch.setattr(urllib.request, "urlopen", dead)
    with pytest.raises(OSError):
        forms.fetch_direct(ctx, href)
    # the stock route turns that into "file did not land" (retry helps), not a script bug
    c = run_downloads.Collector(str(tmp_path), "S")
    c.locale = "es-CL"
    c.give_up(None, "stock", "download-not-landed", step="x", shot=False)
    assert download_diag.diagnose_store(["stock"], collector=c)[0]["cause"] == "DOWNLOAD_NOT_LANDED"


def test_missing_invoice_pdfs_by_type_and_by_number():
    # AIS: numbers filled in
    rows = [("520878", "Factura"), ("505871", "Nota de Credito"), (505872.0, "Nota de crédito"),
            ("1", "Nota de Debito")]
    assert forms.missing_invoice_pdfs(rows, ["520878_FACTURA.pdf", "505871_NOTA_DE_CREDITO.pdf"]) \
        == ["505872_NOTA_DE_CREDITO.pdf"]
    # DAB: the number column is empty — check by type (production 2026-10-19: the
    # factura was missing and nobody noticed)
    rows = [("", "Factura"), ("", "Nota de Credito"), ("", "Nota de Debito")]
    assert forms.missing_invoice_pdfs(rows, ["505871_NOTA_DE_CREDITO.pdf"]) == ["FACTURA"]
    assert forms.missing_invoice_pdfs(rows, ["520878_FACTURA.pdf",
                                             "505871_NOTA_DE_CREDITO.pdf"]) == []
    assert forms.missing_invoice_pdfs([("", "Factura")], []) == ["FACTURA"]
    assert forms.missing_invoice_pdfs([], []) == []


# ------------------------------------------------------------------ robustness (2026-10 review)

class _Drv(object):
    """Just enough driver for the language gate: an address and the page-load timeout."""
    def __init__(self, url):
        self.current_url = url
        self.caps = []

        class T(object):
            page_load = 300
        self.timeouts = T()

    def set_page_load_timeout(self, s):
        self.caps.append(s)


def test_language_gate_signs_in_when_the_first_page_hangs_on_the_login_loop(monkeypatch):
    from selenium.common.exceptions import TimeoutException
    from platforms.falabella import login
    from platforms.falabella import page_kit as pk
    sso = "https://access-key-corp.falabella.tech/auth/realms/x/protocol/openid-connect/auth?a=1"
    drv = _Drv(sso)
    gone = []

    class C(_Ctx):
        def go(self, path, settle=0):
            gone.append(path)
            if len(gone) == 1:
                raise TimeoutException("timeout: Timed out receiving message from renderer")
    ctx = C()
    ctx.d = drv
    signed = []
    monkeypatch.setattr(login, "sign_in", lambda c: signed.append(1) or True)
    monkeypatch.setattr(pk, "page_locale", lambda d: {"locale": "es-CL"})
    forms.check_language(ctx)
    assert signed == [1] and len(gone) == 2
    assert drv.caps == [forms.FIRST_LOAD_CAP, 300]          # capped, then put back


def test_language_gate_slow_page_that_is_not_login_just_goes_on(monkeypatch):
    from selenium.common.exceptions import TimeoutException
    from platforms.falabella import login
    from platforms.falabella import page_kit as pk

    class C(_Ctx):
        def go(self, path, settle=0):
            raise TimeoutException("slow")
    ctx = C()
    ctx.d = _Drv("https://sellercenter.falabella.com" + forms.PAYMENTS)
    monkeypatch.setattr(login, "sign_in", lambda c: pytest.fail("not a login page"))
    monkeypatch.setattr(pk, "page_locale", lambda d: {"locale": "es-CL"})
    forms.check_language(ctx)
    assert ctx.flags["locale"] == "es-CL"


def test_late_file_of_the_previous_chunk_is_not_this_chunk():
    before = {"Ordenes_a.xlsx", "orders_export.xlsx.crdownload"}
    ours, late = forms.own_new_files(before, ["orders_export (1).xlsx", "orders_export.xlsx"])
    assert ours == ["orders_export (1).xlsx"] and late == ["orders_export.xlsx"]
    assert forms.own_new_files(set(), ["a.xlsx"]) == (["a.xlsx"], [])


class _Landing(_Ctx):
    def __init__(self, out_dir):
        _Ctx.__init__(self)
        self.out_dir = out_dir
        self.gave, self.renamed = [], []

    def give_up(self, form, signal, **kw):
        self.gave.append((form, signal))

    def rename(self, name, new):
        self.renamed.append((name, new))
        return new


def test_orders_chunk_refuses_to_guess_between_two_new_files(tmp_path):
    a, b = D(2026, 9, 15), D(2026, 9, 28)
    ctx = _Landing(str(tmp_path))
    forms.land_order_export(ctx, "order/x", a, b, set(), ["x.xlsx", "y.xlsx"], 0)
    assert ctx.gave == [("order/x", "file-rejected")] and not ctx.renamed
    c = run_downloads.Collector(str(tmp_path), "S")
    c.locale = "es-CL"
    c.give_up(None, "order/x", "file-rejected", files=["x.xlsx", "y.xlsx"], step="2", shot=False)
    assert download_diag.diagnose_store(["order/x"], collector=c)[0]["cause"] == "FILE_REJECTED"
    # one file: named for the period
    ctx = _Landing(str(tmp_path))
    forms.land_order_export(ctx, "order/x", a, b, set(), ["x.xlsx"], 0)
    assert ctx.renamed == [("x.xlsx", "Ordenes_2026-09-15_2026-09-28.xlsx")]


def test_orders_chunk_waits_for_its_own_file_after_a_late_one(tmp_path, monkeypatch):
    from platforms.falabella import page_kit as pk
    a, b = D(2026, 9, 15), D(2026, 9, 28)
    seen = []
    monkeypatch.setattr(pk, "wait_new_files",
                        lambda out, before, **k: seen.append(set(before)) or ["mine.xlsx"])
    ctx = _Landing(str(tmp_path))
    forms.land_order_export(ctx, "order/x", a, b, {"prev.xlsx.crdownload"}, ["prev.xlsx"], 0)
    assert "prev.xlsx" in seen[0]
    assert ctx.renamed == [("mine.xlsx", "Ordenes_2026-09-15_2026-09-28.xlsx")]


def _epoch(*a):
    return datetime.datetime(*a, tzinfo=datetime.timezone.utc).timestamp()


OCT = "NewReportTransaction_2026-10-05T12_00_00.xlsx 05-10-2026 Listo"
OCT_LATER = "NewReportTransaction_2026-10-06T09_00_00.xlsx 06-10-2026 Listo"
SEP = "NewReportTransaction_2026-09-05T12_00_00.xlsx 05-09-2026 Listo"


def test_owed_transacciones_takes_only_a_row_made_after_the_request():
    asked = _epoch(2026, 10, 5, 11, 59)
    # never seen while waiting: the row that came after the one newest before
    e = {"route": "billing", "row": None, "seen_before": SEP, "requested_at": asked}
    assert forms.owed_tx_row(e, [OCT, SEP]) == OCT
    # nothing new yet: the old September report must not be taken (it was before)
    assert forms.owed_tx_row(e, [SEP]) is None
    # the table was empty before; an old report listed now is still too old
    e2 = dict(e, seen_before=None)
    assert forms.owed_tx_row(e2, [SEP]) is None
    # a later run's request above ours: ours is the oldest of the new ones
    assert forms.owed_tx_row(e, [OCT_LATER, OCT, SEP]) == OCT
    # seen while waiting: by file name, as before
    e3 = {"route": "billing", "row": "NewReportTransaction_2026-10-05T12_00_00.xlsx Procesando"}
    assert forms.owed_tx_row(e3, [OCT_LATER, OCT]) == OCT
    # entries written before seen_before existed behave as before
    assert forms.owed_tx_row({"route": "billing", "row": None}, [SEP]) == SEP


def test_an_older_request_turning_listo_is_not_the_new_one(monkeypatch):
    from platforms.falabella import page_kit as pk
    asked = _epoch(2026, 10, 5, 11, 59)
    before = "NewReportTransaction_2026-10-05T11_50_00.xlsx 05-10-2026 Procesando"
    now = "NewReportTransaction_2026-10-05T11_50_00.xlsx 05-10-2026 Listo"
    assert forms.new_tx_rows([now], before, asked) == []
    assert forms.new_tx_rows([OCT, now], before, asked) == [OCT]
    # downloads table not rendered when "before" was read: an old Listo row is not new
    monkeypatch.setattr(pk, "download_rows", lambda d: [SEP])
    monkeypatch.setattr(pk, "refresh_downloads", lambda d: None)
    monkeypatch.setattr(pk.time, "sleep", lambda s: None)
    row, state = pk.wait_new_ready_row(None, None, timeout=0, prefix="NewReportTransaction",
                                       select=lambda rows: forms.new_tx_rows(rows, None, asked))
    assert (row, state) == (None, "absent")


def test_row_timestamp_reads_both_spellings():
    from platforms.falabella import page_kit as pk
    assert pk.row_timestamp(OCT) == _epoch(2026, 10, 5, 12, 0, 0)
    assert pk.row_timestamp("NewReportTransaction_2026-10-16T07:24:06.469Z.xlsx Listo") == \
        _epoch(2026, 10, 16, 7, 24, 6)
    assert pk.row_timestamp("FSC_LIQ_2026-09-01_2026-09-30.xlsx Listo") is None
    assert pk.row_not_older_than("FSC_LIQ_x.xlsx", _epoch(2026, 10, 5))     # no time: no verdict


def test_owed_stock_takes_only_an_export_newer_than_the_request():
    e = {"route": "stock", "seen_before": "https://h/fby-exports/Product_details_1.xlsx"}
    old = ("Products Export Finalizado", "https://h/fby-exports/Product_details_1.xlsx")
    new = ("Products Export Finalizado", "https://h/fby-exports/Product_details_2.xlsx")
    assert forms.owed_stock_row(e, [new, old]) == new
    assert forms.owed_stock_row(e, [old]) is None
    assert forms.owed_stock_row({"route": "stock"}, [old]) == old       # old entry: as before
