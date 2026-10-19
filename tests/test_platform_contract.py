# -*- coding: utf-8 -*-
"""The contract every platform package keeps with the runner and the window.

core/run_monthly.build_plan starts, as child processes (`python -m …`):

  MercadoLibre (units MX / CL / PE)
    run_batch   --month M [--stores …]           (+ --collect-only for the catch-up)
    run_reports --month M [--region R] [--stores … --no-combined] --out-dir D
  every other platform (units "<CODE>-<COUNTRY>", _build_plan_market)
    run_batch   --month M --market U --stores …  (+ --collect-only)
    run_reports --month M --market U [--stores … --no-combined] --out-dir D
  the window's preview (region_choice.preview_cmd)
    run_reports --month M --dry-run [--market U] [--stores …]

and reads their exit codes (0 fine · 1 look · 2 failed · 3 store not allowed ·
4 configuration). So this test runs each package's modules exactly as the child
process would (runpy, run_name="__main__", so the exception -> exit code mapping
of the `if __name__ == "__main__"` block is exercised too), always with
--dry-run or with arguments that must be refused before anything opens. Ziniao
is replaced by a trap; config.json is a stub.

Known difference, documented rather than changed: MercadoLibre's run_reports
does not check the platform of the stores it is given (its batch does). Its
main() also reads sys.argv instead of taking argv — the runner never imports
it, so only the command line is the contract.

Platform packages that are registered in core/markets.py but not importable yet
(being built on another branch) are skipped with the reason.
"""
import importlib
import runpy
import sys
import warnings

import pytest

from core import markets
from core import paths
from core import pending_store
from core import run_monthly
from core import store_config

MONTH = "2026-09"


def _importable(pkg):
    try:
        importlib.import_module(pkg)
        return True, ""
    except ImportError as e:
        return False, str(e)


def _params():
    out = []
    for p in markets.PLATFORMS.values():
        ok, why = _importable(p.package)
        marks = [] if ok else [pytest.mark.skip(
            reason="%s (%s) 还不能 import：%s" % (p.code, p.package, why))]
        out.append(pytest.param(p, id=p.code, marks=marks))
    return out


PLATFORMS = _params()


def own_store(p):
    return "CONTRACT_%s_1" % p.code


def other_platform(p):
    for q in markets.PLATFORMS.values():
        if q.code != p.code:
            return q
    return None


def config_for_all():
    """One store per registered platform, in its first country."""
    allowed, catalog, batches = [], {}, {}
    for p in markets.PLATFORMS.values():
        n = own_store(p)
        allowed.append(n)
        catalog[n] = {"platform": p.code, "country": p.countries[0]}
        batches[markets.market_id(p.code, p.countries[0])] = [n]
    return {"default": allowed[0], "allowed": allowed, "catalog": catalog,
            "batches": batches}


@pytest.fixture
def sandbox(tmp_path, monkeypatch):
    """Stub config, a temporary ROOT (logs, downloads, reports land there), and a
    Ziniao that fails the test if anyone tries to start it."""
    cfg = config_for_all()
    monkeypatch.setattr(store_config, "_stores", lambda: cfg)
    monkeypatch.setattr(paths, "ROOT", str(tmp_path))
    monkeypatch.setattr(pending_store, "prune", lambda *a, **k: [])
    import core.ziniao_client as zc

    class Trap(object):
        def __init__(self, *a, **k):
            raise AssertionError("a contract test must never start Ziniao")
    monkeypatch.setattr(zc, "ZiniaoClient", Trap)
    return tmp_path


def run_cli(module, argv, monkeypatch):
    """Run `python -m module argv…` in this process; return its exit code.
    An exception that escapes is returned as the exception (the test decides
    whether that is acceptable)."""
    monkeypatch.setattr(sys, "argv", [module] + list(argv))
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", RuntimeWarning)   # "found in sys.modules"
        try:
            rv = runpy.run_module(module, run_name="__main__", alter_sys=False)
        except SystemExit as e:
            return e.code if e.code is not None else 0
        except Exception as e:  # noqa: BLE001 — returned for the caller to judge
            return e
    return rv.get("__return__", 0)


def is_ml(p):
    return p.code == markets.DEFAULT_PLATFORM


def unit(p):
    return markets.unit_of_market(markets.market_id(p.code, p.countries[0]))


def plan(p, out_dir, stores=None):
    """The commands the runner would start for this platform's first market."""
    u = unit(p)
    return run_monthly.build_plan("both", MONTH, stores, out_dir, py="PY", region=u,
                                  all_stores=[own_store(p)]) + \
        run_monthly.build_plan("collect", MONTH, stores, out_dir, py="PY", region=u,
                               all_stores=[own_store(p)])


def runner_args(cmd):
    """['PY', '-u', '-m', 'pkg.mod', …] -> ('pkg.mod', […])."""
    i = cmd.index("-m")
    return cmd[i + 1], cmd[i + 2:]


# ------------------------------------------------------------------ the modules exist

@pytest.mark.parametrize("p", PLATFORMS)
def test_package_has_runner_modules(p):
    for mod in ("run_batch", "run_reports"):
        m = importlib.import_module("%s.%s" % (p.package, mod))
        assert callable(getattr(m, "main", None)), "%s.%s.main" % (p.package, mod)


@pytest.mark.parametrize("p", PLATFORMS)
def test_registry_entry_is_consistent(p):
    assert p.countries and all(c in markets.COUNTRIES for c in p.countries)
    assert p.package == "platforms." + p.package.split(".")[-1]
    assert markets.ziniao_matches(p.code, p.name_en)


# ------------------------------------------------------------------ what the runner starts

@pytest.mark.parametrize("p", PLATFORMS)
def test_every_command_the_runner_starts_parses_and_dry_runs(p, sandbox, monkeypatch, capsys):
    """Each command of build_plan, with --dry-run added, is accepted and opens
    nothing. A dry-run download exits 0; a dry-run report exits 0, or 2 when no
    store has data (there is none here) — and writes no workbook."""
    out = sandbox / "reports"
    for stores in (None, [own_store(p)]):
        for step in plan(p, str(out), stores):
            module, args = runner_args(step["cmd"])
            assert module.startswith(p.package + "."), step["cmd"]
            if not is_ml(p):
                assert "--market" in args
            code = run_cli(module, args + ["--dry-run"], monkeypatch)
            allowed = (0,) if module.endswith(".run_batch") else (0, 2)
            assert code in allowed, (step["key"], step["cmd"], code, capsys.readouterr().out)
    assert not out.exists() or not list(out.iterdir()), "dry-run wrote a workbook"


@pytest.mark.parametrize("p", PLATFORMS)
def test_window_preview_command_dry_runs(p, sandbox, monkeypatch):
    from core import region_choice as RC
    ml_script = ["-m", "platforms.mercadolibre.run_reports"]
    for stores in (None, [own_store(p)]):
        cmd = RC.preview_cmd("PY", ml_script, MONTH, stores, unit(p))
        module, args = runner_args(cmd)
        assert module == p.package + ".run_reports"
        assert run_cli(module, args, monkeypatch) in (0, 2)


@pytest.mark.parametrize("p", PLATFORMS)
def test_batch_flags(p, sandbox, monkeypatch):
    mod = p.package + ".run_batch"
    base = ["--month", MONTH, "--stores", own_store(p), "--dry-run"]
    if not is_ml(p):
        base += ["--market", markets.market_id(p.code, p.countries[0])]
    assert run_cli(mod, base, monkeypatch) == 0
    assert run_cli(mod, base + ["--collect-only"], monkeypatch) == 0


# ------------------------------------------------------------------ refusals

def _refused(code, expected, exc_types):
    return code == expected or isinstance(code, exc_types)


@pytest.mark.parametrize("p", PLATFORMS)
def test_batch_refuses_a_store_not_on_the_allowlist(p, sandbox, monkeypatch):
    args = ["--month", MONTH, "--stores", "NOT_ALLOWED_STORE", "--dry-run"]
    if not is_ml(p):
        args += ["--market", markets.market_id(p.code, p.countries[0])]
    code = run_cli(p.package + ".run_batch", args, monkeypatch)
    assert _refused(code, 3, store_config.StoreNotAllowed), code


@pytest.mark.parametrize("p", PLATFORMS)
def test_batch_refuses_another_platforms_store(p, sandbox, monkeypatch):
    q = other_platform(p)
    if q is None:
        pytest.skip("only one platform registered")
    args = ["--month", MONTH, "--stores", own_store(q), "--dry-run"]
    if not is_ml(p):
        args += ["--market", markets.market_id(p.code, p.countries[0])]
    code = run_cli(p.package + ".run_batch", args, monkeypatch)
    assert _refused(code, 4, store_config.StoreConfigError), code


@pytest.mark.parametrize("p", [x for x in PLATFORMS if x.values[0].code != "ML"])
def test_batch_refuses_another_platforms_market(p, sandbox, monkeypatch):
    q = other_platform(p)
    if q is None:
        pytest.skip("only one platform registered")
    code = run_cli(p.package + ".run_batch",
                   ["--month", MONTH, "--market", markets.market_id(q.code, q.countries[0]),
                    "--stores", own_store(p), "--dry-run"], monkeypatch)
    assert _refused(code, 4, (store_config.StoreConfigError, markets.UnknownMarket)), code


@pytest.mark.parametrize("p", PLATFORMS)
def test_reports_refuse_a_store_not_on_the_allowlist(p, sandbox, monkeypatch):
    args = ["--month", MONTH, "--stores", "NOT_ALLOWED_STORE", "--dry-run",
            "--out-dir", str(sandbox / "r")]
    if not is_ml(p):
        args += ["--market", markets.market_id(p.code, p.countries[0])]
    code = run_cli(p.package + ".run_reports", args, monkeypatch)
    assert _refused(code, 3, store_config.StoreNotAllowed), code


@pytest.mark.parametrize("p", [x for x in PLATFORMS if x.values[0].code != "ML"])
def test_reports_refuse_another_platforms_store(p, sandbox, monkeypatch):
    """Not asked of MercadoLibre's run_reports (see the module docstring)."""
    q = other_platform(p)
    code = run_cli(p.package + ".run_reports",
                   ["--month", MONTH, "--market", markets.market_id(p.code, p.countries[0]),
                    "--stores", own_store(q), "--dry-run", "--out-dir", str(sandbox / "r")],
                   monkeypatch)
    assert _refused(code, 4, store_config.StoreConfigError), code


# ------------------------------------------------------------------ expected forms

@pytest.mark.parametrize("p", [x for x in PLATFORMS if x.values[0].code != "ML"])
def test_periods_declare_the_expected_forms(p):
    """run_downloads and download_diag need, per month: every form id a store
    should yield (form_ids) and whether the report reads it (in_report), so a
    missing form is either an error or only a note. The runner does not import
    these; the platform's own run_downloads does. MercadoLibre predates the
    convention (download_diag.expected_forms holds its list)."""
    periods = importlib.import_module(p.package + ".periods")
    ids = periods.form_ids(MONTH)
    assert ids and all(isinstance(f, str) and f for f in ids)
    assert len(ids) == len(set(ids)), "form ids must be unique"
    assert any(periods.in_report(f) for f in ids)
    for f in ids:
        assert periods.in_report(f) in (True, False)
    assert periods.previous_month(__import__("datetime").date(2026, 10, 19)) == MONTH
