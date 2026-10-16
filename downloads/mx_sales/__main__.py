"""Command line entry point: ``python -m mx_sales <command>``.

Commands
--------
``discover``  list raw files and the report type each one matched
``build``     clean the raw downloads to Parquet and refresh the DuckDB views
``query``     run SQL against the warehouse
``validate``  apply the Ventas MX validation rules to one export

Exit codes
----------
``0``  everything processed
``1``  some files failed; the rest were written and the warehouse refreshed
``2``  nothing usable -- every file failed, or none matched a report type

1 and 2 are kept apart on purpose. ``pipeline.run`` already isolates a bad
export ("one bad export must not sink the run"), but this used to collapse any
error into a single non-zero code, so a caller had no way to tell 1 bad file out
of 55 from a run that produced nothing -- and reported both as "cleaning
failed".

``build`` walks the whole raw tree every time, so a file that cannot be cleaned
fails every later build too. ``--current-run <folder name>`` says which run
folders are the ones just downloaded: failures elsewhere are then printed as a
short note instead of ``ERROR`` lines, and the ``[BUILD_RESULT_JSON]`` line marks
each failed file ``current`` or not. The exit code does not change with the
flag -- 1 still means "some file failed" -- so that a caller which does not pass
it sees exactly what it always did; the caller that does pass it reads the
split from the JSON line.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import traceback
from pathlib import Path

import pandas as pd

from mx_sales import pipeline, warehouse
from mx_sales.accounting import runner
from mx_sales.accounting import ventas_mx as rules
from mx_sales.config import SETTINGS
from mx_sales.registry import SPECS


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="mx-sales", description=__doc__)
    parser.add_argument("-v", "--verbose", action="store_true", help="log per-file progress")
    subparsers = parser.add_subparsers(dest="command", required=True)

    discover_parser = subparsers.add_parser("discover", help="list recognized raw downloads")
    discover_parser.add_argument("--unmatched", action="store_true", help="show unrecognized files too")

    build_parser = subparsers.add_parser("build", help="clean raw downloads to Parquet")
    build_parser.add_argument("--report", help="limit to one report, e.g. ventas_mx")
    build_parser.add_argument("--file", help="process a single download (name or relative path)")
    build_parser.add_argument("--no-warehouse", action="store_true", help="skip DuckDB view refresh")
    build_parser.add_argument(
        "--current-run", action="append", default=None, metavar="FOLDER",
        help="name of a run folder downloaded just now (repeatable); failures in any "
             "other folder are reported as older ones, in a note",
    )

    query_parser = subparsers.add_parser("query", help="run SQL against the warehouse")
    query_parser.add_argument("sql", help="SQL to execute, e.g. 'SELECT * FROM ventas_mx LIMIT 5'")

    validate_parser = subparsers.add_parser("validate", help="apply the Ventas MX rules")
    validate_parser.add_argument("file", help="the sales export to validate (name or path)")
    validate_parser.add_argument("--excel", type=Path, help="also write an .xlsx workbook here")

    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s %(message)s",
    )
    pd.set_option("display.width", 200)

    if args.command == "discover":
        return _discover(args.unmatched)
    if args.command == "build":
        return _build(args.report, args.file, warehouse_refresh=not args.no_warehouse,
                      current_runs=args.current_run)
    if args.command == "query":
        print(warehouse.query(args.sql).to_string(index=False))
        return 0
    if args.command == "validate":
        return _validate(args.file, args.excel)
    return 2          # unreachable (subparsers are required); never borrow 1


def _discover(show_unmatched: bool) -> int:
    items = pipeline.discover()
    print(f"raw dir: {SETTINGS.raw_dir}")
    print(f"{len(items)} file(s) matched a registered report:\n")
    for item in items:
        print(f"  {item.spec.name:<14} {item.path.relative_to(SETTINGS.raw_dir)}")

    if show_unmatched:
        matched = {item.path for item in items}
        unmatched = [
            p
            for p in sorted(SETTINGS.raw_dir.rglob("*"))
            if p.is_file()
            and p.suffix.lower() in pipeline.RAW_SUFFIXES
            and p not in matched
            and not (set(p.relative_to(SETTINGS.raw_dir).parts) & {"data", "mx_sales"})
        ]
        print(f"\n{len(unmatched)} file(s) with no registered report:\n")
        for path in unmatched:
            print(f"  {'-':<14} {path.relative_to(SETTINGS.raw_dir)}")

    print(f"\nregistered reports: {', '.join(spec.name for spec in SPECS)}")
    return 0


#: Prefix of the one machine-readable line ``build`` prints. ``transform.py`` at the
#: repository root reads it; keep the two in step.
BUILD_RESULT_MARKER = "[BUILD_RESULT_JSON] "


def _build(
    report: str | None,
    only: str | None,
    *,
    warehouse_refresh: bool,
    current_runs: list[str] | None = None,
) -> int:
    results = pipeline.run(report, only=only)
    if not results:
        print("nothing to do: no raw files matched a registered report")
        return 2

    # `files` counts only the files that made it through -- pipeline.run skips a
    # failed one before incrementing -- so errors and files never double count.
    written = sum(result.files for result in results)
    failed = sum(len(result.errors) for result in results)

    # Without --current-run every failure is this run's, which is what it always meant.
    # A full path is accepted too: only the folder name identifies the run.
    runs = None if current_runs is None else frozenset(
        name for name in (run.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1]
                          for run in current_runs) if name
    )
    failures: list[dict] = []
    for result in results:
        for relative, error in zip(result.failed_files, result.errors):
            failures.append({
                "report": result.report,
                "file": relative,
                "error": error,
                "current": runs is None or pipeline.from_runs(relative, runs),
            })
    older = [f for f in failures if not f["current"]]

    for result in results:
        note = f", {len(result.errors)} FAILED" if result.errors else ""
        print(f"{result.report}: {result.rows} rows from {result.files} file(s)"
              f"{note} -> {result.output}")
        for failure in failures:
            if failure["report"] == result.report and failure["current"]:
                print(f"  ERROR {failure['error']}")

    if older:
        # Not ERROR lines: nothing downloaded just now is wrong, and a file that fails
        # on every build would otherwise make every month look broken.
        print(f"note: {len(older)} file(s) from earlier runs still cannot be cleaned "
              "(not part of this run):")
        for failure in older:
            print(f"  {failure['report']:<14} {failure['file']}: "
                  f"{failure['error'].split(': ', 1)[-1]}")

    if warehouse_refresh:
        views = warehouse.refresh_views()
        print(f"warehouse: {SETTINGS.warehouse_path} (views: {', '.join(views) or 'none'})")

    print(BUILD_RESULT_MARKER + json.dumps(
        {"written": written, "failed": failures, "current_runs": sorted(runs or [])},
        ensure_ascii=False,
    ))

    if not failed:
        return 0
    total = written + failed
    if written:
        split = f" ({failed - len(older)} from this run, {len(older)} older)" if runs is not None else ""
        print(f"{failed} of {total} file(s) failed to clean{split}; the other {written} "
              f"were written and the warehouse was refreshed")
        return 1
    print(f"all {total} file(s) failed to clean; nothing was written")
    return 2


def _validate(file: str, excel: Path | None) -> int:
    source, frame, charges = runner.prepare(file)
    store = runner.store_of(source)
    output = runner.write(frame, store=store)
    failed = rules.exceptions(frame)

    print(f"{store or '(store unknown)'}: {len(frame)} rows -> {output}\n")
    print("control -- money in the sheet vs money in the output:")
    print(rules.control_totals(source, frame).to_string(index=False))
    print()
    print(rules.summarize(frame).to_string(index=False))
    print()
    print(rules.monthly_summary(frame, charges).to_string(index=False))

    if not failed.empty:
        columns = [
            "row_index", "order_id", "sold_at", "order_status",
            "components_sum_mxn", "total_mxn", "total_diff_mxn",
        ]
        print(f"\nrule 1 differences ({len(failed)}):")
        print(failed[columns].to_string(index=False))

    if excel is not None:
        print(f"\nworkbook: {runner.to_excel(frame, excel, source, charges, export=file, store=store)}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:
        # 1 now means exactly one thing -- "some files failed, the rest were
        # written" -- and callers key off that. An unhandled crash (a bad
        # --file, an unreadable warehouse) must not borrow that code, or a total
        # failure gets reported as a partial one. Traceback still printed: this
        # widens the exit code, it does not hide the error.
        traceback.print_exc()
        sys.exit(2)
