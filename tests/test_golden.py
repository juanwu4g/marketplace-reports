# -*- coding: utf-8 -*-
"""Golden files: real store-months rebuilt with the current code must equal the
accepted workbook cell by cell (values, formulas as text, number formats).

The data are real exports with customer details, so they live OUTSIDE the repo:

    $REPORT_GOLDEN_ROOT (default ~/report-golden)/
        mx/BOCINA_TA02/                 the store's download folder for 2026-09
        mx/BOCINA_TA02.baseline.json    the accepted workbook, dumped
        cl/DAB_CL/ … cl/DAB_CL.baseline.json
        …

Missing data -> the case is SKIPPED with a message (other machines, CI). Run
`pytest tests/test_golden.py -rs` to see which cases ran.

A diff here means a number or a cell changed. If finance accepted the change,
regenerate that baseline deliberately:

    python tests/test_golden.py --rebaseline cl/DAB_CL

The build date is the only cell that changes with the day; it is pinned.
"""
import datetime
import json
import os
import sys

import pytest
from openpyxl import load_workbook

if __name__ == "__main__":                              # run as a script from tests/
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from platforms.mercadolibre import report_build

ROOT = os.environ.get("REPORT_GOLDEN_ROOT", os.path.expanduser("~/report-golden"))
BUILD_DATE = "2026-10-14"

# case id -> (store name as in config.json, site, month)          MercadoLibre
# case id -> (store name, "<platform code>", month)                 other platforms
# (other platforms build with platforms.<pkg>.report_build.build(pairs, month, out,
#  build_date=...); the case folder is the store's download folder for the month)
CASES = {
    "cl-fa/DAB_CL-Fala": ("DAB_CL-Fala (Coco)", "FA", "2026-09"),
    "mx-az/SM-Amazon": ("SM-Amazon", "AZ", "2026-09"),
    "mx/BOCINA_TA02": ("BOCINA_TA02", "mx", "2026-09"),
    "cl/DAB_CL": ("DAB_CL (Coco) 子账号", "cl", "2026-09"),
    "cl/Tecnohub_CL": ("Tecnohub_CL (个人店)", "cl", "2026-09"),
    "pe/IDG_PE": ("IDG_PE", "pe", "2026-09"),
}


def dump(path):
    """{sheet: {cell: value}, sheet#fmt: {cell: number format}, _sheets: [...]}"""
    wb = load_workbook(path)
    out = {"_sheets": wb.sheetnames}
    for ws in wb.worksheets:
        cells, fm = {}, {}
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is not None:
                    if isinstance(v, (datetime.date, datetime.datetime)):
                        v = v.isoformat()
                    elif not isinstance(v, (int, float, str, bool)):
                        v = str(v)
                    cells[c.coordinate] = v
                if c.number_format != "General":
                    fm[c.coordinate] = c.number_format
        out[ws.title] = cells
        out[ws.title + "#fmt"] = fm
    return json.loads(json.dumps(out))


def build(case, out):
    store, site, month = CASES[case]
    if site.isupper():                                   # another platform's report
        import importlib
        from core import markets
        pkg = markets.platform(site).package
        rb = importlib.import_module(pkg + ".report_build")
        r = rb.build([(store, os.path.join(ROOT, case))], month, out, build_date=BUILD_DATE)
        return report_build.EXIT_OK if r["outcome"] == "ok" else report_build.EXIT_CHECKS
    # start from another site on purpose: the build must switch by the folder
    report_build.set_site("cl" if site != "cl" else "mx")
    rc = report_build.main(["--month", month, "--store", "%s=%s" % (store, os.path.join(ROOT, case)),
                            "--out", out, "--no-recalc", "--build-date", BUILD_DATE])
    report_build.set_site("mx")
    return rc


def diff(want, got):
    out = []
    for key in sorted(set(want) | set(got)):
        a, b = want.get(key), got.get(key)
        if isinstance(a, dict) and isinstance(b, dict):
            for c in sorted(set(a) | set(b)):
                if a.get(c) != b.get(c):
                    out.append("%s!%s: %r -> %r" % (key, c, a.get(c), b.get(c)))
        elif a != b:
            out.append("%s: %r -> %r" % (key, a, b))
    return out


@pytest.mark.parametrize("case", sorted(CASES))
def test_golden_workbook_identical(case, tmp_path):
    folder, baseline = os.path.join(ROOT, case), os.path.join(ROOT, case + ".baseline.json")
    if not (os.path.isdir(folder) and os.path.isfile(baseline)):
        pytest.skip("no golden data for %s under %s" % (case, ROOT))
    out = str(tmp_path / "golden.xlsx")
    rc = build(case, out)
    assert rc in (report_build.EXIT_OK, report_build.EXIT_CHECKS)
    with open(baseline, encoding="utf-8") as fh:
        want = json.load(fh)
    d = diff(want, dump(out))
    assert not d, "%s changed in %d cells, first ones:\n%s" % (case, len(d), "\n".join(d[:20]))


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--rebaseline" and sys.argv[2] in CASES:
        import tempfile
        case = sys.argv[2]
        out = os.path.join(tempfile.mkdtemp(), "golden.xlsx")
        print("build ->", build(case, out))
        with open(os.path.join(ROOT, case + ".baseline.json"), "w", encoding="utf-8") as fh:
            json.dump(dump(out), fh, ensure_ascii=False)
        print("baseline written for", case)
    else:
        print("usage: python tests/test_golden.py --rebaseline <%s>" % "|".join(sorted(CASES)))
