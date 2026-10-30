# -*- coding: utf-8 -*-
"""The figures of a period report: final monthly results, gathered and added up.

Input is only core/results (final() per store-month, versions() to say why a
store-month has no final). Nothing here reads platform files or workbooks.

    Period          kind ('quarterly' / 'annual' / 'custom'), label, months
    gather()        one market × basis × period -> PeriodData
    PeriodData      cells {(store, month): Cell}; the P&L spec and the sums of
                    every section; what is missing and why; which months are
                    before the store's first data month

A store's first data month is the earliest month with any results file for it
(any basis, any grade). A store-month before it is 「未上线」 (not started: the
system was not producing results yet), not missing — it is left out of the
totals like any month without a final, but it is not a gap anyone can fill, so
missing() leaves it out and not_started() lists it. Only a month on or after
the first data month without a final version is missing.

Money is added as the decimal numbers the results files hold (Decimal of each
figure's shortest repr) and rounded half-up to the currency's decimals once
per total, so a total equals the sum of the final monthly results to the cent
whatever binary noise the figures carry.
"""
import collections
import decimal
import math

from core import markets
from core import results as R

Period = collections.namedtuple("Period", "kind label months")

KIND_CN = {"quarterly": "季度", "annual": "全年", "custom": "自定义期间"}


# ------------------------------------------------------------------ periods

def quarter(text):
    months = R.quarter_months(text)
    return Period("quarterly", R.period_label("quarterly", months), months)


def year(text):
    y = str(text).strip()
    if not (len(y) == 4 and y.isdigit()):
        raise ValueError("年份写法应为 2026：%r" % text)
    months = R.year_months(int(y))
    return Period("annual", R.period_label("annual", months), months)


def custom(first, last):
    for m in (first, last):
        if not R.MONTH_RE.match(str(m or "")) or not 1 <= int(str(m)[5:7]) <= 12:
            raise ValueError("月份写法应为 2026-07：%r" % m)
    months = R.months_between(first, last)
    return Period("custom", R.period_label("custom", months), months)


def period_cn(p):
    """'2026 年第 3 季度（2026-07 至 2026-09）' / '2026 年全年' / '2026-07 至 2026-11'."""
    span = "%s 至 %s" % (p.months[0], p.months[-1])
    if p.kind == "quarterly":
        return "%s 年第 %s 季度（%s）" % (p.label[:4], p.label[-1], span)
    if p.kind == "annual":
        return "%s 年全年（%s）" % (p.label, span)
    return "%s（%d 个月）" % (span, len(p.months))


def basis_of(platform, basis=None):
    """The basis a market's report uses: ML bill (default) or order; main elsewhere."""
    if platform != "ML":
        return "main"
    return basis if basis in ("bill", "order") else R.official_basis(platform)


def basis_cn(basis):
    return R.BASIS_CN.get(basis, "")


# ------------------------------------------------------------------ cells

NOT_STARTED = "未上线"


def not_started_text(first_month):
    """'未上线（系统从 2026-09 起有数据）' — the wording of a not-started store-month."""
    return "%s（系统从 %s 起有数据）" % (NOT_STARTED, first_month)


class Cell(object):
    """One store-month: its final version (or None), its newest version of any
    grade, the files of the month that cannot be used, and whether the month is
    before the store's first data month (not_started; first_month says which)."""

    def __init__(self, store, month, final, newest, problems=(), first_month=None):
        self.store, self.month = store, month
        self.final, self.newest = final, newest
        self.problems = list(problems)
        self.first_month = first_month
        self.not_started = bool(first_month) and month < first_month and final is None

    @property
    def ok(self):
        return self.final is not None

    def reason(self):
        """Why there is no final version (Chinese), or '' when there is one.
        A not-started month says so: '未上线（系统从 2026-09 起有数据）'."""
        if self.final:
            return ""
        if self.not_started:
            return not_started_text(self.first_month)
        if self.newest:
            g = self.newest["grade"]
            why = "；".join(g.get("reasons") or []) or "没有说明原因"
            text = "最新版本（%s）的等级是「%s」：%s" % (
                self.newest["version"], R.GRADE_CN.get(g["level"], g["level"]), why)
        else:
            text = "这个月没有任何版本（没出过月度报表，或结果文件读不出来）"
        if self.problems:
            text += "；有 %d 个结果文件不能用：%s" % (len(self.problems), "；".join(
                "%s（%s）" % p for p in self.problems[:3]))
        return text


def _read(platform, country, store, month, basis, config_path=None):
    """(final, newest, problems) of one store-month. results.final() decides the
    final version; a store config.json no longer places in this market is read
    by its results folder with the same rule (results.final_at)."""
    from core import store_config
    try:
        same = store_config.market_parts(store) == (platform, country)
    except Exception:                                  # noqa: BLE001 - not in config any more
        same = False
    # Only a store-month without a final version reads its other versions and
    # broken files (to say why) — a year of thirty stores stays quick.
    if same:
        final = R.final(store, month, basis, config_path)
        if final:
            return final, final, []
        vs = R.versions(store, month, basis, config_path)
        try:
            probs = R.problems(store, month, config_path)
        except Exception:                              # noqa: BLE001
            probs = []
    else:
        final = R.final_at(platform, country, store, month, basis, config_path)
        if final:
            return final, final, []
        vs = R.versions_at(platform, country, store, month, basis, config_path)
        probs = []
    # this basis's broken files, and files whose name says nothing about a basis
    probs = [p for p in probs if ".%s." % basis in p[0] or not R.FILE_RE.match(p[0])]
    return final, (vs[0] if vs else None), probs


def first_month(platform, country, store, config_path=None):
    """The store's first data month (results.first_month_at): the earliest month
    with any results file, any basis, any grade; None when it has none."""
    return R.first_month_at(platform, country, store, config_path)


def market_stores(market, months, basis, stores=None, config_path=None):
    """The stores of a market's period report.

    Without `stores`: every store with any version of this basis in one of the
    months (the results tree under <PF>/<CC>). With `stores`: those of them that
    belong to this market (by the results tree, else by config.json), whether
    they have results or not — a listed store without any results anywhere is
    reported missing in every month (it has no first data month)."""
    pf, cc = markets.split(markets.normalize(market))
    found = R.stores(pf, cc, config_path)
    if stores:
        from core import store_config
        out = []
        for s in stores:
            if s in found:
                out.append(s)
                continue
            try:
                if store_config.market_parts(s) == (pf, cc):
                    out.append(s)
            except Exception:                          # noqa: BLE001
                pass
        return out
    return [s for s in found if any(R.has_files_at(pf, cc, s, m, basis, config_path) for m in months)]


# ------------------------------------------------------------------ sums

def fsum(values, decimals):
    """The sum of `values` to `decimals` places (half-up), as a float."""
    total = sum((decimal.Decimal(repr(float(v))) for v in values), decimal.Decimal(0))
    q = decimal.Decimal(1).scaleb(-int(decimals or 0))
    return float(total.quantize(q, rounding=decimal.ROUND_HALF_UP))


class PeriodData(object):
    """One market's period: the cells and every sum the workbook shows."""

    def __init__(self, market, basis, period, stores, cells, built):
        self.market = markets.normalize(market)
        self.platform, self.country = markets.split(self.market)
        self.basis, self.period = basis, period
        self.months = list(period.months)
        self.stores = list(stores)
        self.cells = cells
        self.built = built
        docs = [c.final for c in cells.values() if c.final] + \
               [c.newest for c in cells.values() if c.newest]
        self.currency = docs[0]["currency"] if docs else ""
        self.decimals = int(docs[0]["decimals"]) if docs else 2
        self._pl_spec = None

    # ---- cells

    def cell(self, store, month):
        return self.cells[(store, month)]

    def final(self, store, month):
        return self.cells[(store, month)].final

    def finals(self):
        """[(store, month, doc)] of the store-months that count."""
        return [(s, m, self.final(s, m)) for s in self.stores for m in self.months
                if self.final(s, m)]

    def missing(self):
        """[Cell] without a final version that should have one, store by store:
        months on or after the store's first data month. Not-started months
        (before it) are in not_started(), not here."""
        return [self.cell(s, m) for s in self.stores for m in self.months
                if not self.final(s, m) and not self.cell(s, m).not_started]

    def not_started(self):
        """[Cell] before the store's first data month (未上线), store by store."""
        return [self.cell(s, m) for s in self.stores for m in self.months
                if self.cell(s, m).not_started]

    def first_month(self, store):
        """The store's first data month ('YYYY-MM'), or None when it has no results."""
        return self.cell(store, self.months[0]).first_month

    def data_months(self, store):
        """(first data month, [months of the period before it]) for a store —
        ([], no first month) when none of its months is not started."""
        before = [m for m in self.months if self.cell(store, m).not_started]
        return (self.first_month(store), before) if before else (None, [])

    def not_started_stores(self):
        """Stores with at least one not-started month, in report order."""
        return [s for s in self.stores if self.data_months(s)[1]]

    def not_started_cn(self, store, short=False):
        """'系统从 2026-09 起有该店数据，2026-07、2026-08 不在合计内' — or, short,
        '系统从 2026-09 起有数据，07、08 月不在合计内'; '' when nothing is not started."""
        first, before = self.data_months(store)
        if not before:
            return ""
        if short:
            return "系统从 %s 起有数据，%s 月不在合计内" % (first, "、".join(m[5:7] for m in before))
        return "系统从 %s 起有该店数据，%s 不在合计内" % (first, "、".join(before))

    def complete_stores(self, month):
        return [s for s in self.stores if self.final(s, month)]

    def last_month_with_data(self):
        for m in reversed(self.months):
            if self.complete_stores(m):
                return m
        return None

    def sum(self, values):
        return fsum(values, self.decimals)

    # ---- ② P&L

    def pl_spec(self):
        """Every P&L row in the order of the latest month's results; rows only
        older months have follow, in the order they first appear."""
        if self._pl_spec is not None:
            return self._pl_spec
        seen = collections.OrderedDict()
        for m in reversed(self.months):
            for s in self.stores:
                doc = self.final(s, m)
                if not doc:
                    continue
                for row in doc["pl"]:
                    if row["key"] not in seen:
                        seen[row["key"]] = {"key": row["key"], "label": row.get("label") or row["key"],
                                            "original": row.get("original") or "",
                                            "kind": row.get("kind") or "d"}
        self._pl_spec = list(seen.values())
        return self._pl_spec

    def pl_value(self, store, month, key):
        """The store-month's figure; None when the store-month has no final
        version; 0 when the final version has no such row (pl_gaps lists it)."""
        doc = self.final(store, month)
        if not doc:
            return None
        return math.fsum(float(r["value"] or 0) for r in doc["pl"] if r["key"] == key)

    def pl_gaps(self):
        """[(store, month, key, label)]: a final version without one of the rows."""
        out = []
        for row in self.pl_spec():
            for s, m, doc in self.finals():
                if not any(r["key"] == row["key"] for r in doc["pl"]):
                    out.append((s, m, row["key"], row["label"]))
        return out

    def store_total(self, store, key):
        return self.sum(v for v in (self.pl_value(store, m, key) for m in self.months)
                        if v is not None)

    def month_total(self, month, key):
        return self.sum(v for v in (self.pl_value(s, month, key) for s in self.stores)
                        if v is not None)

    def total(self, key):
        return self.sum(v for s in self.stores for m in self.months
                        for v in [self.pl_value(s, m, key)] if v is not None)

    # ---- memo (② 备查项), fees (④), returns (⑤): keyed lists

    def _keyed(self, field, key_of, label_of):
        seen = collections.OrderedDict()
        for m in reversed(self.months):
            for s in self.stores:
                doc = self.final(s, m)
                for row in (doc or {}).get(field) or []:
                    k = key_of(row)
                    if k not in seen:
                        seen[k] = label_of(row)
        return seen

    def memo_spec(self):
        return self._keyed("memo", lambda r: r["key"], lambda r: r.get("label") or r["key"])

    def fee_spec(self):
        """{fee name: (label, pl_key)}."""
        return self._keyed("fees", lambda r: r["name"],
                           lambda r: (r.get("label") or r["name"], r.get("pl_key") or ""))

    def return_spec(self):
        return self._keyed("returns", lambda r: r["key"], lambda r: r.get("label") or r["key"])

    def field_value(self, store, month, field, key_field, key, value_field):
        doc = self.final(store, month)
        if not doc:
            return None
        return math.fsum(float(r.get(value_field) or 0) for r in doc.get(field) or []
                         if r.get(key_field) == key)

    def field_store_total(self, store, field, key_field, key, value_field, money=True):
        vals = [v for v in (self.field_value(store, m, field, key_field, key, value_field)
                            for m in self.months) if v is not None]
        return self.sum(vals) if money else int(round(math.fsum(vals)))

    def field_month_total(self, month, field, key_field, key, value_field, money=True):
        vals = [v for v in (self.field_value(s, month, field, key_field, key, value_field)
                            for s in self.stores) if v is not None]
        return self.sum(vals) if money else int(round(math.fsum(vals)))

    # ---- ⑦ SKU

    SKU_MONEY = ("net_sales", "fees", "contribution")

    def sku_rows(self, store):
        """One row per SKU over the period, biggest net sales first."""
        rows = collections.OrderedDict()
        for m in self.months:
            doc = self.final(store, m)
            for r in (doc or {}).get("sku") or []:
                sku = str(r.get("sku") or "")
                x = rows.setdefault(sku, {"sku": sku, "title": "", "units": 0.0, "units_returned": 0.0,
                                          "net_sales": [], "fees": [], "contribution": [], "cost": [],
                                          "months": set(), "cost_missing": set()})
                if r.get("title"):
                    x["title"] = r["title"]                  # the latest month's title wins
                x["units"] += float(r.get("units") or 0)
                x["units_returned"] += float(r.get("units_returned") or 0)
                for k in self.SKU_MONEY:
                    x[k].append(float(r.get(k) or 0))
                if r.get("cost") is None:
                    x["cost_missing"].add(m)
                else:
                    x["cost"].append(float(r["cost"]))
                x["months"].add(m)
        out = []
        for x in rows.values():
            row = dict(x)
            for k in self.SKU_MONEY:
                row[k] = self.sum(x[k])
            row["cost"] = self.sum(x["cost"]) if x["cost"] else None
            out.append(row)
        out.sort(key=lambda r: (-r["net_sales"], r["sku"]))
        return out

    # ---- ⑨ inventory

    def inventory(self, store):
        """(month, snapshot) of the latest month in the period whose final version
        has one; (None, None) when none has."""
        for m in reversed(self.months):
            doc = self.final(store, m)
            if doc and doc.get("inventory"):
                return m, doc["inventory"]
        return None, None

    # ---- ⑩ checks carried over

    def carried_checks(self):
        """[(store, month, check)]: checks that did not pass in the final versions."""
        out = []
        for s, m, doc in self.finals():
            for c in doc.get("checks") or []:
                if not c.get("ok") and not c.get("skipped"):
                    out.append((s, m, c))
        return out

    # ---- versions

    def version_map(self):
        """{(store, month): version id or ''} — what this report used."""
        return {(s, m): (self.final(s, m) or {}).get("version", "")
                for s in self.stores for m in self.months}


def gather(market, period, basis=None, stores=None, config_path=None, built=None):
    """Read the final monthly results of one market for a period."""
    import datetime
    market = markets.normalize(market)
    pf, cc = markets.split(market)
    basis = basis_of(pf, basis)
    names = market_stores(market, period.months, basis, stores, config_path)
    cells = {}
    for s in names:
        first = first_month(pf, cc, s, config_path)
        for m in period.months:
            final, newest, probs = _read(pf, cc, s, m, basis, config_path)
            cells[(s, m)] = Cell(s, m, final, newest, probs, first)
    return PeriodData(market, basis, period, names, cells, built or datetime.datetime.now())
