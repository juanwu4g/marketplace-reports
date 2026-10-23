"""Excel readers that cope with ML's decorated worksheets.

ML workbooks put a title block, an "updated at" stamp and help links above the real
table, so the header row sits at a different offset in almost every report. Several
reports also use a two-level header where the upper row groups columns and the lower
row names them -- and the lower names repeat (``Unidades`` appears under Ventas,
Devoluciones and Reclamos), so the group is what disambiguates them.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from mx_sales.clean.text import slugify


def _row_filled(row: pd.Series) -> int:
    return int(row.notna().sum())


def find_header_row(frame: pd.DataFrame, *, min_columns: int = 3, search_rows: int = 20) -> int:
    """Return the index of the densest row in the top of the sheet: the header.

    Title and note rows occupy one or two cells; the header spans the full table, so
    the first row reaching ``min_columns`` and holding the maximum fill is the header.
    """
    window = frame.head(search_rows)
    counts = [(_row_filled(row), -idx) for idx, row in window.iterrows()]
    best_count, negative_index = max(counts)
    if best_count < min_columns:
        raise ValueError(f"no header row with >= {min_columns} populated cells in first {search_rows} rows")
    return -negative_index


#: How far down the sheet the header is looked for. ML's title block is under ten rows;
#: a banner or two on top of that still leaves plenty of slack.
HEADER_SEARCH_ROWS = 20


def find_anchor_row(
    frame: pd.DataFrame, anchors: tuple[str, ...] | list[str], *, search_rows: int = HEADER_SEARCH_ROWS
) -> int | None:
    """Return the first row in the top of the sheet that holds every anchor column name.

    ML does not keep the header on a fixed row: it inserts a banner above the table when
    it has something to say ("No pudimos mostrar algunos datos...") and every row below
    moves down by one. A fixed offset then reads the group row as the header and the
    whole file fails on "expected columns missing". Anchoring on names the report cannot
    work without finds the header wherever it landed.

    Names are compared slugified, the same way the columns are named afterwards, so
    accents, case and stray whitespace in the export do not matter. ``None`` means no
    row matched -- the caller decides whether that is fatal.
    """
    wanted = {slugify(a) for a in anchors}
    if not wanted:
        return None
    for position in range(min(search_rows, len(frame))):
        present = {slugify(v) for v in frame.iloc[position] if pd.notna(v)}
        if wanted <= present:
            return position
    return None


def locate_header_row(
    path: Path | str,
    sheet_name: str | int,
    anchors: tuple[str, ...] | list[str],
    *,
    default: int,
    search_rows: int = HEADER_SEARCH_ROWS,
) -> int:
    """Header row of a sheet on disk, by anchor, falling back to ``default``.

    For callers that hand the row to ``pd.read_excel(header=...)`` themselves. Only the
    top of the sheet is read, so this costs little next to the full read that follows.
    """
    top = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=search_rows, dtype=object)
    found = find_anchor_row(top, anchors, search_rows=search_rows)
    return default if found is None else found


def _drop_empty_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Drop fully empty columns (several reports start at column B).

    Only columns are trimmed here: empty rows above the table are load-bearing,
    because callers address the group and header rows by position.
    """
    trimmed = frame.dropna(axis=1, how="all").reset_index(drop=True)
    trimmed.columns = range(trimmed.shape[1])
    return trimmed


def read_sheet(
    path: Path | str,
    sheet_name: str | int = 0,
    *,
    header_row: int | None = None,
    anchors: tuple[str, ...] | list[str] | None = None,
    trim_empty: bool = True,
) -> pd.DataFrame:
    """Read one sheet with a single header row, auto-detecting the offset when unset.

    With ``anchors`` the header is the first row carrying all of those column names, and
    ``header_row`` is only the fallback for a sheet where none does.
    """
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
    if trim_empty:
        raw = _drop_empty_columns(raw)

    if anchors:
        found = find_anchor_row(raw, anchors)
        if found is not None:
            header_row = found
    if header_row is None:
        header_row = find_header_row(raw)

    columns = [slugify(v) for v in raw.iloc[header_row]]
    data = raw.iloc[header_row + 1 :].reset_index(drop=True)
    data.columns = _deduplicate(columns)
    return data.dropna(axis=0, how="all").reset_index(drop=True)


def read_grouped_sheet(
    path: Path | str,
    sheet_name: str | int = 0,
    *,
    group_row: int,
    header_row: int,
    anchors: tuple[str, ...] | list[str] | None = None,
    trim_empty: bool = True,
) -> pd.DataFrame:
    """Read a sheet whose header spans two rows, naming columns ``<group>__<header>``.

    The group row is forward-filled because ML writes each group label only above its
    first column.

    With ``anchors`` the header is the first row carrying all of those names (matched on
    the lower, un-grouped name) and the group row keeps its distance above it: a banner
    pushes both down together. ``group_row`` / ``header_row`` are then the usual
    positions, used as given when no row matches.
    """
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
    if trim_empty:
        raw = _drop_empty_columns(raw)

    if anchors:
        found = find_anchor_row(raw, anchors)
        # A header on the very first rows would leave no room for its group row; that
        # is not this layout, so the usual positions stand.
        if found is not None and found - (header_row - group_row) >= 0:
            group_row, header_row = found - (header_row - group_row), found

    groups = raw.iloc[group_row].ffill()
    headers = raw.iloc[header_row]
    columns = [
        f"{slugify(g)}__{slugify(h)}" if pd.notna(g) and slugify(g) else slugify(h)
        for g, h in zip(groups, headers)
    ]

    data = raw.iloc[header_row + 1 :].reset_index(drop=True)
    data.columns = _deduplicate(columns)
    return data.dropna(axis=0, how="all").reset_index(drop=True)


def _deduplicate(columns: list[str]) -> list[str]:
    """Suffix any remaining duplicate names so the frame keeps unique columns."""
    seen: dict[str, int] = {}
    result: list[str] = []
    for name in columns:
        name = name or "unnamed"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result
