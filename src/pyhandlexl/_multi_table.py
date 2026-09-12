"""Internal support for multiple named tables on one worksheet.

Not part of the public API. A workbook-wide schema sheet (``_SCHEMA_SHEET``)
records where each named table lives. Each table is also self-describing on
the sheet: its very first cell holds the literal marker ``"TABLE NAME"`` and
the cell to its right holds the table's name. The schema is a fast-path cache;
the marker is how a lookup verifies the cache is still correct, and how it
recovers if a table has moved.

Layout of one named table, anchored at (anchor_row, anchor_col):

    anchor_row      TABLE NAME | <name>
    anchor_row + 1  <corner>   | <column headers ...>
    anchor_row + 2..           <row label> | <data ...>

Tables on the same sheet stack left-to-right with exactly one empty column
between them; they always start at row 1.
"""

from __future__ import annotations

from dataclasses import dataclass, replace

from pyhandlexl.errors import TableExistsError, TableNotFoundError

SCHEMA_SHEET = "_pyhandlexl_tables"
MARKER = "TABLE NAME"
_SCHEMA_HEADER = ("name", "sheet", "anchor_row", "anchor_col", "n_rows", "n_cols")


@dataclass(frozen=True)
class TableEntry:
    """One row of the schema: where a named table lives and how big it is."""

    name: str
    sheet: str
    anchor_row: int
    anchor_col: int
    n_rows: int  # data rows (not counting the marker or header row)
    n_cols: int  # data columns (not counting the label column)

    @property
    def height(self) -> int:
        """Total rows on the sheet: marker row + header row + data rows."""
        return self.n_rows + 2

    @property
    def width(self) -> int:
        """Total columns on the sheet: label column + data columns."""
        return self.n_cols + 1


# ---------------------------------------------------------------- the schema


def load_schema(workbook) -> dict[str, TableEntry]:
    """Read the schema sheet from an open workbook. Empty if it doesn't exist yet."""
    if SCHEMA_SHEET not in workbook.sheetnames:
        return {}
    ws = workbook[SCHEMA_SHEET]
    entries: dict[str, TableEntry] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        name, sheet, anchor_row, anchor_col, n_rows, n_cols = row[:6]
        entries[name] = TableEntry(
            name, sheet, int(anchor_row), int(anchor_col), int(n_rows), int(n_cols)
        )
    return entries


def save_schema(workbook, entries: dict[str, TableEntry]) -> None:
    """Replace the schema sheet's contents in an open workbook."""
    if SCHEMA_SHEET in workbook.sheetnames:
        ws = workbook[SCHEMA_SHEET]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = workbook.create_sheet(SCHEMA_SHEET)
    ws.append(list(_SCHEMA_HEADER))
    for e in entries.values():
        ws.append([e.name, e.sheet, e.anchor_row, e.anchor_col, e.n_rows, e.n_cols])


# ------------------------------------------------------------- region access


def read_region(ws, top: int, left: int, height: int, width: int) -> list[list[object]]:
    return [
        [ws.cell(row=r, column=c).value for c in range(left, left + width)]
        for r in range(top, top + height)
    ]


def write_region(ws, top: int, left: int, grid: list[list[object]]) -> None:
    # ws.cell(..., value=None) is a no-op in openpyxl (it means "no value given",
    # not "clear it"), so None must be assigned via .value directly.
    for i, row in enumerate(grid):
        for j, value in enumerate(row):
            ws.cell(row=top + i, column=left + j).value = value


def clear_region(ws, top: int, left: int, height: int, width: int) -> None:
    for r in range(top, top + height):
        for c in range(left, left + width):
            ws.cell(row=r, column=c).value = None


# --------------------------------------------------------- marker verify/heal


def verify_or_locate(workbook, entry: TableEntry) -> TableEntry:
    """Confirm *entry*'s recorded position still holds its marker.

    If it does not, scan the recorded sheet for the marker/name pair and
    return a corrected entry (position only — size is assumed unchanged).
    Raises TableNotFoundError if the marker cannot be found at all.
    """
    if entry.sheet not in workbook.sheetnames:
        raise TableNotFoundError(f"table {entry.name!r}: sheet {entry.sheet!r} no longer exists")
    ws = workbook[entry.sheet]
    if (
        ws.cell(entry.anchor_row, entry.anchor_col).value == MARKER
        and ws.cell(entry.anchor_row, entry.anchor_col + 1).value == entry.name
    ):
        return entry

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column):
            if ws.cell(r, c).value == MARKER and ws.cell(r, c + 1).value == entry.name:
                return replace(entry, anchor_row=r, anchor_col=c)
    raise TableNotFoundError(
        f"table {entry.name!r} was not found on sheet {entry.sheet!r} "
        "(its marker is missing — it may have been deleted or edited by hand)"
    )


# ------------------------------------------------------------------ lookups


def get_entry(entries: dict[str, TableEntry], name: str) -> TableEntry:
    try:
        return entries[name]
    except KeyError:
        raise TableNotFoundError(f"no table named {name!r}") from None


def check_not_exists(entries: dict[str, TableEntry], name: str) -> None:
    if name in entries:
        raise TableExistsError(f"table {name!r} already exists")


# ------------------------------------------------------------------ layout


def find_placement(entries: dict[str, TableEntry], sheet: str) -> tuple[int, int]:
    """Where a brand-new table on *sheet* should start: row 1, one column past the rest."""
    same_sheet = [e for e in entries.values() if e.sheet == sheet]
    if not same_sheet:
        return 1, 1
    rightmost_used = max(e.anchor_col + e.width - 1 for e in same_sheet)
    return 1, rightmost_used + 2  # +1 empty separator, +1 to land past it


def tables_to_shift(entries: dict[str, TableEntry], entry: TableEntry) -> list[TableEntry]:
    """Every table on the same sheet sitting to the right of *entry*, left to right."""
    return sorted(
        (e for e in entries.values() if e.sheet == entry.sheet and e.anchor_col > entry.anchor_col),
        key=lambda e: e.anchor_col,
    )


def shift_right(workbook, entries: dict[str, TableEntry], entry: TableEntry, growth: int) -> None:
    """Move every table to the right of *entry* on its sheet further right by *growth* columns.

    Reads every affected region into memory before clearing or writing
    anything, so overlapping moves can never clobber each other. Updates
    *entries* in place.
    """
    to_shift = tables_to_shift(entries, entry)
    if not to_shift:
        return
    ws = workbook[entry.sheet]
    contents = [
        (e, read_region(ws, e.anchor_row, e.anchor_col, e.height, e.width)) for e in to_shift
    ]
    for e, _ in contents:
        clear_region(ws, e.anchor_row, e.anchor_col, e.height, e.width)
    for e, grid in contents:
        new_col = e.anchor_col + growth
        write_region(ws, e.anchor_row, new_col, grid)
        entries[e.name] = replace(e, anchor_col=new_col)
