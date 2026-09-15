"""Public functions for reading and writing worksheets, preserving cell types."""

from __future__ import annotations

import csv
import os
from collections.abc import Iterable
from itertools import zip_longest
from pathlib import Path
from secrets import token_hex
from typing import Literal

from openpyxl import Workbook

from pyhandlexl import _multi_table as mt
from pyhandlexl._safety import atomic_save, safe_delete, safe_load
from pyhandlexl.errors import SheetNotFoundError
from pyhandlexl.validate import check_cell_value, check_dimensions, check_sheet_name

Orientation = Literal["rows", "columns"]

# Extensions openpyxl recognises as Excel workbooks.
_WORKBOOK_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})
_CSV_SUFFIX = ".csv"


def _is_csv_path(path: Path) -> bool:
    return path.suffix.lower() == _CSV_SUFFIX


def _check_cell_values(grid: list[list[object]]) -> None:
    """Raise CellTypeError on the first value that is not a type Excel can store."""
    for row in grid:
        for value in row:
            check_cell_value(value)


def _checked_grid(rows: Iterable[Iterable[object]]) -> list[list[object]]:
    """Materialise *rows* into a grid, raising CellTypeError on any bad value."""
    grid = [list(row) for row in rows]
    _check_cell_values(grid)
    return grid


def _read_csv_rows(path: Path) -> list[list[str]]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        return [list(row) for row in csv.reader(f)]


def _write_csv_atomic(
    path: Path, grid: Iterable[Iterable[object]], *, encoding: str = "utf-8"
) -> None:
    """Write *grid* to *path* as CSV, stringifying every value (``None`` becomes an
    empty field), replacing any existing content atomically."""
    tmp = path.parent / f".{path.stem}.{token_hex(6)}.tmp.csv"
    try:
        with tmp.open("w", newline="", encoding=encoding) as f:
            writer = csv.writer(f)
            for row in grid:
                writer.writerow("" if value is None else str(value) for value in row)
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)


def _append_csv(path: Path, grid: Iterable[Iterable[object]], *, encoding: str = "utf-8") -> None:
    """Append *grid* to the end of an existing CSV file, stringifying every value.

    Deliberately does not read the file's existing content first — that's
    the whole point (see append_rows's docstring for why). Writes directly
    in append mode instead of through the temp-file-then-replace dance
    ``_write_csv_atomic`` uses: a crash mid-write can leave a malformed
    trailing row, but can never touch or lose a single byte that was
    already there, so the stronger (and more expensive) guarantee isn't
    needed for an append.
    """
    needs_leading_newline = False
    if path.stat().st_size > 0:
        with path.open("rb") as f:
            f.seek(-1, os.SEEK_END)
            needs_leading_newline = f.read(1) != b"\n"

    with path.open("a", newline="", encoding=encoding) as f:
        if needs_leading_newline:
            f.write("\r\n")
        writer = csv.writer(f)
        for row in grid:
            writer.writerow("" if value is None else str(value) for value in row)


def create_workbook(path: str | Path, *, sheet: str = "Sheet") -> None:
    """Create a new empty .xlsx file with one worksheet.

    Files are never created implicitly — call this first. ``write_sheet``,
    ``append_rows``, ``create_sheet``, and ``Table.write`` all require the file
    to exist already.

    Raises:
        FileExistsError: something is already at *path*.
        SheetNameError: *sheet* is not a valid worksheet name.
    """
    check_sheet_name(sheet)
    path = Path(path)
    if path.exists():
        raise FileExistsError(f"{path} already exists")
    workbook = Workbook()
    workbook.active.title = sheet
    try:
        atomic_save(workbook, path)
    finally:
        workbook.close()


def create_csv(path: str | Path) -> None:
    """Create a new empty .csv file.

    Files are never created implicitly — call this first, the same as
    :func:`create_workbook` for .xlsx. ``write_sheet`` and ``append_rows``
    both require the file to exist already.

    Raises:
        FileExistsError: something is already at *path*.
        ValueError: *path* does not have a .csv extension.
    """
    path = Path(path)
    if not _is_csv_path(path):
        raise ValueError(f"not a .csv path: {path}")
    if path.exists():
        raise FileExistsError(f"{path} already exists")
    path.touch()


def delete_workbook(path: str | Path) -> None:
    """Delete a workbook file, retrying while it is locked.

    Raises:
        ValueError: *path* does not name an Excel workbook (.xlsx/.xlsm/.xltx/.xltm).
        FileNotFoundError: no file at *path*.
        FileLockedError: the file stayed locked (open in Excel) through every retry.
    """
    path = Path(path)
    if path.suffix.lower() not in _WORKBOOK_SUFFIXES:
        raise ValueError(f"not a workbook path (need one of {sorted(_WORKBOOK_SUFFIXES)}): {path}")
    safe_delete(path)


def read_sheet(
    path: str | Path,
    sheet: str | None = None,
    *,
    pad: bool = False,
) -> list[list[object]]:
    """Read a worksheet — or a .csv file — as a list of rows.

    For an **.xlsx** file, values come back as ``str``, ``int``, ``float``,
    ``bool``, ``datetime``, ``date``, ``time``, or ``timedelta``, with an
    empty cell as ``None``; trailing ``None`` values are trimmed from each
    row, so a fully empty row becomes ``[]``. For a **.csv** file every
    value is a plain ``str`` instead, exactly as written — CSV has no other
    type to preserve, and rows are read exactly as they are, with no
    trimming.

    Args:
        path: the .xlsx or .csv file.
        sheet: worksheet name, or ``None`` for the active sheet. Must be
            ``None`` for a .csv file — it has no sheets.
        pad: if true, right-pad every row with ``None`` (``""`` for a .csv
            file) to the length of the longest row, making the result
            rectangular.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file is not a readable .xlsx.
        SheetNotFoundError: *sheet* names a worksheet that does not exist.
        ValueError: *sheet* is given for a .csv file.
    """
    path = Path(path)
    if _is_csv_path(path):
        if sheet is not None:
            raise ValueError("sheet is not meaningful for a .csv file")
        if not path.exists():
            raise FileNotFoundError(f"no file at {path}")
        rows: list[list[object]] = _read_csv_rows(path)
        fill: object = ""
    else:
        workbook = safe_load(path)
        try:
            if sheet is None:
                worksheet = workbook.active
            elif sheet in workbook.sheetnames:
                worksheet = workbook[sheet]
            else:
                raise SheetNotFoundError(sheet)

            rows = []
            for raw_row in worksheet.iter_rows(values_only=True):
                row = list(raw_row)
                while row and row[-1] is None:
                    row.pop()
                rows.append(row)
        finally:
            workbook.close()
        fill = None

    if pad and rows:
        width = max(len(row) for row in rows)
        rows = [row + [fill] * (width - len(row)) for row in rows]
    return rows


def write_sheet(
    path: str | Path,
    rows: Iterable[Iterable[object]],
    sheet: str | None = None,
    *,
    orientation: Orientation = "rows",
) -> None:
    """Replace a worksheet's contents with *rows* — or a whole .csv file's.

    For an **.xlsx** file, other worksheets are left untouched, *sheet* is
    added if it does not exist, values are written with their type
    preserved (``None`` leaves the cell empty), and every value must be a
    type Excel can store (see :func:`pyhandlexl.check_cell_value`) within
    its row/column limits. For a **.csv** file, the entire file is
    replaced — there's no sheet to isolate a change to — every value is
    stringified with ``str()`` (``None`` becomes an empty field), and there
    is no size limit. Either way the file must already exist (see
    :func:`create_workbook` / :func:`create_csv`), and the write is atomic.

    Args:
        path: the .xlsx or .csv file.
        rows: an iterable of iterables of cell values.
        sheet: worksheet name, or ``None`` for the active sheet. Must be
            ``None`` for a .csv file.
        orientation: ``"rows"`` writes each inner iterable as a row;
            ``"columns"`` writes each inner iterable down a column.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNameError: *sheet* is not a valid worksheet name.
        CellTypeError: a value is not a type Excel can store (.xlsx only).
        DimensionError: the data exceeds the .xlsx row or column limits
            (.xlsx only — a .csv file has no size limit).
        ValueError: *orientation* is not ``"rows"`` or ``"columns"``, or
            *sheet* is given for a .csv file.
    """
    if orientation not in ("rows", "columns"):
        raise ValueError(f"orientation must be 'rows' or 'columns', got {orientation!r}")

    path = Path(path)
    grid = [list(row) for row in rows]
    if orientation == "columns":
        grid = [list(column) for column in zip_longest(*grid, fillvalue=None)]

    if _is_csv_path(path):
        if sheet is not None:
            raise ValueError("sheet is not meaningful for a .csv file")
        if not path.exists():
            raise FileNotFoundError(f"no file at {path}")
        _write_csv_atomic(path, grid)
        return

    _check_cell_values(grid)
    check_dimensions(len(grid), max((len(row) for row in grid), default=0))

    if sheet is not None:
        check_sheet_name(sheet)

    workbook = safe_load(path)
    try:
        name = sheet if sheet is not None else workbook.active.title
        if name in workbook.sheetnames:
            index = workbook.sheetnames.index(name)
            workbook.remove(workbook[name])
            worksheet = workbook.create_sheet(title=name, index=index)
        else:
            worksheet = workbook.create_sheet(title=name)

        for row in grid:
            worksheet.append(row)
        atomic_save(workbook, path)
    finally:
        workbook.close()


def append_rows(
    path: str | Path,
    rows: Iterable[Iterable[object]],
    sheet: str | None = None,
) -> None:
    """Append *rows* to the end of a worksheet — or a .csv file.

    This exists as its own function, rather than being
    ``write_sheet(path, grid.append_row(read_sheet(path), values))``, for
    one reason: performance. That composition would have to read *every*
    existing row into Python first, just to add a few more at the end. This
    function never does — it adds only the new rows, however large the
    existing sheet or file already is, at a cost proportional to *rows*
    alone rather than to everything already there.

    The two formats get to that guarantee differently, because they can:
    for an **.xlsx** file, ``append_rows`` loads the workbook (openpyxl's
    normal cost of opening one at all) but calls ``worksheet.append()``
    only for the new rows — the existing ones are never walked or
    materialised as Python values — then resaves the whole workbook
    atomically, the same as every other .xlsx write. A **.csv** file can do
    even better, since unlike a zip-based .xlsx it can be modified without
    rewriting it: this opens the file in append mode and writes only the
    new rows directly, touching none of the existing bytes at all. That
    makes a .csv append cheaper than an .xlsx one, but also means it isn't
    wrapped in the temp-file-then-replace safety every other write in this
    library gets — a crash mid-write can leave a malformed trailing row,
    but (unlike a failed *replace*) can never lose or corrupt a single byte
    that was already there.

    An empty *rows* is a no-op. For an **.xlsx** file, *sheet* is added if
    it does not exist, and values follow the same rules as
    :func:`write_sheet`. For a **.csv** file, every value is stringified
    with ``str()`` and there is no size limit. Either way the file must
    already exist (see :func:`create_workbook` / :func:`create_csv`).

    Args:
        path: the .xlsx or .csv file.
        rows: an iterable of iterables of cell values.
        sheet: worksheet name, or ``None`` for the active sheet. Must be
            ``None`` for a .csv file.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNameError: *sheet* is not a valid worksheet name.
        CellTypeError: a value is not a type Excel can store (.xlsx only).
        DimensionError: appending would exceed the .xlsx row or column
            limits (.xlsx only).
        ValueError: *sheet* is given for a .csv file.
    """
    path = Path(path)
    grid = [list(row) for row in rows]
    if not grid:
        return

    if _is_csv_path(path):
        if sheet is not None:
            raise ValueError("sheet is not meaningful for a .csv file")
        if not path.exists():
            raise FileNotFoundError(f"no file at {path}")
        _append_csv(path, grid)
        return

    _check_cell_values(grid)
    if sheet is not None:
        check_sheet_name(sheet)

    workbook = safe_load(path)
    try:
        if sheet is None:
            worksheet = workbook.active
        elif sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
        else:
            worksheet = workbook.create_sheet(title=sheet)

        widest = max(len(row) for row in grid)
        check_dimensions(worksheet.max_row + len(grid), max(widest, worksheet.max_column))

        for row in grid:
            worksheet.append(row)
        atomic_save(workbook, path)
    finally:
        workbook.close()


def import_csv_to_xl(
    csv_path: str | Path,
    path: str | Path,
    *,
    sheet: str | None = None,
    encoding: str = "utf-8-sig",
) -> None:
    """Import a CSV file into a worksheet, as plain text.

    A one-shot copy, not a live link. Every field becomes a ``str`` cell —
    CSV has no other type, so nothing is inferred as a number, date, or
    boolean. *sheet* is replaced if it already exists, or created if it
    doesn't (same target semantics as :func:`write_sheet`). The .xlsx file
    must already exist (see :func:`create_workbook`).

    Args:
        csv_path: the source .csv file.
        path: the .xlsx file to import into.
        sheet: worksheet name, or ``None`` for the active sheet.
        encoding: text encoding to read *csv_path* with. Defaults to
            ``"utf-8-sig"``, which also transparently handles a leading
            byte-order mark (common in CSVs saved by Excel on Windows).

    Raises:
        FileNotFoundError: no file at *csv_path*, or none at *path*.
        SheetNameError: *sheet* is not a valid worksheet name.
        DimensionError: the CSV has more rows or columns than an .xlsx
            worksheet can hold.
        ValueError: *csv_path* is not a .csv path, or *path* is not a
            workbook path.
    """
    csv_path = Path(csv_path)
    path = Path(path)
    if not _is_csv_path(csv_path):
        raise ValueError(f"not a .csv path: {csv_path}")
    if path.suffix.lower() not in _WORKBOOK_SUFFIXES:
        raise ValueError(f"not a workbook path (need one of {sorted(_WORKBOOK_SUFFIXES)}): {path}")
    if not csv_path.exists():
        raise FileNotFoundError(f"no file at {csv_path}")
    with csv_path.open(newline="", encoding=encoding) as f:
        grid = [list(row) for row in csv.reader(f)]
    write_sheet(path, grid, sheet)


def export_xl_to_csv(
    path: str | Path,
    csv_path: str | Path,
    *,
    sheet: str | None = None,
    encoding: str = "utf-8",
) -> None:
    """Export a worksheet to a new CSV file, as plain text.

    A one-shot copy, not a live link. Every value is stringified with
    ``str()`` (``None`` becomes an empty field) — CSV has no concept of type,
    so this is a lossy conversion; re-importing the result reproduces the
    text, not the original types. Written atomically, the same way workbook
    writes are (see `Safe writes` in the README).

    Args:
        path: the .xlsx file to export from.
        csv_path: the .csv file to create. Refused if something is already
            there — export never overwrites a file it didn't create.
        sheet: worksheet name, or ``None`` for the active sheet.
        encoding: text encoding to write *csv_path* with.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: *path* is not a readable .xlsx.
        SheetNotFoundError: *sheet* names a worksheet that does not exist.
        FileExistsError: something is already at *csv_path*.
        ValueError: *path* is not a workbook path, or *csv_path* is not a
            .csv path.
    """
    path = Path(path)
    csv_path = Path(csv_path)
    if path.suffix.lower() not in _WORKBOOK_SUFFIXES:
        raise ValueError(f"not a workbook path (need one of {sorted(_WORKBOOK_SUFFIXES)}): {path}")
    if not _is_csv_path(csv_path):
        raise ValueError(f"not a .csv path: {csv_path}")
    if csv_path.exists():
        raise FileExistsError(f"{csv_path} already exists")
    rows = read_sheet(path, sheet)
    _write_csv_atomic(csv_path, rows, encoding=encoding)


def list_sheets(path: str | Path) -> list[str]:
    """Return the worksheet names in *path*, in order.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file is not a readable .xlsx.
    """
    workbook = safe_load(path)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def list_tables(path: str | Path) -> list[str]:
    """Return the names of every named table in the workbook (see ``Table.create``).

    If the reserved schema sheet is missing entirely, it's automatically
    rebuilt by scanning the workbook for table markers first — see
    :class:`~pyhandlexl.errors.SchemaRebuiltWarning`. This can make listing
    tables write to the file, so the scan isn't repeated on the next call.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file is not a readable .xlsx.
    """
    workbook = safe_load(path)
    try:
        schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
        entries = mt.load_schema(workbook)
        if not schema_existed and entries:
            mt.save_schema(workbook, entries)
            atomic_save(workbook, path)
        return list(entries.keys())
    finally:
        workbook.close()


def delete_table(path: str | Path, name: str) -> None:
    """Remove the named table called *name*.

    The space it occupied is left empty — other tables on the sheet are not
    shifted to close the gap.

    Raises:
        FileNotFoundError: no file at *path*.
        TableNotFoundError: no table named *name*, or its marker cannot be found.
    """
    workbook = safe_load(path)
    try:
        entries = mt.load_schema(workbook)
        entry = mt.get_entry(entries, name)
        entry = mt.verify_or_locate(workbook, entry)
        ws = workbook[entry.sheet]
        mt.clear_region(ws, entry.anchor_row, entry.anchor_col, entry.height, entry.width)
        del entries[name]
        mt.save_schema(workbook, entries)
        atomic_save(workbook, path)
    finally:
        workbook.close()


def sheet_exists(path: str | Path, name: str) -> bool:
    """Return whether *path* contains a worksheet called *name*."""
    return name in list_sheets(path)


def create_sheet(path: str | Path, name: str) -> None:
    """Add an empty worksheet called *name*.

    The file must already exist (see :func:`create_workbook`).

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNameError: *name* is not a valid worksheet name.
        ValueError: a worksheet called *name* already exists.
    """
    check_sheet_name(name)
    workbook = safe_load(path)
    try:
        if name in workbook.sheetnames:
            raise ValueError(f"sheet {name!r} already exists")
        workbook.create_sheet(title=name)
        atomic_save(workbook, path)
    finally:
        workbook.close()


def delete_sheet(path: str | Path, name: str) -> None:
    """Remove the worksheet called *name*.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNotFoundError: no worksheet called *name*.
        ValueError: *name* is the only worksheet (a workbook needs at least one).
    """
    workbook = safe_load(path)
    try:
        if name not in workbook.sheetnames:
            raise SheetNotFoundError(name)
        if len(workbook.sheetnames) == 1:
            raise ValueError("cannot delete the only sheet in the workbook")
        del workbook[name]
        atomic_save(workbook, path)
    finally:
        workbook.close()


def rename_sheet(path: str | Path, old: str, new: str) -> None:
    """Rename worksheet *old* to *new*.

    Raises:
        SheetNameError: *new* is not a valid worksheet name.
        FileNotFoundError: no file at *path*.
        SheetNotFoundError: no worksheet called *old*.
        ValueError: a different worksheet called *new* already exists.
    """
    check_sheet_name(new)
    workbook = safe_load(path)
    try:
        if old not in workbook.sheetnames:
            raise SheetNotFoundError(old)
        if new != old and new in workbook.sheetnames:
            raise ValueError(f"sheet {new!r} already exists")
        workbook[old].title = new
        atomic_save(workbook, path)
    finally:
        workbook.close()
