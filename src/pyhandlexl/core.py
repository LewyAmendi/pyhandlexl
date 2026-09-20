"""Public functions for reading and writing worksheets, preserving cell types."""

from __future__ import annotations

import csv
import os
import warnings
from collections.abc import Iterable
from itertools import zip_longest
from pathlib import Path
from secrets import token_hex
from typing import Any, Literal, cast

from openpyxl import Workbook
from openpyxl.cell.read_only import EmptyCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from pyhandlexl import _multi_table as mt
from pyhandlexl._multi_table import _to_label
from pyhandlexl._safety import (
    atomic_save,
    damage_is_invalid,
    ensure_writable,
    keep_permissions,
    open_for_reading,
    resolve_target,
    safe_delete,
    safe_load,
)
from pyhandlexl.errors import (
    InvalidFileError,
    SheetExistsError,
    SheetKindError,
    SheetNotFoundError,
)
from pyhandlexl.table import TableInfo
from pyhandlexl.validate import (
    check_cell_value,
    check_dimensions,
    check_sheet_name,
    normalize_cell_value,
    normalize_newlines,
)

Orientation = Literal["rows", "columns"]
SheetKind = mt.SheetKind

# Extensions openpyxl recognises as Excel workbooks.
_WORKBOOK_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})
_CSV_SUFFIX = ".csv"

# The csv module refuses any single field over 131,072 characters unless told otherwise;
# a CSV "has no size limit" (see the README), so lift it (the largest a C long holds).
_CSV_FIELD_LIMIT = 2**31 - 1


def _keep_text(worksheet: Worksheet, row: list[object]) -> None:
    """Store any string in the row just appended that starts with ``=`` as text.

    ``Worksheet.append`` makes a formula of it; pyhandlexl doesn't support formulas.
    """
    number = cast(Any, worksheet)._current_row  # the row append() just wrote
    for column, value in enumerate(row, start=1):
        if isinstance(value, str) and value.startswith("="):
            mt.store_text(worksheet.cell(row=number, column=column), value)


def _active(workbook: Workbook) -> Worksheet:
    """The workbook's active worksheet (``workbook.active`` may be ``None``, or a chart sheet)."""
    sheet = workbook.active
    if not isinstance(sheet, (Worksheet, ReadOnlyWorksheet)):
        raise InvalidFileError("the workbook has no active worksheet to use")
    return sheet


def _is_csv_path(path: Path) -> bool:
    return path.suffix.lower() == _CSV_SUFFIX


def _check_cell_values(grid: list[list[object]]) -> None:
    """Raise CellTypeError on the first value that is not a type Excel can store."""
    for row in grid:
        for value in row:
            check_cell_value(value)


def _as_row(row: object) -> list[object]:
    """One row of a grid as a list, refusing a bare ``str``/``bytes`` — which would be
    silently split into single characters (``["hello"]`` is one row, not five cells)."""
    if isinstance(row, (str, bytes)):
        raise TypeError(
            f"each row must be a sequence of values, not a single {type(row).__name__} "
            f"({row!r}) — iterating it would split it into individual characters; "
            "wrap it in a list"
        )
    return list(cast("Iterable[object]", row))


def _normalized(grid: list[list[object]]) -> list[list[object]]:
    """*grid* as cells hold it: numpy and pandas values made plain and every missing value
    empty (see :func:`~pyhandlexl.validate.normalize_cell_value`), and every string's line
    breaks stored the way Excel stores them (see
    :func:`~pyhandlexl.validate.normalize_newlines`)."""
    return [[normalize_newlines(normalize_cell_value(value)) for value in row] for row in grid]


def _checked_grid(rows: Iterable[Iterable[object]]) -> list[list[object]]:
    """Materialise *rows* into a grid, raising CellTypeError on any bad value."""
    grid = [_as_row(row) for row in rows]
    _check_cell_values(grid)
    return grid


def _refuse_case_clash(workbook: Workbook, name: str, *, ignoring: str | None = None) -> None:
    """Refuse a name that differs from an existing sheet's only by capitalisation.

    Excel and openpyxl treat ``Data`` and ``data`` as one name; openpyxl would
    quietly create ``data1`` instead of what was asked for, so a call that
    reports success would have written to (or named) the wrong sheet. A name
    that already exists exactly is fine — that's just the sheet itself.
    *ignoring* leaves one sheet out, for renaming a sheet's own capitalisation.
    """
    if name in workbook.sheetnames:
        return
    clash = mt.case_clash([n for n in workbook.sheetnames if n != ignoring], name)
    if clash is not None:
        raise SheetExistsError(
            f"a sheet called {clash!r} already exists — Excel treats sheet names as "
            f"case-insensitive, so {name!r} would collide with it"
        )


def _read_csv_rows(path: Path) -> list[list[str]]:
    previous = csv.field_size_limit(_CSV_FIELD_LIMIT)
    try:
        with path.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            try:
                return [list(row) for row in reader]
            except csv.Error as error:
                # e.g. a NUL byte on Python 3.10 (3.11+ reads it): name the line
                raise InvalidFileError(
                    f"{path} is not a readable CSV (line {reader.line_num}): {error}"
                ) from error
    finally:
        csv.field_size_limit(previous)


def _write_csv_atomic(
    path: Path, grid: Iterable[Iterable[object]], *, encoding: str = "utf-8"
) -> None:
    """Write *grid* to *path* as CSV, stringifying every value (``None`` becomes an
    empty field), replacing any existing content atomically."""
    path = resolve_target(path)
    ensure_writable(path)
    tmp = path.parent / f".{path.stem}.{token_hex(6)}.tmp.csv"
    try:
        with tmp.open("w", newline="", encoding=encoding) as f:
            writer = csv.writer(f)
            for row in grid:
                writer.writerow("" if value is None else str(value) for value in row)
        keep_permissions(path, tmp)
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

    Only ``.xlsx`` files can be created: a workbook saved under ``.xlsm``,
    ``.xltx``, ``.csv``, ``.xls`` (or anything else) would hold ``.xlsx`` content
    Excel refuses to open under that name.

    Raises:
        FileExistsError: something is already at *path*.
        FileNotFoundError: the directory *path* would go in does not exist.
        SheetNameError: *sheet* is not a valid worksheet name.
        ValueError: *path* does not end in ``.xlsx``.
    """
    check_sheet_name(sheet)
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        shown = path.suffix or "a file with no extension"
        raise ValueError(f"create_workbook only makes .xlsx files, not {shown!r}: {path}")
    if path.exists():
        raise FileExistsError(f"{path} already exists")
    if not path.parent.is_dir():
        raise FileNotFoundError(f"the directory {path.parent} does not exist")
    workbook = Workbook()
    _active(workbook).title = sheet
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
        FileReadOnlyError: the file is read-only, so nothing is written (a kind of
            ``FileLockedError``).
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
    orientation: Orientation = "rows",
) -> list[list[object]]:
    """Read a worksheet — or a .csv file — as a list of rows (or columns).

    For an **.xlsx** file, values come back as ``str``, ``int``, ``float``,
    ``bool``, ``datetime``, ``date``, ``time``, or ``timedelta``, with an
    empty cell as ``None``; trailing ``None`` values are trimmed from each
    row (or column, with ``orientation="columns"``), so a fully empty one
    becomes ``[]``. For a **.csv** file every value is a plain ``str``
    instead, exactly as written — CSV has no other type to preserve, and
    rows are read exactly as they are, with no trimming.

    Args:
        path: the .xlsx or .csv file.
        sheet: worksheet name, or ``None`` for the active sheet. Must be
            ``None`` for a .csv file — it has no sheets.
        pad: if true, right-pad every row (or column) with ``None`` (``""``
            for a .csv file) to the length of the longest one, making the
            result rectangular.
        orientation: ``"rows"`` (default) returns each worksheet row as an
            inner list; ``"columns"`` returns each worksheet column as an
            inner list instead — the transpose of ``"rows"``.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file is not a readable .xlsx.
        SheetNotFoundError: *sheet* names a worksheet that does not exist.
        ValueError: *orientation* is not ``"rows"`` or ``"columns"``, or
            *sheet* is given for a .csv file.
    """
    if orientation not in ("rows", "columns"):
        raise ValueError(f"orientation must be 'rows' or 'columns', got {orientation!r}")

    path = Path(path)
    if _is_csv_path(path):
        if sheet is not None:
            raise ValueError("sheet is not meaningful for a .csv file")
        if not path.exists():
            raise FileNotFoundError(f"no file at {path}")
        rows: list[list[object]] = [list(r) for r in _read_csv_rows(path)]
        fill: object = ""
    else:
        workbook = open_for_reading(path)
        try:
            if sheet is None:
                worksheet = _active(workbook)
            elif sheet in workbook.sheetnames:
                worksheet = workbook[sheet]
            else:
                raise SheetNotFoundError(f"no worksheet named {sheet!r}")

            rows = []
            formulas: list[str] = []
            last = 0  # how many rows hold a cell at all, however empty
            with damage_is_invalid(path):
                for number, cells in enumerate(mt.all_rows(worksheet), start=1):
                    row: list[object] = [cell.value for cell in cells]
                    for position, cell in enumerate(cells, start=1):
                        if cell.data_type == "f":
                            formulas.append(f"{get_column_letter(position)}{number}")
                    if any(type(cell) is not EmptyCell for cell in cells):
                        last = number
                    while row and row[-1] is None:
                        row.pop()
                    rows.append(row)
            # A read-only sheet also reports the rows a file lists but gives no cells (what a
            # deleted table leaves behind); a loaded sheet ends at its last cell. Same here.
            del rows[last:]
            title = worksheet.title
        finally:
            workbook.close()
        fill = None
        if formulas:
            warnings.warn(
                mt.formula_warning(f"sheet {title!r}", formulas),
                stacklevel=2 + mt.EXTRA_FRAMES.get(),
            )

    if orientation == "columns":
        rows = [list(column) for column in zip_longest(*rows, fillvalue=fill)]
        if fill is None:
            for column in rows:
                while column and column[-1] is None:
                    column.pop()

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
        SheetKindError: *sheet* already holds table data, or is the reserved
            schema sheet (.xlsx only — a sheet has no such distinction in a
            .csv file). See :class:`~pyhandlexl.errors.SheetKindError`.
        CellTypeError: a value is not a type Excel can store (.xlsx only).
        DimensionError: the data exceeds the .xlsx row or column limits
            (.xlsx only — a .csv file has no size limit).
        ValueError: *orientation* is not ``"rows"`` or ``"columns"``, *sheet*
            is given for a .csv file, or *sheet* differs from an existing
            worksheet's name only by capitalisation (Excel treats those as
            the same name).
    """
    if orientation not in ("rows", "columns"):
        raise ValueError(f"orientation must be 'rows' or 'columns', got {orientation!r}")

    path = Path(path)
    grid = [_as_row(row) for row in rows]
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
    grid = _normalized(grid)
    check_dimensions(len(grid), max((len(row) for row in grid), default=0))

    if sheet is not None:
        check_sheet_name(sheet)

    workbook = safe_load(path)
    try:
        name = sheet if sheet is not None else _active(workbook).title
        if mt.is_schema_name(name):
            raise mt.reserved_error(name, "cannot be written to")
        _refuse_case_clash(workbook, name)
        if name in workbook.sheetnames:
            schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
            entries = mt.load_schema(workbook)
            if mt.sheet_kind(workbook, entries, name) == "table":
                raise SheetKindError(
                    f"sheet {name!r} holds table data — write_sheet() would corrupt it; "
                    "use Table.write()/Table.create(), or clear_all_sheet_data() first"
                )
            if not schema_existed and entries:  # keep a rebuilt schema, as the reads do
                mt.save_schema(workbook, entries)
            index = workbook.sheetnames.index(name)
            workbook.remove(workbook[name])
            worksheet = workbook.create_sheet(title=name, index=index)
        else:
            worksheet = workbook.create_sheet(title=name)

        for row in grid:
            worksheet.append(row)
            _keep_text(worksheet, row)
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
        SheetKindError: *sheet* already holds table data, or its name is the
            reserved schema sheet's, in any capitalisation (.xlsx only).
        CellTypeError: a value is not a type Excel can store (.xlsx only).
        DimensionError: appending would exceed the .xlsx row or column
            limits (.xlsx only).
        ValueError: *sheet* is given for a .csv file, or differs from an
            existing worksheet's name only by capitalisation.
    """
    path = Path(path)
    grid = [_as_row(row) for row in rows]
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
    grid = _normalized(grid)
    if sheet is not None:
        check_sheet_name(sheet)

    workbook = safe_load(path)
    try:
        if sheet is not None:
            if mt.is_schema_name(sheet):  # checked before anything is created under that name
                raise mt.reserved_error(sheet, "cannot be written to")
            _refuse_case_clash(workbook, sheet)
        if sheet is None:
            worksheet = _active(workbook)
        elif sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
        else:
            worksheet = workbook.create_sheet(title=sheet)

        if mt.is_schema_name(worksheet.title):
            raise mt.reserved_error(worksheet.title, "cannot be written to")
        schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
        entries = mt.load_schema(workbook)
        if mt.sheet_kind(workbook, entries, worksheet.title) == "table":
            raise SheetKindError(
                f"sheet {worksheet.title!r} holds table data — append_rows() would corrupt "
                "it; use Table.write()/Table.create(), or clear_all_sheet_data() first"
            )
        if not schema_existed and entries:  # keep a rebuilt schema, as the reads do
            mt.save_schema(workbook, entries)

        widest = max(len(row) for row in grid)
        check_dimensions(worksheet.max_row + len(grid), max(widest, worksheet.max_column))

        for row in grid:
            worksheet.append(row)
            _keep_text(worksheet, row)
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
    token = mt.EXTRA_FRAMES.set(1)  # this function sits between the user and write_sheet
    try:
        write_sheet(path, grid, sheet)
    finally:
        mt.EXTRA_FRAMES.reset(token)


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
    token = mt.EXTRA_FRAMES.set(1)  # this function sits between the user and read_sheet
    try:
        rows = read_sheet(path, sheet)
    finally:
        mt.EXTRA_FRAMES.reset(token)
    _write_csv_atomic(csv_path, rows, encoding=encoding)


def list_sheets(path: str | Path) -> list[str]:
    """Return the worksheet names in *path*, in order.

    The reserved ``_pyhandlexl_tables`` schema sheet is never included — it
    isn't a sheet a caller created or can write to.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file is not a readable .xlsx.
    """
    workbook = open_for_reading(path)  # the names are in workbook.xml: no sheet is parsed
    try:
        return [name for name in workbook.sheetnames if name != mt.SCHEMA_SHEET]
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
    with mt.cheap_schema(path) as cheap:
        if cheap is not None:  # the schema sheet is all it takes
            return list(cheap[1])

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


def sheet_kind(path: str | Path, sheet: str) -> SheetKind:
    """Whether *sheet* currently holds table data, grid data, or neither.

    A sheet holds either named tables or plain grid data, never both — the
    first successful write claims it: ``"table"`` if ``Table.create`` has
    placed one there, ``"grid"`` if ``write_sheet``/``append_rows`` has, or
    ``"empty"`` if neither has (a fresh sheet, or one just wiped by
    :func:`clear_all_sheet_data`). Writing the other kind to a sheet that
    isn't ``"empty"`` raises :class:`~pyhandlexl.errors.SheetKindError`.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNotFoundError: no worksheet called *sheet*.
        SheetKindError: *sheet* is the reserved schema sheet (or a lookalike
            of its name) — it has no
            meaningful kind of its own.
    """
    with mt.cheap_schema(path) as cheap:
        if cheap is not None:
            workbook, entries = cheap
            if mt.is_schema_name(sheet):
                raise mt.reserved_error(sheet, "has no meaningful kind")
            if sheet not in workbook.sheetnames:
                raise SheetNotFoundError(f"no worksheet named {sheet!r}")
            return mt.sheet_kind(workbook, entries, sheet)

    workbook = safe_load(path)
    try:
        if mt.is_schema_name(sheet):
            raise mt.reserved_error(sheet, "has no meaningful kind")
        if sheet not in workbook.sheetnames:
            raise SheetNotFoundError(f"no worksheet named {sheet!r}")
        schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
        entries = mt.load_schema(workbook)
        kind = mt.sheet_kind(workbook, entries, sheet)
        if not schema_existed and entries:  # persist a rebuild, like the other reads do
            mt.save_schema(workbook, entries)
            atomic_save(workbook, path)
        return kind
    finally:
        workbook.close()


def clear_all_sheet_data(path: str | Path, sheet: str) -> None:
    """Wipe *sheet* back to a blank, unclaimed worksheet.

    Every cell's value and style is removed, and any tables that lived
    there are forgotten from the schema — afterwards :func:`sheet_kind`
    reports ``"empty"`` again, so the sheet can be freely claimed by
    ``write_sheet``/``append_rows`` (grid data) or ``Table.create`` (table
    data), whichever writes to it first. Its position among the workbook's
    other sheets is unchanged.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNotFoundError: no worksheet called *sheet*.
        SheetKindError: *sheet* is the reserved schema sheet (or a lookalike
            of its name).
    """
    workbook = safe_load(path)
    try:
        if mt.is_schema_name(sheet):
            raise mt.reserved_error(sheet, "cannot be cleared")
        if sheet not in workbook.sheetnames:
            raise SheetNotFoundError(f"no worksheet named {sheet!r}")

        schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
        entries = mt.load_schema(workbook)
        remaining = mt.drop_sheet_from_entries(entries, sheet)
        on_disk = entries if schema_existed else {}
        if remaining != on_disk:
            mt.save_schema(workbook, remaining)

        index = workbook.sheetnames.index(sheet)
        workbook.remove(workbook[sheet])
        workbook.create_sheet(title=sheet, index=index)
        atomic_save(workbook, path)
    finally:
        workbook.close()


def _table_info(ws: Any, entry: mt.TableEntry) -> TableInfo:
    """The metadata for the table at *entry*'s (verified) position: its header row is read, to
    pair each column type with its column's name, and nothing else."""
    header_row = mt.read_region(ws, entry.anchor_row + 1, entry.anchor_col + 1, 1, entry.n_cols)
    headers = [_to_label(h) for h in header_row[0]]
    return TableInfo(
        created_at=entry.created_at,
        modified_at=entry.modified_at,
        n_rows=entry.n_rows,
        n_cols=entry.n_cols,
        sheet=entry.sheet,
        style=entry.style,
        column_types=dict(zip(headers, entry.column_types, strict=True)),
    )


def table_info(path: str | Path, name: str) -> TableInfo:
    """Metadata for the named table, without reading its actual row data.

    Cheaper than ``Table.read(path, name).info`` when all you need is size,
    dates, sheet, style, or column types — this reads the table's header
    row (to pair each :class:`~pyhandlexl.column_type.ColumnType` with its
    column name), but never its row labels or data.

    ``created_at``/``modified_at`` are ``None`` if the table's schema entry
    was recovered by an automatic rebuild — see
    :class:`~pyhandlexl.errors.SchemaRebuiltWarning`.

    Raises:
        FileNotFoundError: no file at *path*.
        TableNotFoundError: no such table exists, or its marker cannot be
            found on its recorded sheet.
    """
    with mt.cheap_schema(path) as cheap:
        if cheap is not None:
            workbook, entries = cheap
            entry = mt.get_entry(entries, name)
            if entry.sheet in workbook.sheetnames:
                ws = workbook[entry.sheet]
                if mt.marker_matches(ws, entry):
                    return _table_info(ws, entry)

    workbook = safe_load(path)
    try:
        schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
        entries = mt.load_schema(workbook)
        entry = mt.get_entry(entries, name)
        located = mt.verify_or_locate(workbook, entry)
        healed = located != entry

        info = _table_info(workbook[located.sheet], located)

        if healed:
            entries[name] = located
        if healed or not schema_existed:
            mt.save_schema(workbook, entries)
            atomic_save(workbook, path)

        return info
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
        SheetKindError: *name* is the reserved schema sheet's name in any
            capitalisation (``_PYHANDLEXL_TABLES`` would push the real one off
            its exact name, since Excel ignores case).
        SheetExistsError: a worksheet called *name* already exists — including one that
            differs only by capitalisation, which Excel treats as the same name.
    """
    check_sheet_name(name)
    workbook = safe_load(path)
    try:
        if mt.is_schema_name(name):
            raise mt.reserved_error(name, "cannot be created")
        if name in workbook.sheetnames:
            raise SheetExistsError(f"sheet {name!r} already exists")
        _refuse_case_clash(workbook, name)
        workbook.create_sheet(title=name)
        atomic_save(workbook, path)
    finally:
        workbook.close()


def delete_sheet(path: str | Path, name: str) -> None:
    """Remove the worksheet called *name*.

    Any tables that lived on *name* are forgotten from the schema too — not
    just left behind as stale, unreachable entries.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNotFoundError: no worksheet called *name*.
        SheetKindError: *name* is the reserved schema sheet — deleting it
            directly is refused; it's only ever meant to disappear as a
            side effect of deleting every table that lives on it.
        ValueError: *name* is the only worksheet (a workbook needs at least one).
    """
    workbook = safe_load(path)
    try:
        if mt.is_schema_name(name):
            raise mt.reserved_error(name, "cannot be deleted")
        if name not in workbook.sheetnames:
            raise SheetNotFoundError(f"no worksheet named {name!r}")
        if len(workbook.sheetnames) == 1:
            raise ValueError("cannot delete the only sheet in the workbook")
        del workbook[name]

        schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
        entries = mt.load_schema(workbook)
        remaining = mt.drop_sheet_from_entries(entries, name)
        on_disk = entries if schema_existed else {}
        if remaining != on_disk:
            mt.save_schema(workbook, remaining)

        atomic_save(workbook, path)
    finally:
        workbook.close()


def rename_sheet(path: str | Path, old: str, new: str) -> None:
    """Rename worksheet *old* to *new*.

    Any tables that live on *old* are updated to point at *new* — they stay
    readable under their same names, just on the renamed sheet.

    Raises:
        SheetNameError: *new* is not a valid worksheet name.
        FileNotFoundError: no file at *path*.
        SheetNotFoundError: no worksheet called *old*.
        SheetKindError: *old* is the reserved schema sheet (or a lookalike of
            its name) — renaming it
            away would orphan it under a name pyhandlexl no longer
            recognizes, same as deleting it — or *new* is the reserved
            name, which would collide with (or masquerade as) it.
        SheetExistsError: a different worksheet called *new* already exists, including
            one that differs only by capitalisation. Changing just *old*'s own
            capitalisation (``log`` -> ``LOG``) is allowed.
    """
    check_sheet_name(new)
    workbook = safe_load(path)
    try:
        for reserved in (old, new):
            if mt.is_schema_name(reserved):
                raise mt.reserved_error(reserved, "cannot be renamed to or from")
        if old not in workbook.sheetnames:
            raise SheetNotFoundError(f"no worksheet named {old!r}")
        if new != old and new in workbook.sheetnames:
            raise SheetExistsError(f"sheet {new!r} already exists")
        _refuse_case_clash(workbook, new, ignoring=old)
        worksheet = workbook[old]
        if new != old and new.lower() == old.lower():
            # A change of capitalisation only. openpyxl would take "data" -> "DATA" for a
            # duplicate of itself and name the sheet "DATA1", so go via a throwaway name.
            worksheet.title = f"_rename_{token_hex(6)}"
        worksheet.title = new

        if new != old:
            schema_existed = mt.SCHEMA_SHEET in workbook.sheetnames
            entries = mt.load_schema(workbook)
            updated = mt.rename_sheet_in_entries(entries, old, new)
            on_disk = entries if schema_existed else {}
            if updated != on_disk:
                mt.save_schema(workbook, updated)

        atomic_save(workbook, path)
    finally:
        workbook.close()
