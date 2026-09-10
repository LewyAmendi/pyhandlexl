"""Public functions for reading and writing worksheets, preserving cell types."""

from __future__ import annotations

from collections.abc import Iterable
from itertools import zip_longest
from pathlib import Path
from typing import Literal

from openpyxl import Workbook

from pyhandlexl._safety import atomic_save, safe_delete, safe_load
from pyhandlexl.errors import SheetNotFoundError
from pyhandlexl.validate import check_cell_value, check_dimensions, check_sheet_name

Orientation = Literal["rows", "columns"]

# Extensions openpyxl recognises as Excel workbooks.
_WORKBOOK_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})


def _checked_grid(rows: Iterable[Iterable[object]]) -> list[list[object]]:
    """Materialise *rows* into a grid, raising CellTypeError on any bad value."""
    grid = [list(row) for row in rows]
    for row in grid:
        for value in row:
            check_cell_value(value)
    return grid


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
    """Read a worksheet as a list of rows, keeping each value's type.

    Values come back as ``str``, ``int``, ``float``, ``bool``, ``datetime``,
    ``date``, ``time``, or ``timedelta``. An empty cell is ``None``. Trailing
    ``None`` values are trimmed from each row, so a fully empty row becomes
    ``[]``.

    Args:
        path: the .xlsx file.
        sheet: worksheet name, or ``None`` for the active sheet.
        pad: if true, right-pad every row with ``None`` to the length of the
            longest row, making the result rectangular.

    Raises:
        FileNotFoundError: no file at *path*.
        InvalidFileError: the file is not a readable .xlsx.
        SheetNotFoundError: *sheet* names a worksheet that does not exist.
    """
    workbook = safe_load(path)
    try:
        if sheet is None:
            worksheet = workbook.active
        elif sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
        else:
            raise SheetNotFoundError(sheet)

        rows: list[list[object]] = []
        for raw_row in worksheet.iter_rows(values_only=True):
            row = list(raw_row)
            while row and row[-1] is None:
                row.pop()
            rows.append(row)
    finally:
        workbook.close()

    if pad and rows:
        width = max(len(row) for row in rows)
        rows = [row + [None] * (width - len(row)) for row in rows]
    return rows


def write_sheet(
    path: str | Path,
    rows: Iterable[Iterable[object]],
    sheet: str | None = None,
    *,
    orientation: Orientation = "rows",
) -> None:
    """Replace a worksheet's contents with *rows*.

    Other worksheets in the file are left untouched. *sheet* is added if it
    does not exist. The file must already exist (see :func:`create_workbook`).

    Values are written with their type preserved — no conversion. ``None``
    leaves the cell empty. Every value must be a type Excel can store (see
    :func:`pyhandlexl.check_cell_value`).

    Args:
        path: the .xlsx file.
        rows: an iterable of iterables of cell values.
        sheet: worksheet name, or ``None`` for the active sheet.
        orientation: ``"rows"`` writes each inner iterable as a row;
            ``"columns"`` writes each inner iterable down a column.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNameError: *sheet* is not a valid worksheet name.
        CellTypeError: a value is not a type Excel can store.
        DimensionError: the data exceeds the .xlsx row or column limits.
        ValueError: *orientation* is not ``"rows"`` or ``"columns"``.
    """
    if orientation not in ("rows", "columns"):
        raise ValueError(f"orientation must be 'rows' or 'columns', got {orientation!r}")

    grid = _checked_grid(rows)
    if orientation == "columns":
        grid = [list(column) for column in zip_longest(*grid, fillvalue=None)]

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
    """Append *rows* to the end of a worksheet.

    *sheet* is added if it does not exist. An empty *rows* is a no-op. The file
    must already exist (see :func:`create_workbook`). Values follow the same rules
    as :func:`write_sheet`.

    Raises:
        FileNotFoundError: no file at *path*.
        SheetNameError: *sheet* is not a valid worksheet name.
        CellTypeError: a value is not a type Excel can store.
        DimensionError: appending would exceed the .xlsx row or column limits.
    """
    grid = _checked_grid(rows)
    if not grid:
        return

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
