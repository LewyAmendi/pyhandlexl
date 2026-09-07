"""Editing helpers for the plain grids that :func:`pyhandlexl.read_sheet` returns.

Every function takes a grid (a list of rows) and returns a **new** grid — the
input is never modified — so calls compose:

    write_sheet(path, grid.delete_column(read_sheet(path), 2))

Rows and columns are addressed with 1-based, Excel-style numbers (row 1 is the
first row), matching ``Table``. Ragged rows are handled: column operations pad
short rows with ``""`` as needed.
"""

from __future__ import annotations

from collections.abc import Iterable

Grid = list[list[object]]


def _copy(grid: Iterable[Iterable[object]]) -> Grid:
    return [list(row) for row in grid]


def _width(grid: Grid) -> int:
    return max((len(row) for row in grid), default=0)


def _check_row(grid: Grid, row: int, *, extra: int = 0) -> None:
    upper = len(grid) + extra
    if not 1 <= row <= upper:
        raise IndexError(f"row {row} is out of range (1..{upper})")


def _check_column(grid: Grid, col: int, *, extra: int = 0) -> None:
    upper = _width(grid) + extra
    if not 1 <= col <= upper:
        raise IndexError(f"column {col} is out of range (1..{upper})")


# ------------------------------------------------------------------- reading


def get_row(grid: Iterable[Iterable[object]], row: int) -> list[object]:
    """Return row *row* (1-based) as a list."""
    grid = _copy(grid)
    _check_row(grid, row)
    return grid[row - 1]


def get_column(grid: Iterable[Iterable[object]], col: int) -> list[object]:
    """Return column *col* (1-based) as a list; short rows contribute ``""``."""
    grid = _copy(grid)
    _check_column(grid, col)
    return [row[col - 1] if col - 1 < len(row) else "" for row in grid]


# ------------------------------------------------------------------- editing


def set_value(grid: Iterable[Iterable[object]], row: int, col: int, value: object) -> Grid:
    """Return a new grid with the cell at *row*, *col* (both 1-based) set to *value*."""
    out = _copy(grid)
    _check_row(out, row)
    _check_column(out, col)
    target = out[row - 1]
    if len(target) < col:
        target.extend([""] * (col - len(target)))
    target[col - 1] = value
    return out


def set_row(grid: Iterable[Iterable[object]], row: int, values: Iterable[object]) -> Grid:
    """Return a new grid with row *row* (1-based) replaced by *values*."""
    out = _copy(grid)
    _check_row(out, row)
    out[row - 1] = list(values)
    return out


def set_column(grid: Iterable[Iterable[object]], col: int, values: Iterable[object]) -> Grid:
    """Return a new grid with column *col* (1-based) replaced by *values*.

    ``len(values)`` must equal the number of rows.
    """
    out = _copy(grid)
    _check_column(out, col)
    values = list(values)
    if len(values) != len(out):
        raise ValueError(f"expected {len(out)} values, got {len(values)}")
    for row, value in zip(out, values, strict=True):
        if len(row) < col:
            row.extend([""] * (col - len(row)))
        row[col - 1] = value
    return out


def insert_row(grid: Iterable[Iterable[object]], row: int, values: Iterable[object]) -> Grid:
    """Return a new grid with *values* inserted as row *row* (1-based)."""
    out = _copy(grid)
    _check_row(out, row, extra=1)
    out.insert(row - 1, list(values))
    return out


def insert_column(grid: Iterable[Iterable[object]], col: int, values: Iterable[object]) -> Grid:
    """Return a new grid with *values* inserted as column *col* (1-based).

    ``len(values)`` must equal the number of rows.
    """
    out = _copy(grid)
    _check_column(out, col, extra=1)
    values = list(values)
    if len(values) != len(out):
        raise ValueError(f"expected {len(out)} values, got {len(values)}")
    for row, value in zip(out, values, strict=True):
        if len(row) < col - 1:
            row.extend([""] * (col - 1 - len(row)))
        row.insert(col - 1, value)
    return out


def append_row(grid: Iterable[Iterable[object]], values: Iterable[object]) -> Grid:
    """Return a new grid with *values* added as the last row."""
    out = _copy(grid)
    out.append(list(values))
    return out


def append_column(grid: Iterable[Iterable[object]], values: Iterable[object]) -> Grid:
    """Return a new grid with *values* added as the last column.

    Rows are padded to equal width first, so the new value lands in the same
    column for every row. ``len(values)`` must equal the number of rows.
    """
    out = pad(grid)
    values = list(values)
    if len(values) != len(out):
        raise ValueError(f"expected {len(out)} values, got {len(values)}")
    for row, value in zip(out, values, strict=True):
        row.append(value)
    return out


def delete_row(grid: Iterable[Iterable[object]], row: int) -> Grid:
    """Return a new grid without row *row* (1-based)."""
    out = _copy(grid)
    _check_row(out, row)
    del out[row - 1]
    return out


def delete_column(grid: Iterable[Iterable[object]], col: int) -> Grid:
    """Return a new grid without column *col* (1-based)."""
    out = _copy(grid)
    _check_column(out, col)
    for row in out:
        if col - 1 < len(row):
            del row[col - 1]
    return out


# ------------------------------------------------------------------- shape


def pad(grid: Iterable[Iterable[object]]) -> Grid:
    """Return a new rectangular grid, every row right-padded with ``""``."""
    out = _copy(grid)
    width = _width(out)
    for row in out:
        row.extend([""] * (width - len(row)))
    return out


def transpose(grid: Iterable[Iterable[object]]) -> Grid:
    """Return the grid with rows and columns swapped.

    Ragged rows are padded with ``""`` before transposing.
    """
    out = pad(grid)
    if not out:
        return []
    return [list(column) for column in zip(*out, strict=True)]
