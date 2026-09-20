"""Editing helpers for the plain grids that :func:`pyhandlexl.read_sheet` returns.

Every function takes a grid (a list of rows) and returns a **new** grid — the
input is never modified — so calls compose:

    write_sheet(path, grid.delete_column(read_sheet(path), 2))

Rows and columns are addressed with 1-based, Excel-style numbers (row 1 is the
first row), matching ``Table``. Ragged rows are handled: column operations pad
short rows with ``None`` as needed.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from pyhandlexl import _frames

Grid = list[list[object]]


def _values(values: Iterable[object]) -> list[object]:
    """A row or column's values as a list — refusing a bare ``str``/``bytes``, which
    iterating would silently split into single characters."""
    if isinstance(values, (str, bytes)):
        raise TypeError(
            f"values must be a sequence, not a single {type(values).__name__} ({values!r}) — "
            "iterating it would split it into individual characters"
        )
    return list(values)


def _copy(grid: Iterable[Iterable[object]]) -> Grid:
    return [_values(row) for row in grid]


def _width(grid: Grid) -> int:
    return max((len(row) for row in grid), default=0)


def _check_row(grid: Grid, row: int, *, extra: int = 0) -> None:
    if not isinstance(row, int):
        raise TypeError(f"row must be int, got {type(row).__name__}: {row!r}")
    upper = len(grid) + extra
    if not 1 <= row <= upper:
        raise IndexError(f"row {row} is out of range (1..{upper})")


def _check_column(grid: Grid, col: int, *, extra: int = 0) -> None:
    if not isinstance(col, int):
        raise TypeError(f"col must be int, got {type(col).__name__}: {col!r}")
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
    """Return column *col* (1-based) as a list; short rows contribute ``None``."""
    grid = _copy(grid)
    _check_column(grid, col)
    return [row[col - 1] if col - 1 < len(row) else None for row in grid]


# ------------------------------------------------------------------- editing


def set_value(grid: Iterable[Iterable[object]], row: int, col: int, value: object) -> Grid:
    """Return a new grid with the cell at *row*, *col* (both 1-based) set to *value*."""
    out = _copy(grid)
    _check_row(out, row)
    _check_column(out, col)
    target = out[row - 1]
    if len(target) < col:
        target.extend([None] * (col - len(target)))
    target[col - 1] = value
    return out


def set_row(grid: Iterable[Iterable[object]], row: int, values: Iterable[object]) -> Grid:
    """Return a new grid with row *row* (1-based) replaced by *values*."""
    out = _copy(grid)
    _check_row(out, row)
    out[row - 1] = _values(values)
    return out


def set_column(grid: Iterable[Iterable[object]], col: int, values: Iterable[object]) -> Grid:
    """Return a new grid with column *col* (1-based) replaced by *values*.

    ``len(values)`` must equal the number of rows.
    """
    out = _copy(grid)
    _check_column(out, col)
    values = _values(values)
    if len(values) != len(out):
        raise ValueError(f"expected {len(out)} values, got {len(values)}")
    for row, value in zip(out, values, strict=True):
        if len(row) < col:
            row.extend([None] * (col - len(row)))
        row[col - 1] = value
    return out


def insert_row(grid: Iterable[Iterable[object]], row: int, values: Iterable[object]) -> Grid:
    """Return a new grid with *values* inserted as row *row* (1-based)."""
    out = _copy(grid)
    _check_row(out, row, extra=1)
    out.insert(row - 1, _values(values))
    return out


def insert_column(grid: Iterable[Iterable[object]], col: int, values: Iterable[object]) -> Grid:
    """Return a new grid with *values* inserted as column *col* (1-based).

    ``len(values)`` must equal the number of rows.
    """
    out = _copy(grid)
    _check_column(out, col, extra=1)
    values = _values(values)
    if len(values) != len(out):
        raise ValueError(f"expected {len(out)} values, got {len(values)}")
    for row, value in zip(out, values, strict=True):
        if len(row) < col - 1:
            row.extend([None] * (col - 1 - len(row)))
        row.insert(col - 1, value)
    return out


def append_row(grid: Iterable[Iterable[object]], values: Iterable[object]) -> Grid:
    """Return a new grid with *values* added as the last row."""
    out = _copy(grid)
    out.append(_values(values))
    return out


def append_column(grid: Iterable[Iterable[object]], values: Iterable[object]) -> Grid:
    """Return a new grid with *values* added as the last column.

    Rows are padded to equal width first, so the new value lands in the same
    column for every row. ``len(values)`` must equal the number of rows.
    """
    out = pad(grid)
    values = _values(values)
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
    """Return a new rectangular grid, every row right-padded with ``None``."""
    out = _copy(grid)
    width = _width(out)
    for row in out:
        row.extend([None] * (width - len(row)))
    return out


def transpose(grid: Iterable[Iterable[object]]) -> Grid:
    """Return the grid with rows and columns swapped.

    Ragged rows are padded with ``None`` before transposing.
    """
    out = pad(grid)
    if not out:
        return []
    return [list(column) for column in zip(*out, strict=True)]


# ----------------------------------------------------------------- display


def to_dataframe(grid: Iterable[Iterable[object]], *, header: bool = True) -> Any:
    """A pandas ``DataFrame`` from a grid.

    With *header* (the default) the first row becomes the column names; without it the
    columns are numbered from 0. Rows are padded to the widest one, and a blank is ``NaN``
    (``None`` in a text column). Needs pandas (``pip install "pyhandlexl[pandas]"``).

    For a *table* — a sheet with row labels and column headers — use
    :meth:`Table.to_dataframe <pyhandlexl.Table.to_dataframe>` instead.

    Raises:
        ImportError: pandas is not installed.
    """
    return _frames.grid_to_dataframe(grid, header)


def from_dataframe(df: Any, *, header: bool = True, index: bool = False) -> Grid:
    """A grid from a pandas ``DataFrame``, ready for ``write_sheet``.

    With *header* (the default) the column names are the first row; with *index* the index
    is the first column (and its name the first header). Values are stored as the plain
    Python values they hold, and a missing value — ``NaN``, ``NaT``, ``pd.NA`` — is a blank
    cell. (``write_sheet(path, df)`` would not do: iterating a DataFrame yields its column
    names.) Needs pandas (``pip install "pyhandlexl[pandas]"``).

    Raises:
        ImportError: pandas is not installed.
        TypeError: *df* isn't a DataFrame, or a ``MultiIndex`` would have to be written.
    """
    return _frames.dataframe_to_grid(df, header, index)


def show(
    grid: Iterable[Iterable[object]],
    *,
    rows: int | None = None,
    head: int | None = 5,
    tail: int | None = 5,
) -> None:
    """Print a grid to the console as a plain aligned block.

    Same truncation rules as ``Table.show``: by default the first 5 and last
    5 rows with a ``...`` divider between them (nothing hidden at 10 rows or
    fewer). ``rows=n`` overrides that and prints only the first *n* rows; pass
    ``head=None, tail=None`` for every row. A console convenience only — it
    doesn't touch the grid. Ragged rows are padded with ``""`` for display.
    """
    for param_name, value in (("rows", rows), ("head", head), ("tail", tail)):
        if value is None:
            continue
        if not isinstance(value, int):
            raise TypeError(f"{param_name} must be int, got {type(value).__name__}: {value!r}")
        if value < 0:
            raise ValueError(f"{param_name} must not be negative, got {value}")

    data = _copy(grid)
    if not data:
        print("(empty grid)")
        return

    divider_after: int | None = None
    if rows is not None:
        data = data[:rows]
    elif head is not None or tail is not None:
        h, t = head or 0, tail or 0
        if h + t < len(data):
            data = data[:h] + data[len(data) - t :]
            divider_after = h

    if not data:
        return

    width = _width(data)
    display_rows: list[list[str]] = []
    for i, row in enumerate(data):
        padded = row + [None] * (width - len(row))
        display_rows.append(["" if v is None else str(v) for v in padded])
        if divider_after is not None and i == divider_after - 1 and divider_after < len(data):
            display_rows.append(["..."] * width)

    col_widths = [max(len(r[c]) for r in display_rows) for c in range(width)]
    for line in display_rows:
        print("  ".join(cell.ljust(w) for cell, w in zip(line, col_widths, strict=True)))
