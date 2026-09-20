"""Conversions between pyhandlexl's tables and grids and pandas / numpy.

Internal module — the public entry points are ``Table.to_dataframe``, ``Table.from_dataframe``,
``Table.to_numpy`` and ``pyhandlexl.grid.to_dataframe`` / ``from_dataframe``. pandas and numpy
stay optional: they are imported here, by name at the moment a conversion needs one, and
nowhere else — so importing pyhandlexl never imports either.
"""

from __future__ import annotations

import importlib
import warnings
from collections.abc import Iterable, Mapping
from datetime import datetime, timedelta
from typing import TYPE_CHECKING, Any

from pyhandlexl.column_type import ColumnType
from pyhandlexl.validate import normalize_cell_value

if TYPE_CHECKING:
    from pyhandlexl.table import TableData


def _library(name: str, extra: str) -> Any:
    """The optional library *name*, or an ImportError that says how to get it."""
    try:
        return importlib.import_module(name)
    except ImportError:
        raise ImportError(
            f"{name} is needed for this and is not installed — "
            f'install it with: pip install "pyhandlexl[{extra}]"'
        ) from None


# ---------------------------------------------------------------- Table -> pandas


def _hinted(pd: Any, series: Any, column_type: ColumnType) -> Any | None:
    """*series* converted to the dtype its column's type calls for, or None to leave it as
    inferred. Inference already does most of it; this settles what it can't: a column of
    only blanks (all ``None``), a boolean column with blanks (pandas' nullable
    ``boolean`` instead of ``object``), and dates that arrive as a mix of ``date`` and
    ``datetime``.
    """
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # pandas' own notes about parsing odd values
            return _convert(pd, series, column_type)
    except (TypeError, ValueError, OverflowError):
        return None  # values that don't fit their column's type (only in memory): keep as is


def _convert(pd: Any, series: Any, column_type: ColumnType) -> Any | None:
    if column_type is ColumnType.NUMBER:
        return series.astype("float64") if series.isna().all() else pd.to_numeric(series)
    if column_type is ColumnType.BOOLEAN:
        return series.astype("boolean") if series.isna().any() else series.astype(bool)
    if column_type is ColumnType.DATE:
        return pd.to_datetime(series)
    if column_type is ColumnType.DURATION:
        return pd.to_timedelta(series)
    return None


def table_to_dataframe(data: TableData, column_types: list[ColumnType]) -> Any:
    """A DataFrame: row labels as the index (the corner is its name), headers as the columns."""
    pd = _library("pandas", "pandas")
    index = pd.Index(data.row_labels, name=data.corner or None)
    frame = pd.DataFrame(data.rows, columns=data.column_headers, index=index)
    for position, column_type in enumerate(column_types):
        hinted = _hinted(pd, frame.iloc[:, position], column_type)
        if hinted is not None:
            frame.isetitem(position, hinted)
    return frame


# ---------------------------------------------------------------- pandas -> Table

_INFERRED: Mapping[str, ColumnType] = {
    "integer": ColumnType.NUMBER,
    "floating": ColumnType.NUMBER,
    "mixed-integer-float": ColumnType.NUMBER,
    "boolean": ColumnType.BOOLEAN,
    "datetime64": ColumnType.DATE,
    "datetime": ColumnType.DATE,
    "date": ColumnType.DATE,
    "timedelta64": ColumnType.DURATION,
    "timedelta": ColumnType.DURATION,
    "time": ColumnType.TIME,
    "string": ColumnType.TEXT,
}


def _labels(axis: Any, what: str, fix: str) -> list[str]:
    """The labels on one axis of a DataFrame, which must all be strings."""
    if getattr(axis, "nlevels", 1) > 1:
        raise TypeError(
            f"a table has one level of {what}s, but this DataFrame's has {axis.nlevels} — "
            f"flatten it first (for the index, df.reset_index())"
        )
    labels = list(axis)
    for label in labels:
        if not isinstance(label, str):
            raise TypeError(
                f"the DataFrame's {what}s must be strings — a table's {what}s are text — but "
                f"one is {type(label).__name__} ({label!r}). Convert them with {fix}"
            )
    return labels


def dataframe_to_table_parts(
    df: Any, infer_column_types: bool
) -> tuple[list[list[object]], list[str], str, list[str], dict[str, ColumnType]]:
    """``(rows, row_labels, corner, column_headers, inferred_types)`` to build a Table from *df*."""
    pd = _library("pandas", "pandas")
    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"expected a pandas DataFrame, got {type(df).__name__}")
    row_labels = _labels(
        df.index,
        "index label",
        "df.index = df.index.astype(str), or pick a column to be the index with "
        "df.set_index('column')",
    )
    headers = _labels(df.columns, "column name", "df.columns = df.columns.astype(str)")
    corner = df.index.name
    if corner is not None and not isinstance(corner, str):
        raise TypeError(
            f"the index name becomes the table's corner, which must be a string, "
            f"not {type(corner).__name__} ({corner!r})"
        )
    rows = df.astype(object).to_numpy().tolist()
    inferred: dict[str, ColumnType] = {}
    if infer_column_types:
        for position, header in enumerate(headers):
            kind = pd.api.types.infer_dtype(df.iloc[:, position], skipna=True)
            column_type = _INFERRED.get(kind)
            if column_type is not None:
                inferred[header] = column_type
    return rows, row_labels, corner or "", headers, inferred


# ------------------------------------------------------------------ Table -> numpy


def rows_to_numpy(rows: list[list[object]], width: int, dtype: Any) -> Any:
    """The data block as a 2-D array.

    With no *dtype* the array gets the natural one for what it holds: ``int64``, ``float64``,
    ``bool``, ``datetime64[us]`` or ``timedelta64[us]`` when every cell is of that kind (a
    blank is ``nan`` / ``NaT``), and ``object`` for anything else — text, or a mixture.
    """
    np = _library("numpy", "numpy")
    if dtype is None:
        dtype = _natural_dtype(rows)
    try:
        array = np.array(rows, dtype=dtype)
    except (TypeError, ValueError) as error:
        raise TypeError(
            f"can't convert this table to {dtype!r}: {error}. A blank cell is None, which "
            "becomes nan in a float array (and can't be an integer) — use a float or object dtype"
        ) from None
    return array.reshape(len(rows), width)


def _natural_dtype(rows: list[list[object]]) -> Any:
    kinds = {type(value) for row in rows for value in row}
    blanks = type(None) in kinds
    kinds.discard(type(None))
    if not kinds:
        return "float64"  # nothing but blanks: nan
    if kinds == {int}:
        return "float64" if blanks else "int64"
    if kinds <= {int, float}:
        return "float64"
    if kinds == {bool} and not blanks:
        return bool
    if kinds == {datetime}:
        return "datetime64[us]"
    if kinds == {timedelta}:
        return "timedelta64[us]"
    return object


# ---------------------------------------------------------------- grids <-> pandas


def grid_to_dataframe(grid: Iterable[Iterable[object]], header: bool) -> Any:
    """A DataFrame from a grid: the first row as the column names when *header*, and every row
    padded to the widest one."""
    pd = _library("pandas", "pandas")
    rows = [list(row) for row in grid]
    width = max((len(row) for row in rows), default=0)
    rows = [row + [None] * (width - len(row)) for row in rows]
    if header and rows:
        return pd.DataFrame(rows[1:], columns=rows[0])
    return pd.DataFrame(rows)


def dataframe_to_grid(df: Any, header: bool, index: bool) -> list[list[object]]:
    """A grid from a DataFrame: values as a cell holds them (blanks for ``nan``, ``NaT`` and
    ``NA``), the column names as the first row when *header*, the index as the first column
    when *index*."""
    pd = _library("pandas", "pandas")
    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"expected a pandas DataFrame, got {type(df).__name__}")
    for axis, used in ((df.columns, header), (df.index, index)):
        if used and getattr(axis, "nlevels", 1) > 1:
            raise TypeError("a MultiIndex can't be written to a grid — flatten it first")
    body = [[normalize_cell_value(v) for v in row] for row in df.astype(object).to_numpy().tolist()]
    if index:
        body = [
            [normalize_cell_value(label), *row] for label, row in zip(df.index, body, strict=True)
        ]
    if not header:
        return body
    names: list[object] = [normalize_cell_value(name) for name in df.columns]
    if index:
        names.insert(0, normalize_cell_value(df.index.name))
    return [names, *body]
