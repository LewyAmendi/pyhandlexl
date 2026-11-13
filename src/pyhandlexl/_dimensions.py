"""Internal: column widths, row heights and text fitting on an openpyxl worksheet.

Not part of the public API.

openpyxl keeps column widths as ranges (a ``<col min= max=>`` element can cover many columns),
and setting a width on a column that lies inside such a range would write an overlapping
second range, which Excel rejects. So every change here first splits the range around the
column. Row heights are one per row and need no such care.
"""

from __future__ import annotations

from copy import copy
from datetime import date, datetime, time, timedelta
from typing import Any

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension

DEFAULT_COLUMN_WIDTH = 13.0  # what openpyxl gives a column dimension that has no width of its own

MAX_WIDTH = 255.0  # Excel's limit, in characters
MAX_HEIGHT = 409.0  # Excel's limit, in points
DEFAULT_MAX_FIT = 60  # a longer line is wrapped rather than widening the column further
_PADDING = 3  # room for the bold header and the cell border
_MIN_FIT = 9.0  # never narrower than a default column


def _span(dim: Any) -> tuple[int, int]:
    lo = dim.min or column_index_from_string(dim.index)
    return lo, dim.max or lo


def _covering(ws: Any, idx: int) -> Any | None:
    for dim in ws.column_dimensions.values():
        lo, hi = _span(dim)
        if lo <= idx <= hi:
            return dim
    return None


def get_column_width(ws: Any, idx: int) -> float | None:
    """The explicit width of column *idx* (1-based), or ``None`` if it has none."""
    dim = _covering(ws, idx)
    return None if dim is None or dim.width is None else float(dim.width)


def _clone(ws: Any, dim: Any, lo: int, hi: int) -> Any:
    new: Any = ColumnDimension(
        ws,
        index=get_column_letter(lo),
        width=dim.width,
        bestFit=dim.bestFit,
        hidden=dim.hidden,
        outlineLevel=dim.outlineLevel,
        collapsed=dim.collapsed,
        min=lo,
        max=hi,
    )
    if getattr(dim, "_style", None) is not None:
        new._style = copy(dim._style)
    return new


def _isolate(ws: Any, idx: int) -> None:
    """Make column *idx* a dimension of its own, splitting any range that holds it."""
    dim = _covering(ws, idx)
    if dim is None:
        return
    lo, hi = _span(dim)
    if lo == hi:
        return
    del ws.column_dimensions[dim.index]
    if lo < idx:
        ws.column_dimensions[get_column_letter(lo)] = _clone(ws, dim, lo, idx - 1)
    ws.column_dimensions[get_column_letter(idx)] = _clone(ws, dim, idx, idx)
    if idx < hi:
        ws.column_dimensions[get_column_letter(idx + 1)] = _clone(ws, dim, idx + 1, hi)


def set_column_width(ws: Any, idx: int, width: float | None) -> None:
    """Give column *idx* (1-based) *width* characters, or restore the default for ``None``."""
    _isolate(ws, idx)
    letter = get_column_letter(idx)
    if width is None:
        dim = ws.column_dimensions.get(letter)
        if dim is not None:
            plain = not (dim.hidden or dim.outlineLevel or dim.collapsed)
            if plain and getattr(dim, "_style", None) is None:
                del ws.column_dimensions[letter]
            else:  # it carries more than a width (hidden, grouped, a style): keep that
                dim.width = DEFAULT_COLUMN_WIDTH
    else:
        ws.column_dimensions[letter].width = float(width)


def get_row_height(ws: Any, idx: int) -> float | None:
    dim = ws.row_dimensions.get(idx)
    return None if dim is None or dim.height is None else float(dim.height)


def set_row_height(ws: Any, idx: int, height: float | None) -> None:
    if height is None:
        dim = ws.row_dimensions.get(idx)
        if dim is not None:
            dim.height = None
    else:
        ws.row_dimensions[idx].height = float(height)


# ------------------------------------------------------------------ validation


def check_width(width: object) -> float:
    if isinstance(width, bool) or not isinstance(width, (int, float)):
        raise TypeError(f"width must be a number, got {type(width).__name__}: {width!r}")
    if not 0 < width <= MAX_WIDTH:
        raise ValueError(f"width must be above 0 and at most {MAX_WIDTH:g}, got {width!r}")
    return float(width)


def check_height(height: object) -> float:
    if isinstance(height, bool) or not isinstance(height, (int, float)):
        raise TypeError(f"height must be a number, got {type(height).__name__}: {height!r}")
    if not 0 < height <= MAX_HEIGHT:
        raise ValueError(f"height must be above 0 and at most {MAX_HEIGHT:g}, got {height!r}")
    return float(height)


def check_max_width(value: object) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"max_width must be an int, got {type(value).__name__}: {value!r}")
    if not 1 <= value <= MAX_WIDTH:
        raise ValueError(f"max_width must be between 1 and {MAX_WIDTH:g}, got {value}")
    return value


# --------------------------------------------------------------------- fitting


def longest_line(value: object) -> int:
    """How many characters wide *value* is when shown (its longest line)."""
    if value is None:
        return 0
    if isinstance(value, str):
        return max((len(line) for line in value.split("\n")), default=0)
    if isinstance(value, datetime):
        return 19
    if isinstance(value, date):
        return 10
    if isinstance(value, (time, timedelta)):
        return 8
    return len(str(value))


def needs_wrap(value: object, max_width: int) -> bool:
    """Whether *value* is text that should wrap: it has a line break or a line past the cap."""
    return isinstance(value, str) and ("\n" in value or len(value) > max_width)


def fit_width(longest: int, max_width: int) -> float:
    """The column width that shows a *longest*-character line, never beyond *max_width*."""
    return float(max(_MIN_FIT, min(longest, max_width) + _PADDING))
