"""Validation helpers: sheet names, data dimensions, cell types, file integrity."""

from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path

from openpyxl import load_workbook

from pyhandlexl.errors import CellTypeError, DimensionError, SheetNameError

# Excel's hard limits for the .xlsx format.
MAX_ROWS = 1_048_576
MAX_COLUMNS = 16_384

# Characters Excel forbids in a worksheet name.
ILLEGAL_SHEET_CHARS = frozenset(r"\/?*[]:")

# Files every valid .xlsx zip must contain.
_REQUIRED_PARTS = frozenset({"xl/workbook.xml", "[Content_Types].xml"})

# The value types Excel can store in a cell. ``bool`` is covered by ``int`` and
# ``datetime`` by ``date``; ``None`` means an empty cell.
_CELL_TYPES = (str, int, float, dt.date, dt.time, dt.timedelta)


def check_dimensions(n_rows: int, n_cols: int) -> None:
    """Raise DimensionError if a grid of this size won't fit in an .xlsx sheet."""
    if n_rows > MAX_ROWS:
        raise DimensionError(f"{n_rows} rows exceeds the .xlsx limit of {MAX_ROWS}")
    if n_cols > MAX_COLUMNS:
        raise DimensionError(f"{n_cols} columns exceeds the .xlsx limit of {MAX_COLUMNS}")


def check_cell_value(value: object) -> None:
    """Raise CellTypeError if *value* is not a type Excel can store in a cell.

    Allowed: ``str``, ``int``, ``float``, ``bool``, ``datetime``, ``date``,
    ``time``, ``timedelta``, and ``None`` (an empty cell). Timezone-aware
    datetimes are rejected — Excel has no concept of a timezone.
    """
    if value is None or isinstance(value, _CELL_TYPES):
        if isinstance(value, (dt.datetime, dt.time)) and value.tzinfo is not None:
            raise CellTypeError(
                f"{type(value).__name__} must be timezone-naive; Excel has no timezone concept"
            )
        return
    raise CellTypeError(f"{type(value).__name__} is not a type Excel can store: {value!r}")


def check_sheet_name(sheet_name: str) -> None:
    """Raise SheetNameError if *sheet_name* is not a valid Excel worksheet name."""
    if not isinstance(sheet_name, str):
        raise SheetNameError(f"sheet name must be a string, got {type(sheet_name).__name__}")

    if sheet_name == "":
        raise SheetNameError("sheet name must not be empty")

    if len(sheet_name) > 31:
        raise SheetNameError(f"sheet name is longer than 31 characters: {sheet_name!r}")

    illegal = sorted(set(sheet_name) & ILLEGAL_SHEET_CHARS)
    if illegal:
        raise SheetNameError(f"sheet name {sheet_name!r} contains illegal character(s): {illegal}")

    if sheet_name.lower() == "history":
        raise SheetNameError("'History' is reserved by Excel and cannot be used as a sheet name")


def is_valid_xlsx(path: str | Path) -> bool:
    """Return True only if *path* is a readable .xlsx workbook.

    Never raises: every failure mode returns False.
    """
    path = Path(path)
    if not path.is_file():
        return False

    try:
        with zipfile.ZipFile(path) as zf:
            if not _REQUIRED_PARTS.issubset(zf.namelist()):
                return False
            if zf.testzip() is not None:
                return False
    except (zipfile.BadZipFile, OSError):
        return False

    # openpyxl raises several unrelated exception types for a malformed
    # workbook; this function's job is to answer yes/no, not to blow up.
    try:
        load_workbook(path, read_only=True).close()
    except Exception:
        return False

    return True
