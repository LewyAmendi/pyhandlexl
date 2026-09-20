"""Validation helpers: sheet names, data dimensions, cell types, file integrity."""

from __future__ import annotations

import datetime as dt
import math
import re
import sys
import zipfile
from pathlib import Path
from typing import Any, cast

from openpyxl import load_workbook

from pyhandlexl.errors import CellTypeError, DimensionError, SheetNameError

# Excel's hard limits for the .xlsx format.
MAX_ROWS = 1_048_576
MAX_COLUMNS = 16_384

# The most characters one cell can hold.
MAX_CELL_CHARS = 32_767

# Excel stores dates as a day count from 1900-01-01; nothing earlier fits, and
# milliseconds are the finest it keeps (openpyxl rounds to them), so anything that
# would round up into year 10000 is out of range too.
_EXCEL_MIN = dt.datetime(1900, 1, 1)
_EXCEL_MAX = dt.datetime(9999, 12, 31, 23, 59, 59, 999000)

# Characters that can't appear in the XML an .xlsx is made of (XML 1.0): most control
# characters, lone surrogates, and the two non-characters U+FFFE and U+FFFF. A single one
# in a cell makes the *whole workbook* unreadable, not just that cell.
_XML_ILLEGAL = re.compile(r"[^\x09\x0a\x0d\x20-\ud7ff\ue000-\ufffd\U00010000-\U0010ffff]")

# Characters Excel forbids in a worksheet name.
ILLEGAL_SHEET_CHARS = frozenset(r"\/?*[]:")

# Files every valid .xlsx zip must contain.
_REQUIRED_PARTS = frozenset({"xl/workbook.xml", "[Content_Types].xml"})

# The value types Excel can store in a cell. ``bool`` is covered by ``int`` and
# ``datetime`` by ``date``; ``None`` means an empty cell.
_CELL_TYPES = (str, int, float, dt.date, dt.time, dt.timedelta)


# The types that need no conversion — by far the common case, so it is one set lookup.
_PLAIN = frozenset({str, int, bool, float, dt.datetime, dt.date, dt.time, dt.timedelta, type(None)})


def normalize_cell_value(value: object) -> object:
    """*value* as a cell holds it: what numpy and pandas produce, made plain.

    * a numpy scalar becomes the Python value it holds (``np.int64(3)`` → ``3``,
      ``np.bool_`` → ``bool``, ``np.datetime64`` → ``datetime``, ``np.timedelta64`` →
      ``timedelta``), and a pandas ``Timestamp`` or ``Timedelta`` the plain ``datetime`` or
      ``timedelta``;
    * **every kind of missing value becomes ``None``, an empty cell**: ``nan`` (float or
      numpy), ``pd.NaT``, ``pd.NA`` and ``np.datetime64('NaT')``;
    * anything else — including a value neither library made — is returned unchanged.

    numpy and pandas are never imported: the types are recognised by where they come from,
    so this costs nothing when they aren't installed, and a plain value is one set lookup.
    Infinities are left alone (Excel can't store them; :func:`check_cell_value` refuses them).
    """
    kind = type(value)
    if kind in _PLAIN:
        if kind is float and value != value:
            return None
        return value
    if kind.__module__.partition(".")[0] not in ("numpy", "pandas"):
        return value
    return _from_analysis_library(cast(Any, value), kind)


def _from_analysis_library(value: Any, kind: type) -> object:
    name = kind.__name__
    if name in ("NAType", "NaTType"):  # pd.NA, pd.NaT
        return None
    if hasattr(value, "to_pydatetime"):  # pd.Timestamp
        return value.to_pydatetime(warn=False)
    if hasattr(value, "to_pytimedelta"):  # pd.Timedelta
        return value.to_pytimedelta()
    if name in ("datetime64", "timedelta64"):
        if value != value:  # NaT
            return None
        unit = "datetime64[us]" if name == "datetime64" else "timedelta64[us]"
        plain = value.astype(unit).item()
        if not isinstance(plain, (dt.datetime, dt.timedelta)):  # too far out for a datetime
            raise CellTypeError(f"{value!r} is outside the dates Excel can store")
        return plain
    if hasattr(value, "item"):  # any other numpy scalar
        plain = value.item()
        if name in ("float16", "float32"):  # keep the digits it prints, not float64 noise
            plain = float(str(value))
        return normalize_cell_value(plain) if type(plain) in _PLAIN else plain
    return value


def check_dimensions(n_rows: int, n_cols: int) -> None:
    """Raise DimensionError if a grid of this size won't fit in an .xlsx sheet."""
    if n_rows > MAX_ROWS:
        raise DimensionError(f"{n_rows} rows exceeds the .xlsx limit of {MAX_ROWS}")
    if n_cols > MAX_COLUMNS:
        raise DimensionError(f"{n_cols} columns exceeds the .xlsx limit of {MAX_COLUMNS}")


def check_cell_value(value: object) -> None:
    """Raise CellTypeError if *value* is not something Excel can store in a cell.

    Allowed: ``str``, ``int``, ``float``, ``bool``, ``datetime``, ``date``,
    ``time``, ``timedelta``, and ``None`` (an empty cell) — and the numpy and pandas
    equivalents, which are stored as the plain value (see :func:`normalize_cell_value`).
    A missing value — ``nan``, ``pd.NaT``, ``pd.NA`` — is an empty cell, so it is allowed.
    Rejected even though they have an allowed type, because Excel would silently change or
    lose them (or, for the first, make the whole file unreadable):

    * a string with a character XML can't hold (a control character such as
      NUL, a lone surrogate, U+FFFE or U+FFFF), or longer than a cell holds
      (32,767 characters — longer is silently truncated);
    * the infinities (``nan`` is fine: it is stored as an empty cell);
    * an ``int`` too large for a float;
    * a date before 1900-01-01, or after 9999-12-31 23:59:59.999;
    * a timezone-aware ``datetime``/``time`` — Excel has no timezone.
    """
    value = normalize_cell_value(value)
    if value is None:
        return
    if not isinstance(value, _CELL_TYPES):
        raise CellTypeError(f"{type(value).__name__} is not a type Excel can store: {value!r}")

    if isinstance(value, str):
        if len(value) > MAX_CELL_CHARS:
            raise CellTypeError(
                f"a string of {len(value):,} characters is longer than an Excel cell "
                f"holds ({MAX_CELL_CHARS:,}) and would be silently truncated"
            )
        bad = _XML_ILLEGAL.search(value)
        if bad:
            raise CellTypeError(
                f"the string contains U+{ord(bad.group()):04X}, which an .xlsx file cannot "
                "hold (it would make the whole workbook unreadable)"
            )
    elif isinstance(value, float):
        if not math.isfinite(value):
            raise CellTypeError(f"Excel cannot store {value!r}")
    elif isinstance(value, int):
        if abs(value) > sys.float_info.max:
            raise CellTypeError("an integer this large cannot be stored as an Excel number")
    elif isinstance(value, (dt.datetime, dt.time)) and value.tzinfo is not None:
        raise CellTypeError(
            f"{type(value).__name__} must be timezone-naive; Excel has no timezone concept"
        )
    if isinstance(value, dt.datetime):
        if not _EXCEL_MIN <= value <= _EXCEL_MAX:
            raise CellTypeError(
                f"{value!r} is outside the dates Excel can store "
                "(1900-01-01 to 9999-12-31 23:59:59.999)"
            )
    elif isinstance(value, dt.date) and value < _EXCEL_MIN.date():
        raise CellTypeError(f"{value!r} is before 1900-01-01, the earliest date Excel can store")


def normalize_newlines(value: object) -> object:
    """A string with each carriage return turned into a newline; anything else unchanged.

    Both ``"\\r\\n"`` and a lone ``"\\r"`` become ``"\\n"``, which is how Excel stores a
    line break in a cell. Left alone, a carriage return is handled by the XML layer
    differently per platform — on Windows ``"a\\r\\nb"`` came back as ``"a\\n\\nb"`` (a
    blank line gained), on Linux as ``"a\\nb"`` — so one call stored different text on
    different machines.
    """
    if isinstance(value, str) and "\r" in value:
        return value.replace("\r\n", "\n").replace("\r", "\n")
    return value


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

    control = sorted({c for c in sheet_name if ord(c) < 0x20 or _XML_ILLEGAL.match(c)})
    if control:
        raise SheetNameError(
            f"sheet name {sheet_name!r} contains control or otherwise unstorable "
            f"character(s): {[f'U+{ord(c):04X}' for c in control]}"
        )

    if sheet_name.startswith("'") or sheet_name.endswith("'"):
        raise SheetNameError(f"sheet name {sheet_name!r} must not begin or end with an apostrophe")

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
