"""Restricting a Table column to one Excel-native data type."""

from __future__ import annotations

import datetime as dt
from enum import Enum


class ColumnType(Enum):
    """A restriction on what values a Table column's data may hold.

    Follows Excel's own type model, not Python's: ``NUMBER`` covers both
    ``int`` and ``float`` (Excel stores every number as a float and doesn't
    distinguish them — see the round-trip notes in the README), and
    ``DATE`` covers both ``datetime.date`` and ``datetime.datetime`` (Excel
    has no date-only type). ``ANY`` — the default for every column unless
    stated otherwise — allows anything :func:`pyhandlexl.check_cell_value`
    itself allows.

    A blank cell (``None``) is always allowed regardless of a column's
    type: the restriction governs what a real value may be, not whether
    the cell has been filled in yet.
    """

    ANY = "any"
    TEXT = "text"
    NUMBER = "number"
    BOOLEAN = "boolean"
    DATE = "date"
    TIME = "time"
    DURATION = "duration"

    def allows(self, value: object) -> bool:
        """Whether *value* is allowed in a column restricted to this type."""
        if value is None or self is ColumnType.ANY:
            return True
        if self is ColumnType.TEXT:
            return isinstance(value, str)
        if self is ColumnType.NUMBER:
            return isinstance(value, (int, float)) and not isinstance(value, bool)
        if self is ColumnType.BOOLEAN:
            return isinstance(value, bool)
        if self is ColumnType.DATE:
            return isinstance(value, dt.date)  # datetime is a subclass of date
        if self is ColumnType.TIME:
            return isinstance(value, dt.time)
        if self is ColumnType.DURATION:
            return isinstance(value, dt.timedelta)
        raise AssertionError(self)  # pragma: no cover
