"""Tests for the ColumnType enum: what each type allows."""

from __future__ import annotations

import datetime as dt

from pyhandlexl import ColumnType

_SAMPLES = {
    "str": "hello",
    "int": 5,
    "float": 5.5,
    "bool": True,
    "date": dt.date(2026, 1, 1),
    "datetime": dt.datetime(2026, 1, 1, 9, 0),
    "time": dt.time(9, 0),
    "timedelta": dt.timedelta(hours=1),
    "dict": {"not": "excel-storable"},
}


class TestAny:
    def test_allows_everything(self):
        for value in _SAMPLES.values():
            assert ColumnType.ANY.allows(value)

    def test_allows_none(self):
        assert ColumnType.ANY.allows(None)


class TestNoneAlwaysAllowed:
    def test_every_type_allows_none(self):
        for column_type in ColumnType:
            assert column_type.allows(None)


class TestText:
    def test_allows_str(self):
        assert ColumnType.TEXT.allows("hello")

    def test_rejects_everything_else(self):
        for key, value in _SAMPLES.items():
            if key != "str":
                assert not ColumnType.TEXT.allows(value), key


class TestNumber:
    def test_allows_int_and_float(self):
        assert ColumnType.NUMBER.allows(5)
        assert ColumnType.NUMBER.allows(5.5)

    def test_rejects_bool(self):
        # bool is technically an int subclass in Python, but not a "number"
        # for this purpose — same gotcha TableStyle's font-size validation
        # already guards against.
        assert not ColumnType.NUMBER.allows(True)
        assert not ColumnType.NUMBER.allows(False)

    def test_rejects_everything_else(self):
        for key, value in _SAMPLES.items():
            if key not in ("int", "float"):
                assert not ColumnType.NUMBER.allows(value), key


class TestBoolean:
    def test_allows_bool(self):
        assert ColumnType.BOOLEAN.allows(True)
        assert ColumnType.BOOLEAN.allows(False)

    def test_rejects_int(self):
        assert not ColumnType.BOOLEAN.allows(1)
        assert not ColumnType.BOOLEAN.allows(0)

    def test_rejects_everything_else(self):
        for key, value in _SAMPLES.items():
            if key != "bool":
                assert not ColumnType.BOOLEAN.allows(value), key


class TestDate:
    def test_allows_date_and_datetime(self):
        assert ColumnType.DATE.allows(dt.date(2026, 1, 1))
        assert ColumnType.DATE.allows(dt.datetime(2026, 1, 1, 9, 0))

    def test_rejects_everything_else(self):
        for key, value in _SAMPLES.items():
            if key not in ("date", "datetime"):
                assert not ColumnType.DATE.allows(value), key


class TestTime:
    def test_allows_time(self):
        assert ColumnType.TIME.allows(dt.time(9, 0))

    def test_rejects_everything_else(self):
        for key, value in _SAMPLES.items():
            if key != "time":
                assert not ColumnType.TIME.allows(value), key


class TestDuration:
    def test_allows_timedelta(self):
        assert ColumnType.DURATION.allows(dt.timedelta(hours=1))

    def test_rejects_everything_else(self):
        for key, value in _SAMPLES.items():
            if key != "timedelta":
                assert not ColumnType.DURATION.allows(value), key
