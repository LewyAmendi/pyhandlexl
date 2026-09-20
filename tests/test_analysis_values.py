"""Values from numpy and pandas are accepted and stored as the plain values they hold, and
every kind of missing value (nan, NaT, NA) is an empty cell.

numpy and pandas are optional: none of the library imports them, so the checks that need
one skip when it isn't installed.
"""

from __future__ import annotations

import datetime as dt
import decimal
import math

import pytest

from pyhandlexl import (
    CellTypeError,
    Table,
    append_rows,
    check_cell_value,
    create_sheet,
    read_sheet,
    write_sheet,
)
from pyhandlexl.validate import normalize_cell_value

np = pytest.importorskip("numpy")
pd = pytest.importorskip("pandas")

WHEN = dt.datetime(2026, 1, 5, 12, 30)


class TestNormalize:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (np.int64(3), 3),
            (np.int32(-7), -7),
            (np.uint8(200), 200),
            (np.float64(2.5), 2.5),
            (np.float32(1.5), 1.5),
            (np.float32(0.1), 0.1),  # the digits it prints, not float64 noise
            (np.float16(0.5), 0.5),
            (np.bool_(True), True),
            (np.bool_(False), False),
            (np.str_("s"), "s"),
            (np.datetime64("2026-01-05T12:30:00"), WHEN),
            (np.datetime64("2026-01-05"), dt.datetime(2026, 1, 5)),
            (np.datetime64("2026-01-05T12:30:00.000000001"), WHEN),  # nanoseconds are dropped
            (np.timedelta64(90, "m"), dt.timedelta(minutes=90)),
            (np.timedelta64(5, "s"), dt.timedelta(seconds=5)),
            (pd.Timestamp("2026-01-05 12:30"), WHEN),
            (pd.Timedelta(hours=2), dt.timedelta(hours=2)),
        ],
        ids=lambda v: repr(v)[:28],
    )
    def test_becomes_the_plain_value(self, value, expected):
        got = normalize_cell_value(value)
        assert got == expected
        assert type(got) in (int, float, bool, str, dt.datetime, dt.timedelta)

    @pytest.mark.parametrize(
        "value",
        [
            float("nan"),
            np.nan,
            np.float64("nan"),
            np.float32("nan"),
            pd.NaT,
            pd.NA,
            np.datetime64("NaT", "s"),
            np.timedelta64("NaT", "s"),
            None,
        ],
        ids=lambda v: repr(v),
    )
    def test_every_kind_of_missing_is_none(self, value):
        assert normalize_cell_value(value) is None

    @pytest.mark.parametrize("value", [1, 2.5, True, "s", "", WHEN, dt.date(2026, 1, 1)])
    def test_plain_values_pass_through_untouched(self, value):
        assert normalize_cell_value(value) is value

    @pytest.mark.parametrize("value", [decimal.Decimal("1.5"), b"x", [1], {"a": 1}, 1j])
    def test_a_foreign_value_is_left_for_the_check_to_refuse(self, value):
        assert normalize_cell_value(value) is value

    def test_infinity_is_not_missing(self):
        assert normalize_cell_value(float("inf")) == math.inf
        assert normalize_cell_value(np.float64("-inf")) == -math.inf

    def test_a_timestamp_with_nanoseconds_does_not_warn(self, recwarn):
        normalize_cell_value(pd.Timestamp("2026-01-05 12:30:00.123456789"))
        assert not [w for w in recwarn if "nanosecond" in str(w.message).lower()]

    def test_a_datetime64_beyond_a_datetime_is_refused_not_turned_into_a_number(self):
        with pytest.raises(CellTypeError, match="outside the dates Excel can store"):
            normalize_cell_value(np.datetime64("12000-01-01"))

    def test_a_numpy_scalar_with_no_python_equivalent_is_refused_by_the_check(self):
        with pytest.raises(CellTypeError):
            check_cell_value(np.clongdouble(1))


class TestCheckCellValue:
    @pytest.mark.parametrize(
        "value",
        [
            np.int64(3),
            np.bool_(True),
            np.float32(1.5),
            np.datetime64("2026-01-05"),
            pd.NaT,
            pd.NA,
            float("nan"),
            np.nan,
            np.timedelta64(1, "h"),
            pd.Timestamp("2026-01-05"),
        ],
        ids=lambda v: repr(v)[:30],
    )
    def test_accepts_what_numpy_and_pandas_produce(self, value):
        check_cell_value(value)

    @pytest.mark.parametrize("value", [float("inf"), -math.inf, np.float64("inf")])
    def test_still_refuses_infinity(self, value):
        with pytest.raises(CellTypeError, match="cannot store"):
            check_cell_value(value)

    def test_the_checks_apply_to_the_plain_value(self):
        with pytest.raises(CellTypeError, match="outside the dates"):
            check_cell_value(np.datetime64("1800-01-01"))
        with pytest.raises(CellTypeError, match="timezone"):
            check_cell_value(pd.Timestamp("2026-01-05", tz="UTC"))
        with pytest.raises(CellTypeError):
            check_cell_value(np.str_("a\x00b"))


class TestGridWrites:
    def test_write_sheet_stores_plain_values_and_blanks(self, book):
        write_sheet(
            book,
            [
                [np.int64(3), np.float64(2.5), np.bool_(True), np.str_("x")],
                [np.nan, pd.NaT, pd.NA, float("nan")],
                [pd.Timestamp("2026-01-05"), np.datetime64("2026-01-06"), pd.Timedelta(hours=2)],
            ],
        )
        assert read_sheet(book) == [
            [3, 2.5, True, "x"],
            [],
            [dt.datetime(2026, 1, 5), dt.datetime(2026, 1, 6), dt.timedelta(hours=2)],
        ]

    def test_the_stored_types_are_the_plain_ones(self, book):
        write_sheet(book, [[np.int64(3), np.bool_(True)]])
        assert [type(v) for v in read_sheet(book)[0]] == [int, bool]

    def test_append_rows_too(self, book):
        write_sheet(book, [[1]])
        append_rows(book, [[np.int64(2), np.nan, np.float32(0.25)]])
        assert read_sheet(book) == [[1], [2, None, 0.25]]

    def test_nan_is_a_blank_cell_not_an_error(self, book):
        write_sheet(book, [[1, float("nan"), 3]])
        assert read_sheet(book) == [[1, None, 3]]

    def test_infinity_is_still_refused_and_nothing_is_written(self, book):
        write_sheet(book, [["keep"]])
        with pytest.raises(CellTypeError):
            write_sheet(book, [[float("inf")]])
        assert read_sheet(book) == [["keep"]]


class TestTableValues:
    @pytest.fixture
    def t(self):
        return Table(
            data=[[1, 2], [3, 4]], row_labels=["a", "b"], column_headers=["x", "y"], name="T"
        )

    def test_a_table_built_from_numpy_holds_plain_values(self):
        arr = np.array([[1, 2], [3, 4]])
        t = Table(data=arr, row_labels=["a", "b"], column_headers=["x", "y"], name="T")
        assert t.data.rows == [[1, 2], [3, 4]]
        assert {type(v) for row in t.data.rows for v in row} == {int}

    def test_every_way_of_setting_a_value_normalises_it(self, t):
        t.set_cell(row="a", column="x", value=np.int64(10))
        t.set_cell("C2", value=np.float64("nan"))  # by position: row a, column y
        t.set_row("b", [np.bool_(True), pd.NaT])
        t.add_row("c", [np.int32(5), pd.NA])
        t.insert_row(1, "d", [np.float32(0.5), np.nan])
        t.add_column("z", [np.int8(1), np.int8(2), np.int8(3), np.int8(4)])
        t.set_column("z", [np.int64(9)] * 4)
        assert t.data.rows == [
            [0.5, None, 9],
            [10, None, 9],
            [True, None, 9],
            [5, None, 9],
        ]
        assert type(t.read_cell(row="a", column="x")) is int

    def test_it_round_trips_through_a_file(self, book, t):
        create_sheet(book, "D")
        t.set_row("a", [np.int64(7), np.nan])
        t.add_row("c", [pd.Timestamp("2026-01-05"), np.timedelta64(30, "m")])
        t.create(book, sheet="D")
        assert Table.read(book, "T").data.rows == [
            [7, None],
            [3, 4],
            [dt.datetime(2026, 1, 5), dt.timedelta(minutes=30)],
        ]

    def test_a_column_type_check_sees_the_plain_value(self, book):
        from pyhandlexl import ColumnType

        create_sheet(book, "D")
        t = Table(
            data=[[np.int64(1)], [np.nan]],
            row_labels=["a", "b"],
            column_headers=["n"],
            name="T",
            column_types={"n": ColumnType.NUMBER},
        )
        t.create(book, sheet="D")  # int is a NUMBER and a blank is allowed everywhere
        assert Table.read(book, "T").data.rows == [[1], [None]]

    def test_editing_a_value_to_a_numpy_equal_is_not_an_edit_conflict(self, book, t):
        create_sheet(book, "D")
        t.create(book, sheet="D")
        a, b = Table.read(book, "T"), Table.read(book, "T")
        a.set_cell(row="a", column="x", value=np.int64(1))  # the value it already has
        b.set_cell(row="b", column="y", value=99)
        a.write(book)
        b.write(book)
        assert b.last_merge is None
