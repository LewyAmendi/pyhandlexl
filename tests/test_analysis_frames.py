"""Table <-> pandas DataFrame, Table -> numpy array, and grid <-> DataFrame.

pandas and numpy are optional extras, so this file skips when either is missing.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
import textwrap

import pytest

from pyhandlexl import (
    CellTypeError,
    ColumnType,
    ColumnTypeError,
    Table,
    TableStyle,
    create_sheet,
    grid,
    read_sheet,
    write_sheet,
)

np = pytest.importorskip("numpy")
pd = pytest.importorskip("pandas")

WHEN = dt.datetime(2026, 1, 5, 12, 30)


@pytest.fixture
def fruit() -> Table:
    return Table(
        data=[
            [10, 2.5, dt.datetime(2026, 1, 5), True, "x", dt.timedelta(hours=2)],
            [20, None, dt.datetime(2026, 2, 6), False, "y", dt.timedelta(minutes=30)],
            [30, 4.5, None, True, None, None],
        ],
        row_labels=["apple", "pear", "plum"],
        column_headers=["qty", "price", "when", "ok", "note", "spent"],
        corner="fruit",
        name="Fruit",
    )


def a_typical_frame() -> pd.DataFrame:
    rng = np.random.default_rng(0)
    df = pd.DataFrame(
        {
            "qty": rng.integers(1, 100, 6),
            "price": rng.normal(10, 2, 6).round(2),
            "when": pd.date_range("2026-01-01", periods=6, freq="D"),
            "ok": [True, False, True, True, False, True],
            "note": list("abcdef"),
            "spent": pd.to_timedelta(rng.integers(1, 90, 6), unit="min"),
        },
        index=[f"r{i}" for i in range(6)],
    )
    df.index.name = "id"
    return df


class TestToDataFrame:
    def test_labels_headers_corner_and_values(self, fruit):
        df = fruit.to_dataframe()
        assert list(df.index) == ["apple", "pear", "plum"] and df.index.name == "fruit"
        assert list(df.columns) == ["qty", "price", "when", "ok", "note", "spent"]
        assert df.loc["apple", "qty"] == 10 and df.loc["pear", "note"] == "y"

    def test_dtypes_are_inferred_from_the_values(self, fruit):
        d = fruit.to_dataframe().dtypes
        assert str(d["qty"]) == "int64" and str(d["price"]) == "float64"
        assert str(d["when"]).startswith("datetime64") and str(d["spent"]).startswith("timedelta64")
        assert str(d["ok"]) == "bool"

    def test_blanks_are_missing_values(self, fruit):
        df = fruit.to_dataframe()
        assert pd.isna(df.loc["pear", "price"]) and pd.isna(df.loc["plum", "when"])
        assert pd.isna(df.loc["plum", "note"]) and pd.isna(df.loc["plum", "spent"])

    def test_no_corner_means_no_index_name(self):
        t = Table(data=[[1]], row_labels=["a"], column_headers=["x"], name="T")
        assert t.to_dataframe().index.name is None

    def test_it_is_a_copy(self, fruit):
        df = fruit.to_dataframe()
        df.loc["apple", "qty"] = 999
        assert fruit.read_cell(row="apple", column="qty") == 10

    def test_a_table_with_no_rows_is_an_empty_frame_with_its_columns(self):
        df = Table(column_headers=["a", "b"], name="T").to_dataframe()
        assert df.shape == (0, 2) and list(df.columns) == ["a", "b"]

    def test_a_table_read_from_a_file(self, book, fruit):
        create_sheet(book, "D")
        fruit.create(book, sheet="D")
        df = Table.read(book, "Fruit").to_dataframe()
        assert df.shape == (3, 6) and df.index.name == "fruit"
        assert df.loc["plum", "qty"] == 30

    def test_column_types_settle_the_dtypes_inference_cannot(self):
        t = Table(
            data=[[True, None, None, None], [None, None, None, None]],
            row_labels=["a", "b"],
            column_headers=["flag", "nothing", "day", "took"],
            name="T",
            column_types={
                "flag": ColumnType.BOOLEAN,
                "nothing": ColumnType.NUMBER,
                "day": ColumnType.DATE,
                "took": ColumnType.DURATION,
            },
        )
        d = t.to_dataframe().dtypes
        assert str(d["flag"]) == "boolean"  # nullable, not object
        assert str(d["nothing"]) == "float64"  # not object
        assert str(d["day"]).startswith("datetime64") and str(d["took"]).startswith("timedelta64")

    def test_a_boolean_column_without_blanks_is_plain_bool(self):
        t = Table(
            data=[[True], [False]],
            row_labels=["a", "b"],
            column_headers=["f"],
            name="T",
            column_types={"f": ColumnType.BOOLEAN},
        )
        assert str(t.to_dataframe().dtypes["f"]) == "bool"

    def test_dates_and_datetimes_mixed_become_one_datetime_column(self):
        t = Table(
            data=[[dt.date(2026, 1, 1)], [dt.datetime(2026, 1, 2, 3)], [None]],
            row_labels=["a", "b", "c"],
            column_headers=["d"],
            name="T",
            column_types={"d": ColumnType.DATE},
        )
        col = t.to_dataframe()["d"]
        assert str(col.dtype).startswith("datetime64") and col.isna().tolist() == [
            False,
            False,
            True,
        ]

    def test_a_value_that_does_not_fit_its_column_type_keeps_its_inferred_dtype(self):
        t = Table(
            data=[["text"], [1]],
            row_labels=["a", "b"],
            column_headers=["n"],
            name="T",
            column_types={"n": ColumnType.DATE},  # nothing stops this until write()
        )
        assert t.to_dataframe()["n"].tolist() == ["text", 1]

    def test_whole_number_floats_come_back_as_int64_because_excel_has_one_number_kind(self, book):
        df = pd.DataFrame({"x": [1.0, 2.0, 3.0]}, index=["a", "b", "c"])
        create_sheet(book, "D")
        Table.from_dataframe(df, name="T").create(book, sheet="D")
        assert str(Table.read(book, "T").to_dataframe().dtypes["x"]) == "int64"


class TestFromDataFrame:
    def test_index_columns_and_values(self):
        t = Table.from_dataframe(a_typical_frame(), name="T")
        assert t.data.row_labels == [f"r{i}" for i in range(6)]
        assert t.data.column_headers == ["qty", "price", "when", "ok", "note", "spent"]
        assert t.data.corner == "id"
        assert {type(v) for row in t.data.rows for v in row[:1]} == {int}
        assert {type(row[2]) for row in t.data.rows} == {dt.datetime}
        assert {type(row[5]) for row in t.data.rows} == {dt.timedelta}

    def test_every_missing_value_is_a_blank_cell(self):
        df = pd.DataFrame(
            {
                "f": [1.5, np.nan],
                "t": [pd.Timestamp("2026-01-05"), pd.NaT],
                "n": pd.array([1, None], dtype="Int64"),
                "s": ["x", None],
                "b": pd.array([True, None], dtype="boolean"),
            },
            index=["a", "b"],
        )
        t = Table.from_dataframe(df, name="T")
        assert t.data.rows[1] == [None, None, None, None, None]
        assert t.data.rows[0][:3] == [1.5, dt.datetime(2026, 1, 5), 1]

    @pytest.mark.parametrize(
        ("series", "expected"),
        [
            (pd.Series([1, 2], dtype="int32"), [1, 2]),
            (pd.Series([1, 2], dtype="uint8"), [1, 2]),
            (pd.Series([0.5, 0.25], dtype="float32"), [0.5, 0.25]),
            (pd.Series([True, False]), [True, False]),
            (pd.Series(["a", "b"], dtype="category"), ["a", "b"]),
            (pd.Series(["a", "b"], dtype="string"), ["a", "b"]),
            (pd.Series([1, 2], dtype="Int64"), [1, 2]),
            (pd.Series([WHEN, WHEN]), [WHEN, WHEN]),
        ],
        ids=lambda v: str(getattr(v, "dtype", v))[:20],
    )
    def test_every_common_dtype_is_stored_as_plain_values(self, series, expected):
        t = Table.from_dataframe(pd.DataFrame({"c": series.to_numpy()}, index=["a", "b"]), name="T")
        assert [row[0] for row in t.data.rows] == expected
        t2 = Table.from_dataframe(pd.DataFrame({"c": series.set_axis(["a", "b"])}), name="T")
        assert [row[0] for row in t2.data.rows] == expected

    def test_the_style_and_column_types_are_passed_on(self):
        t = Table.from_dataframe(
            a_typical_frame(),
            name="T",
            style=TableStyle.MINIMAL,
            column_types={"note": ColumnType.TEXT},
        )
        assert t.style == TableStyle.MINIMAL
        assert t.column_types["note"] is ColumnType.TEXT and t.column_types["qty"] is ColumnType.ANY

    def test_columns_are_unrestricted_by_default(self):
        types = Table.from_dataframe(a_typical_frame(), name="T").column_types
        assert set(types.values()) == {ColumnType.ANY}

    def test_it_can_infer_each_columns_type(self):
        t = Table.from_dataframe(a_typical_frame(), name="T", infer_column_types=True)
        assert t.column_types == {
            "qty": ColumnType.NUMBER,
            "price": ColumnType.NUMBER,
            "when": ColumnType.DATE,
            "ok": ColumnType.BOOLEAN,
            "note": ColumnType.TEXT,
            "spent": ColumnType.DURATION,
        }

    def test_a_mixed_or_empty_column_stays_unrestricted_when_inferring(self):
        df = pd.DataFrame({"mixed": [1, "a"], "empty": [None, None]}, index=["a", "b"])
        t = Table.from_dataframe(df, name="T", infer_column_types=True)
        assert set(t.column_types.values()) == {ColumnType.ANY}

    def test_explicit_column_types_win_over_inferred_ones(self):
        t = Table.from_dataframe(
            a_typical_frame(),
            name="T",
            infer_column_types=True,
            column_types={"qty": ColumnType.ANY},
        )
        assert (
            t.column_types["qty"] is ColumnType.ANY and t.column_types["price"] is ColumnType.NUMBER
        )

    def test_inferred_types_are_enforced_on_write(self, book):
        create_sheet(book, "D")
        t = Table.from_dataframe(a_typical_frame(), name="T", infer_column_types=True)
        t.set_cell(row="r0", column="qty", value="not a number")
        with pytest.raises(ColumnTypeError):
            t.create(book, sheet="D")

    def test_a_timezone_aware_column_is_refused_at_write_with_a_clear_reason(self, book):
        df = pd.DataFrame({"t": pd.date_range("2026-01-01", periods=2, tz="UTC")}, index=["a", "b"])
        create_sheet(book, "D")
        with pytest.raises(CellTypeError, match="timezone"):
            Table.from_dataframe(df, name="T").create(book, sheet="D")

    def test_the_frame_is_not_modified(self):
        df = a_typical_frame()
        before = df.copy()
        Table.from_dataframe(df, name="T")
        pd.testing.assert_frame_equal(df, before)


class TestFromDataFrameRefuses:
    def test_an_integer_index_with_a_way_out(self):
        df = pd.DataFrame({"x": [1, 2]})  # the default RangeIndex
        with pytest.raises(TypeError) as caught:
            Table.from_dataframe(df, name="T")
        message = str(caught.value)
        assert "index labels must be strings" in message and "int (0)" in message
        assert "astype(str)" in message and "set_index" in message

    def test_a_non_string_index_of_any_kind(self):
        for index in ([1.5, 2.5], [dt.date(2026, 1, 1), dt.date(2026, 1, 2)], [("a", 1), ("b", 2)]):
            df = pd.DataFrame({"x": [1, 2]}, index=pd.Index(index, tupleize_cols=False))
            with pytest.raises(TypeError, match="index labels must be strings"):
                Table.from_dataframe(df, name="T")

    def test_a_mixed_index_names_the_offender(self):
        df = pd.DataFrame({"x": [1, 2]}, index=["ok", 7])
        with pytest.raises(TypeError, match=r"int \(7\)"):
            Table.from_dataframe(df, name="T")

    def test_non_string_column_names(self):
        df = pd.DataFrame([[1, 2]], columns=[2024, 2025], index=["a"])
        with pytest.raises(TypeError, match=r"column names must be strings.*astype\(str\)"):
            Table.from_dataframe(df, name="T")

    def test_a_multiindex_on_either_axis(self):
        rows = pd.MultiIndex.from_tuples([("a", "x"), ("a", "y")])
        with pytest.raises(TypeError, match="one level"):
            Table.from_dataframe(pd.DataFrame({"c": [1, 2]}, index=rows), name="T")
        cols = pd.MultiIndex.from_tuples([("a", "x"), ("a", "y")])
        with pytest.raises(TypeError, match="one level"):
            Table.from_dataframe(pd.DataFrame([[1, 2]], index=["r"], columns=cols), name="T")

    def test_a_non_string_index_name(self):
        df = pd.DataFrame({"x": [1]}, index=pd.Index(["a"], name=5))
        with pytest.raises(TypeError, match="corner"):
            Table.from_dataframe(df, name="T")

    @pytest.mark.parametrize(
        "not_a_frame", [None, {"a": 1}, [[1]], pd.Series([1, 2]), np.zeros((2, 2))]
    )
    def test_something_that_is_not_a_dataframe(self, not_a_frame):
        with pytest.raises(TypeError, match="expected a pandas DataFrame"):
            Table.from_dataframe(not_a_frame, name="T")

    def test_the_fix_the_message_suggests_works(self):
        df = pd.DataFrame({"x": [1, 2], "who": ["a", "b"]})
        assert Table.from_dataframe(df.set_index("who"), name="T").data.row_labels == ["a", "b"]
        df.index = df.index.astype(str)
        assert Table.from_dataframe(df.drop(columns="who"), name="T").data.row_labels == ["0", "1"]


class TestRoundTripThroughAFile:
    def test_a_typical_frame_survives(self, book):
        df = a_typical_frame()
        df.loc["r2", "price"] = np.nan
        df.loc["r4", "when"] = pd.NaT
        df.loc["r1", "note"] = None
        create_sheet(book, "D")
        Table.from_dataframe(df, name="T", infer_column_types=True).create(book, sheet="D")
        back = Table.read(book, "T").to_dataframe()
        assert back.index.name == "id" and list(back.index) == list(df.index)
        assert list(back.columns) == list(df.columns)
        for column in df.columns:
            assert df[column].isna().tolist() == back[column].isna().tolist(), column
            assert (
                df[column].dropna().astype(object).tolist()
                == back[column].dropna().astype(object).tolist()
            ), column

    def test_a_big_frame_is_identical(self, book):
        rng = np.random.default_rng(1)
        df = pd.DataFrame(
            rng.normal(size=(500, 6)),
            columns=list("abcdef"),
            index=[f"r{i}" for i in range(500)],
        )
        create_sheet(book, "D")
        Table.from_dataframe(df, name="T").create(book, sheet="D")
        back = Table.read(book, "T").to_dataframe()
        assert np.allclose(df.to_numpy(), back.to_numpy())

    def test_replacing_a_table_from_a_frame(self, book, fruit):
        create_sheet(book, "D")
        fruit.create(book, sheet="D")
        df = pd.DataFrame({"qty": [1, 2]}, index=["kiwi", "fig"])
        Table.from_dataframe(df, name="Fruit").write(book)
        assert Table.read(book, "Fruit").data.rows == [[1], [2]]

    def test_a_table_can_be_edited_through_a_frame_and_written_back(self, book, fruit):
        create_sheet(book, "D")
        fruit.create(book, sheet="D")
        df = Table.read(book, "Fruit").to_dataframe()
        df["qty"] = df["qty"] * 2
        edited = Table.from_dataframe(df, name="Fruit", style=fruit.style)
        edited.write(book)
        assert Table.read(book, "Fruit").read_column("qty") == [20, 40, 60]


class TestToNumpy:
    def test_integers(self):
        t = Table(data=[[1, 2], [3, 4]], row_labels=["a", "b"], column_headers=["x", "y"], name="T")
        arr = t.to_numpy()
        assert arr.dtype == np.int64 and arr.shape == (2, 2) and arr.tolist() == [[1, 2], [3, 4]]

    def test_floats_and_blanks_become_nan(self):
        t = Table(
            data=[[1.5, None], [3, 4]], row_labels=["a", "b"], column_headers=["x", "y"], name="T"
        )
        arr = t.to_numpy()
        assert arr.dtype == np.float64 and np.isnan(arr[0, 1]) and arr[1, 0] == 3.0

    def test_integers_with_a_blank_are_floats(self):
        t = Table(data=[[1], [None]], row_labels=["a", "b"], column_headers=["x"], name="T")
        assert t.to_numpy().dtype == np.float64

    def test_booleans(self):
        t = Table(data=[[True], [False]], row_labels=["a", "b"], column_headers=["x"], name="T")
        assert t.to_numpy().dtype == np.bool_

    def test_booleans_with_a_blank_stay_objects(self):
        t = Table(data=[[True], [None]], row_labels=["a", "b"], column_headers=["x"], name="T")
        assert t.to_numpy().dtype == object and t.to_numpy()[1, 0] is None

    def test_datetimes_and_durations(self):
        t = Table(data=[[WHEN], [None]], row_labels=["a", "b"], column_headers=["x"], name="T")
        arr = t.to_numpy()
        assert arr.dtype == np.dtype("datetime64[us]") and np.isnat(arr[1, 0])
        d = Table(
            data=[[dt.timedelta(hours=1)], [None]],
            row_labels=["a", "b"],
            column_headers=["x"],
            name="T",
        )
        assert d.to_numpy().dtype == np.dtype("timedelta64[us]") and np.isnat(d.to_numpy()[1, 0])

    def test_text_or_a_mixture_is_an_object_array_keeping_none(self):
        t = Table(
            data=[["a", 1], [None, 2.5]], row_labels=["a", "b"], column_headers=["x", "y"], name="T"
        )
        arr = t.to_numpy()
        assert arr.dtype == object and arr.tolist() == [["a", 1], [None, 2.5]]

    def test_no_rows_is_an_empty_array_of_the_right_width(self):
        assert Table(column_headers=["a", "b", "c"], name="T").to_numpy().shape == (0, 3)

    def test_only_blanks_is_nan(self):
        t = Table(data=[[None]], row_labels=["a"], column_headers=["x"], name="T")
        assert np.isnan(t.to_numpy()[0, 0])

    def test_an_explicit_dtype(self):
        t = Table(data=[[1, 2]], row_labels=["a"], column_headers=["x", "y"], name="T")
        assert t.to_numpy(dtype=float).dtype == np.float64
        assert t.to_numpy(dtype=object).dtype == object
        assert t.to_numpy(dtype="float32").dtype == np.float32

    def test_a_blank_cannot_be_an_integer_and_the_message_says_what_to_do(self):
        t = Table(data=[[1], [None]], row_labels=["a", "b"], column_headers=["x"], name="T")
        with pytest.raises(TypeError, match="float or object dtype"):
            t.to_numpy(dtype=int)

    def test_the_labels_are_on_data(self):
        t = Table(data=[[1]], row_labels=["a"], column_headers=["x"], corner="c", name="T")
        assert (t.data.row_labels, t.data.column_headers) == (["a"], ["x"])

    def test_the_array_is_a_copy(self):
        t = Table(data=[[1]], row_labels=["a"], column_headers=["x"], name="T")
        t.to_numpy()[0, 0] = 99
        assert t.read_cell(row="a", column="x") == 1

    def test_numbers_are_ready_for_numpy_arithmetic(self, book, fruit):
        create_sheet(book, "D")
        fruit.create(book, sheet="D")
        qty = Table.read(book, "Fruit").to_numpy(dtype=object)[:, 0].astype(float)
        assert qty.mean() == 20.0


class TestGridFrames:
    def test_to_dataframe_with_a_header_row(self):
        df = grid.to_dataframe([["a", "b"], [1, 2], [3, 4]])
        assert list(df.columns) == ["a", "b"] and df["b"].tolist() == [2, 4]

    def test_to_dataframe_without_a_header(self):
        df = grid.to_dataframe([[1, 2], [3, 4]], header=False)
        assert list(df.columns) == [0, 1] and df.shape == (2, 2)

    def test_ragged_rows_are_padded(self):
        df = grid.to_dataframe([["a", "b", "c"], [1], [1, 2, 3]])
        assert df.shape == (2, 3) and pd.isna(df.loc[0, "b"])

    def test_a_header_only_and_an_empty_grid(self):
        assert grid.to_dataframe([["a", "b"]]).shape == (0, 2)
        assert grid.to_dataframe([]).shape == (0, 0)

    def test_it_takes_what_read_sheet_returns(self, book):
        write_sheet(book, [["n", "when"], [1, WHEN], [2, None]])
        df = grid.to_dataframe(read_sheet(book))
        assert str(df["when"].dtype).startswith("datetime64") and df["n"].tolist() == [1, 2]

    def test_from_dataframe_with_a_header(self):
        df = pd.DataFrame({"a": [1, 2], "b": ["x", None]})
        assert grid.from_dataframe(df) == [["a", "b"], [1, "x"], [2, None]]

    def test_from_dataframe_without_a_header_or_with_the_index(self):
        df = pd.DataFrame({"a": [1.5, np.nan]}, index=pd.Index(["r1", "r2"], name="id"))
        assert grid.from_dataframe(df, header=False) == [[1.5], [None]]
        assert grid.from_dataframe(df, index=True) == [["id", "a"], ["r1", 1.5], ["r2", None]]
        assert grid.from_dataframe(df, header=False, index=True) == [["r1", 1.5], ["r2", None]]

    def test_the_index_can_be_anything_in_a_grid(self):
        df = pd.DataFrame({"a": [1, 2]})  # a RangeIndex is fine here: cells can hold numbers
        assert grid.from_dataframe(df, index=True) == [[None, "a"], [0, 1], [1, 2]]

    def test_a_multiindex_is_refused_when_it_would_be_written(self):
        df = pd.DataFrame({"a": [1, 2]}, index=pd.MultiIndex.from_tuples([("x", 1), ("x", 2)]))
        with pytest.raises(TypeError, match="MultiIndex"):
            grid.from_dataframe(df, index=True)
        assert grid.from_dataframe(df) == [["a"], [1], [2]]

    def test_not_a_dataframe(self):
        with pytest.raises(TypeError, match="expected a pandas DataFrame"):
            grid.from_dataframe([[1]])

    def test_write_sheet_round_trip(self, book):
        df = pd.DataFrame(
            {
                "n": [1, 2, 3],
                "x": [0.5, np.nan, 2.5],
                "when": pd.to_datetime(["2026-01-05", None, "2026-01-07"]),
            }
        )
        write_sheet(book, grid.from_dataframe(df))
        back = grid.to_dataframe(read_sheet(book))
        assert back["n"].tolist() == [1, 2, 3]
        assert back["x"].isna().tolist() == [False, True, False]
        assert back["when"].isna().tolist() == [False, True, False]

    def test_values_are_plain_python(self):
        rows = grid.from_dataframe(pd.DataFrame({"a": np.array([1, 2], dtype="int32")}))
        assert {type(v) for row in rows[1:] for v in row} == {int}


class TestWithoutTheLibraries:
    def run(self, code: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, "-c", textwrap.dedent(code)], capture_output=True, text=True
        )

    def test_a_missing_pandas_says_how_to_get_it(self):
        result = self.run(
            """
            import sys
            sys.modules["pandas"] = None
            import pyhandlexl as p
            t = p.Table(data=[[1]], row_labels=["a"], column_headers=["x"], name="T")
            for call in (t.to_dataframe, lambda: p.grid.to_dataframe([[1]]),
                         lambda: p.Table.from_dataframe(object(), name="T"),
                         lambda: p.grid.from_dataframe(object())):
                try:
                    call()
                except ImportError as error:
                    assert 'pip install "pyhandlexl[pandas]"' in str(error), error
                else:
                    raise SystemExit("no ImportError")
            print("ok")
            """
        )
        assert result.returncode == 0 and "ok" in result.stdout, result.stderr

    def test_a_missing_numpy_says_how_to_get_it(self):
        result = self.run(
            """
            import sys
            sys.modules["numpy"] = None
            import pyhandlexl as p
            t = p.Table(data=[[1]], row_labels=["a"], column_headers=["x"], name="T")
            try:
                t.to_numpy()
            except ImportError as error:
                assert 'pip install "pyhandlexl[numpy]"' in str(error), error
                print("ok")
            """
        )
        assert result.returncode == 0 and "ok" in result.stdout, result.stderr

    def test_the_rest_of_the_library_does_not_need_either(self):
        result = self.run(
            """
            import sys
            sys.modules["numpy"] = None
            sys.modules["pandas"] = None
            import pyhandlexl as p
            t = p.Table(data=[[1, None]], row_labels=["a"], column_headers=["x", "y"], name="T")
            assert t.to_dict() == {"a": {"x": 1, "y": None}}
            print("ok")
            """
        )
        assert result.returncode == 0 and "ok" in result.stdout, result.stderr
