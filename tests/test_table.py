"""Tests for the Table class."""

from __future__ import annotations

import datetime as dt

import pytest
from openpyxl import load_workbook

from pyhandlexl.core import read_sheet
from pyhandlexl.table import Table, TableData


@pytest.fixture
def sample():
    # corner    North  South  East
    # Revenue   100    200    150
    # Costs     40     60     55
    return Table(
        data=[["100", "200", "150"], ["40", "60", "55"]],
        column_headers=["North", "South", "East"],
        row_labels=["Revenue", "Costs"],
        corner="Metric",
        name="Sample",
    )


class TestConstruction:
    def test_name_is_required(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]])

    def test_name_is_stored(self):
        t = Table(data=[["1"]], column_headers=["a"], row_labels=["x"], name="Sales")
        assert t.name == "Sales"

    def test_non_string_name_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]], column_headers=["a"], row_labels=["x"], name=123)

    def test_empty_name_raises(self):
        with pytest.raises(ValueError):
            Table(data=[["1"]], column_headers=["a"], row_labels=["x"], name="")

    def test_repr_includes_name(self):
        t = Table(data=[["1"]], column_headers=["a"], row_labels=["x"], name="Sales")
        assert "name='Sales'" in repr(t)

    def test_mismatched_row_labels_raise(self):
        with pytest.raises(ValueError):
            Table(data=[["1"]], column_headers=["a"], row_labels=["x", "y"], name="T")

    def test_mismatched_row_width_raises(self):
        with pytest.raises(ValueError):
            Table(data=[["1", "2"]], column_headers=["a"], row_labels=["x"], name="T")

    def test_column_headers_is_required(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]], row_labels=["x"], name="T")

    def test_empty_column_headers_raises(self):
        with pytest.raises(ValueError):
            Table(data=[], column_headers=[], name="T")

    def test_headers_with_data_but_no_row_labels_raises(self):
        with pytest.raises(ValueError):
            Table(data=[["1", "2"], ["3", "4"]], column_headers=["a", "b"], name="T")

    def test_headers_with_no_rows_yet_is_allowed(self):
        # building up a table before its first add_row() — row_labels is
        # necessarily empty too, since there's nothing to label yet
        t = Table(data=[], column_headers=["a", "b"], name="T")
        assert t.data.column_headers == ["a", "b"]
        assert t.data.rows == []

    def test_bare_string_column_headers_raises(self):
        # a bare str is Iterable[str] — silently splitting "ab" into ["a", "b"]
        # would be a dangerous, invisible footgun rather than a clear mistake
        with pytest.raises(TypeError):
            Table(data=[[1, 2]], column_headers="ab", row_labels=["x"], name="T")

    def test_bare_string_row_labels_raises(self):
        with pytest.raises(TypeError):
            Table(data=[[1]], column_headers=["a"], row_labels="x", name="T")

    def test_bare_string_data_raises(self):
        with pytest.raises(TypeError):
            Table(data="ab", name="T")

    def test_non_string_row_label_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]], column_headers=["a"], row_labels=[1], name="T")

    def test_non_string_column_header_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]], column_headers=[1], row_labels=["x"], name="T")

    def test_non_string_corner_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]], column_headers=["a"], row_labels=["x"], corner=2026, name="T")

    def test_empty_corner_is_allowed(self):
        t = Table(data=[["1"]], column_headers=["a"], row_labels=["x"], corner="", name="T")
        assert t.data.corner == ""

    def test_empty_column_header_raises(self):
        with pytest.raises(ValueError):
            Table(data=[["1"]], column_headers=[""], row_labels=["x"], name="T")

    def test_empty_row_label_raises(self):
        with pytest.raises(ValueError):
            Table(data=[["1"]], column_headers=["a"], row_labels=[""], name="T")

    def test_constructor_allows_duplicate_labels(self):
        # the API blocks *creating* duplicates; ingesting them is allowed
        t = Table(data=[["1"], ["2"]], column_headers=["a"], row_labels=["dup", "dup"], name="T")
        assert t.data.row_labels == ["dup", "dup"]


class TestData:
    def test_data_snapshot(self, sample):
        d = sample.data
        assert isinstance(d, TableData)
        assert d.rows == [["100", "200", "150"], ["40", "60", "55"]]
        assert d.columns == [["100", "40"], ["200", "60"], ["150", "55"]]
        assert d.row_labels == ["Revenue", "Costs"]
        assert d.column_headers == ["North", "South", "East"]
        assert d.corner == "Metric"

    def test_data_is_a_fresh_copy_each_time(self, sample):
        sample.data.rows.append(["999"])
        sample.data.column_headers.append("West")
        assert len(sample.data.rows) == 2
        assert sample.data.column_headers == ["North", "South", "East"]

    def test_corner_is_settable(self, sample):
        sample.set_corner("Quarter")
        assert sample.data.corner == "Quarter"

    def test_setting_non_string_corner_raises(self, sample):
        with pytest.raises(TypeError):
            sample.set_corner(2026)


class TestToFromDict:
    def test_to_dict(self, sample):
        assert sample.to_dict() == {
            "Revenue": {"North": "100", "South": "200", "East": "150"},
            "Costs": {"North": "40", "South": "60", "East": "55"},
        }

    def test_to_dict_duplicate_header_keeps_last_value(self):
        t = Table(data=[[1, 2]], column_headers=["a", "a"], row_labels=["r"], name="T")
        assert t.to_dict() == {"r": {"a": 2}}

    def test_to_dict_duplicate_row_label_keeps_last_value(self):
        t = Table(data=[[1], [2]], column_headers=["a"], row_labels=["r", "r"], name="T")
        assert t.to_dict() == {"r": {"a": 2}}

    def test_from_dict(self):
        t = Table.from_dict(
            {"Revenue": {"North": 100, "South": 200}, "Costs": {"North": 40, "South": 60}},
            corner="Metric",
            name="T",
        )
        assert t.data.column_headers == ["North", "South"]
        assert t.data.row_labels == ["Revenue", "Costs"]
        assert t.data.corner == "Metric"
        assert t.data.rows == [[100, 200], [40, 60]]

    def test_from_dict_empty_raises(self):
        # no rows means no way to infer column_headers, which is mandatory
        with pytest.raises(ValueError):
            Table.from_dict({}, name="T")

    def test_from_dict_mismatched_keys_raises(self):
        with pytest.raises(ValueError):
            Table.from_dict({"r1": {"a": 1, "b": 2}, "r2": {"a": 3, "c": 4}}, name="T")

    def test_from_dict_different_key_order_raises(self):
        with pytest.raises(ValueError):
            Table.from_dict({"r1": {"a": 1, "b": 2}, "r2": {"b": 4, "a": 3}}, name="T")

    def test_from_dict_non_mapping_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table.from_dict([1, 2, 3], name="T")

    def test_from_dict_non_mapping_row_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table.from_dict({"r1": [1, 2]}, name="T")

    def test_to_dict_round_trips_through_from_dict(self, sample):
        rebuilt = Table.from_dict(sample.to_dict(), corner=sample.data.corner, name="T")
        assert rebuilt.data.rows == sample.data.rows
        assert rebuilt.data.column_headers == sample.data.column_headers
        assert rebuilt.data.row_labels == sample.data.row_labels


class TestReadPreservesTypesButCoercesLabels:
    def test_numeric_header_and_label_cells_are_coerced_to_str(self, book):
        Table(
            data=[[10, 20.5], [None, True]],
            column_headers=["a", "b"],
            row_labels=["r1", "r2"],
            name="T",
        ).create(book, sheet="Sheet")

        # Simulate a header/label cell that happens to hold a number on disk
        # (e.g. someone edited the sheet by hand).
        wb = load_workbook(book)
        ws = wb["Sheet"]
        ws["B2"] = 2024  # the "a" column header
        ws["A3"] = 1  # the "r1" row label
        wb.save(book)

        t = Table.read(book, "T")
        assert t.data.column_headers == ["2024", "b"]
        assert t.data.row_labels == ["1", "r2"]
        assert t.data.rows == [[10, 20.5], [None, True]]  # data keeps its type

    def test_typed_data_round_trips_through_table(self, book):
        t = Table(
            data=[[10, 2.5, dt.datetime(2026, 1, 1, 9, 0)]],
            column_headers=["a", "b", "c"],
            row_labels=["r1"],
            name="T",
        )
        t.create(book, sheet="Sheet")
        assert Table.read(book, "T") == t


class TestLabelAccess:
    def test_row_and_column(self, sample):
        assert sample.read_row("Revenue") == ["100", "200", "150"]
        assert sample.read_column("South") == ["200", "60"]

    def test_unknown_row_or_column_raises_keyerror(self, sample):
        with pytest.raises(KeyError):
            sample.read_row("Profit")
        with pytest.raises(KeyError):
            sample.read_column("West")


class TestCellByLabel:
    def test_reads_data(self, sample):
        assert sample.read_cell(row="Revenue", column="North") == "100"
        assert sample.read_cell(row="Costs", column="East") == "55"

    def test_unknown_label_raises_keyerror(self, sample):
        with pytest.raises(KeyError):
            sample.read_cell(row="Profit", column="North")
        with pytest.raises(KeyError):
            sample.read_cell(row="Revenue", column="West")

    def test_requires_both_row_and_column(self, sample):
        with pytest.raises(TypeError):
            sample.read_cell(row="Revenue")
        with pytest.raises(TypeError):
            sample.read_cell(column="North")

    def test_cannot_mix_position_and_label(self, sample):
        with pytest.raises(TypeError):
            sample.read_cell(row=2, column="North")
        with pytest.raises(TypeError):
            sample.read_cell(row="Revenue", column=2)


class TestCellByPosition:
    def test_data_cell_by_ref_and_by_rowcol(self, sample):
        assert sample.read_cell("B2") == "100"
        assert sample.read_cell(row=2, column=2) == "100"
        assert sample.read_cell("D3") == "55"

    def test_headers_and_labels_and_corner(self, sample):
        assert sample.read_cell(row=1, column=1) == "Metric"
        assert sample.read_cell(row=1, column=2) == "North"
        assert sample.read_cell(row=2, column=1) == "Revenue"

    def test_out_of_range_raises(self, sample):
        with pytest.raises(IndexError):
            sample.read_cell(row=99, column=1)

    def test_ref_and_rowcol_together_is_error(self, sample):
        with pytest.raises(TypeError):
            sample.read_cell("B2", row=2)

    def test_malformed_ref_raises_valueerror(self, sample):
        with pytest.raises(ValueError):
            sample.read_cell("not a ref")
        with pytest.raises(ValueError):
            sample.read_cell("")


class TestDunders:
    def test_equality_ignores_name(self, sample):
        same = Table(
            data=[["100", "200", "150"], ["40", "60", "55"]],
            column_headers=["North", "South", "East"],
            row_labels=["Revenue", "Costs"],
            corner="Metric",
            name="DifferentName",
        )
        assert sample == same
        assert sample != Table(data=[["1"]], column_headers=["a"], row_labels=["x"], name="T")


class TestReadWriteRoundTrip:
    def test_write_then_read_reproduces_the_table(self, book, sample):
        sample.create(book, sheet="Sheet")
        assert Table.read(book, sample.name) == sample

    def test_assembled_grid_layout_on_disk(self, book, sample):
        sample.create(book, sheet="Sheet")
        assert read_sheet(book) == [
            ["TABLE NAME", "Sample"],
            ["Metric", "North", "South", "East"],
            ["Revenue", "100", "200", "150"],
            ["Costs", "40", "60", "55"],
        ]

    def test_read_allows_duplicate_headers(self, book):
        Table(
            data=[[10, 20]], column_headers=["amount", "amount"], row_labels=["Jan"], name="Dup"
        ).create(book, sheet="Sheet")
        t = Table.read(book, "Dup")
        assert t.data.column_headers == ["amount", "amount"]

    def test_create_with_headers_but_no_rows_round_trips(self, book):
        # regression: the corner cell used to be omitted from the persisted
        # header row whenever row_labels was empty, silently shifting every
        # header over by one on read-back
        Table(data=[], column_headers=["a", "b", "c"], name="Empty").create(book, sheet="Sheet")
        t = Table.read(book, "Empty")
        assert t.data.column_headers == ["a", "b", "c"]
        assert t.data.corner == ""
        assert t.data.rows == []

    def test_create_with_headers_then_add_row_round_trips(self, book):
        t = Table(data=[], column_headers=["a", "b"], name="T")
        t.create(book, sheet="Sheet")
        t.add_row("r1", [1, 2])
        t.write(book)
        assert Table.read(book, "T").data.rows == [[1, 2]]


class TestSetCell:
    def test_set_by_label(self, sample):
        sample.set_cell(row="Revenue", column="North", value=120)
        assert sample.read_cell(row="Revenue", column="North") == 120

    def test_set_unknown_label_raises(self, sample):
        with pytest.raises(KeyError):
            sample.set_cell(row="Profit", column="North", value=1)

    def test_set_by_position_via_ref(self, sample):
        sample.set_cell("B2", value=999)
        assert sample.read_cell(row="Revenue", column="North") == 999

    def test_set_by_position_via_rowcol(self, sample):
        sample.set_cell(row=2, column=2, value=999)
        assert sample.read_cell(row="Revenue", column="North") == 999

    def test_cannot_set_header_by_position(self, sample):
        with pytest.raises(ValueError):
            sample.set_cell(row=1, column=2, value="N")

    def test_cannot_set_row_label_by_position(self, sample):
        with pytest.raises(ValueError):
            sample.set_cell(row=2, column=1, value="Rev")

    def test_cannot_set_corner_by_position(self, sample):
        with pytest.raises(ValueError):
            sample.set_cell(row=1, column=1, value="M")

    def test_out_of_range_raises(self, sample):
        with pytest.raises(IndexError):
            sample.set_cell(row=50, column=1, value=1)

    def test_requires_both_row_and_column(self, sample):
        with pytest.raises(TypeError):
            sample.set_cell(row="Revenue", value=1)

    def test_cannot_mix_position_and_label(self, sample):
        with pytest.raises(TypeError):
            sample.set_cell(row=2, column="North", value=1)
        with pytest.raises(TypeError):
            sample.set_cell(row="Revenue", column=2, value=1)

    def test_ref_and_rowcol_together_is_error(self, sample):
        with pytest.raises(TypeError):
            sample.set_cell("B2", row=2, column=2, value=1)


class TestReplaceLine:
    def test_set_row(self, sample):
        sample.set_row("Revenue", [1, 2, 3])
        assert sample.read_row("Revenue") == [1, 2, 3]

    def test_set_row_wrong_length_raises(self, sample):
        with pytest.raises(ValueError):
            sample.set_row("Revenue", [1, 2])

    def test_set_column(self, sample):
        sample.set_column("North", ["x", "y"])
        assert sample.read_column("North") == ["x", "y"]

    def test_set_column_wrong_length_raises(self, sample):
        with pytest.raises(ValueError):
            sample.set_column("North", ["x"])


class TestAdd:
    def test_add_row(self, sample):
        sample.add_row("Profit", [60, 140, 95])
        assert sample.data.row_labels == ["Revenue", "Costs", "Profit"]
        assert sample.read_row("Profit") == [60, 140, 95]
        assert len(sample.data.rows) == 3

    def test_add_row_wrong_length_raises(self, sample):
        with pytest.raises(ValueError):
            sample.add_row("Profit", [1, 2])

    def test_add_row_non_string_label_raises(self, sample):
        with pytest.raises(TypeError):
            sample.add_row(123, [1, 2, 3])

    def test_add_row_duplicate_label_raises(self, sample):
        with pytest.raises(ValueError):
            sample.add_row("Revenue", [1, 2, 3])

    def test_add_row_empty_label_raises(self, sample):
        with pytest.raises(ValueError):
            sample.add_row("", [1, 2, 3])

    def test_add_column(self, sample):
        sample.add_column("West", [10, 20])
        assert sample.data.column_headers == ["North", "South", "East", "West"]
        assert sample.read_column("West") == [10, 20]
        assert sample.read_row("Revenue") == ["100", "200", "150", 10]

    def test_add_column_wrong_length_raises(self, sample):
        with pytest.raises(ValueError):
            sample.add_column("West", [10])

    def test_add_column_non_string_header_raises(self, sample):
        with pytest.raises(TypeError):
            sample.add_column(123, [10, 20])

    def test_add_column_duplicate_header_raises(self, sample):
        with pytest.raises(ValueError):
            sample.add_column("North", [10, 20])

    def test_add_column_empty_header_raises(self, sample):
        with pytest.raises(ValueError):
            sample.add_column("", [10, 20])

    def test_insert_row_at_start(self, sample):
        sample.insert_row(1, "Forecast", [1, 2, 3])
        assert sample.data.row_labels == ["Forecast", "Revenue", "Costs"]
        assert sample.data.rows == [[1, 2, 3], ["100", "200", "150"], ["40", "60", "55"]]

    def test_insert_row_in_middle(self, sample):
        sample.insert_row(2, "Forecast", [1, 2, 3])
        assert sample.data.row_labels == ["Revenue", "Forecast", "Costs"]

    def test_insert_row_at_end_matches_add_row(self, sample):
        sample.insert_row(3, "Profit", [60, 140, 95])
        assert sample.data.row_labels == ["Revenue", "Costs", "Profit"]

    def test_insert_row_out_of_range_raises(self, sample):
        with pytest.raises(IndexError):
            sample.insert_row(0, "Forecast", [1, 2, 3])
        with pytest.raises(IndexError):
            sample.insert_row(4, "Forecast", [1, 2, 3])

    def test_insert_row_wrong_length_raises(self, sample):
        with pytest.raises(ValueError):
            sample.insert_row(1, "Forecast", [1, 2])

    def test_insert_row_duplicate_label_raises(self, sample):
        with pytest.raises(ValueError):
            sample.insert_row(1, "Revenue", [1, 2, 3])

    def test_insert_row_non_string_label_raises(self, sample):
        with pytest.raises(TypeError):
            sample.insert_row(1, 123, [1, 2, 3])

    def test_insert_row_non_int_position_raises_typeerror(self, sample):
        with pytest.raises(TypeError):
            sample.insert_row("1", "Forecast", [1, 2, 3])
        with pytest.raises(TypeError):
            sample.insert_row(1.5, "Forecast", [1, 2, 3])

    def test_insert_column_at_start(self, sample):
        sample.insert_column(1, "West", [10, 20])
        assert sample.data.column_headers == ["West", "North", "South", "East"]
        assert sample.read_row("Revenue") == [10, "100", "200", "150"]

    def test_insert_column_in_middle(self, sample):
        sample.insert_column(2, "West", [10, 20])
        assert sample.data.column_headers == ["North", "West", "South", "East"]

    def test_insert_column_at_end_matches_add_column(self, sample):
        sample.insert_column(4, "West", [10, 20])
        assert sample.data.column_headers == ["North", "South", "East", "West"]

    def test_insert_column_out_of_range_raises(self, sample):
        with pytest.raises(IndexError):
            sample.insert_column(0, "West", [10, 20])
        with pytest.raises(IndexError):
            sample.insert_column(5, "West", [10, 20])

    def test_insert_column_wrong_length_raises(self, sample):
        with pytest.raises(ValueError):
            sample.insert_column(1, "West", [10])

    def test_insert_column_duplicate_header_raises(self, sample):
        with pytest.raises(ValueError):
            sample.insert_column(1, "North", [10, 20])

    def test_insert_column_non_string_header_raises(self, sample):
        with pytest.raises(TypeError):
            sample.insert_column(1, 123, [10, 20])

    def test_insert_column_non_int_position_raises_typeerror(self, sample):
        with pytest.raises(TypeError):
            sample.insert_column("1", "West", [10, 20])
        with pytest.raises(TypeError):
            sample.insert_column(1.5, "West", [10, 20])

    def test_build_table_from_empty(self):
        t = Table(data=[], column_headers=["a", "b"], name="T")
        t.add_row("x", [1, 2])
        t.add_row("y", [3, 4])
        assert t.data.rows == [[1, 2], [3, 4]]
        assert t.data.row_labels == ["x", "y"]


class TestRemoveAndRename:
    def test_drop_row(self, sample):
        sample.drop_row("Revenue")
        assert sample.data.row_labels == ["Costs"]
        assert sample.data.rows == [["40", "60", "55"]]

    def test_drop_column(self, sample):
        sample.drop_column("South")
        assert sample.data.column_headers == ["North", "East"]
        assert sample.read_row("Revenue") == ["100", "150"]

    def test_drop_unknown_raises(self, sample):
        with pytest.raises(KeyError):
            sample.drop_row("Profit")

    def test_rename_row_and_column(self, sample):
        sample.rename_row("Revenue", "Sales")
        sample.rename_column("North", "N")
        assert sample.read_cell(row="Sales", column="N") == "100"

    def test_rename_unknown_raises(self, sample):
        with pytest.raises(KeyError):
            sample.rename_column("West", "W")

    def test_rename_to_non_string_raises(self, sample):
        with pytest.raises(TypeError):
            sample.rename_row("Revenue", 123)
        with pytest.raises(TypeError):
            sample.rename_column("North", 123)

    def test_rename_to_empty_raises(self, sample):
        with pytest.raises(ValueError):
            sample.rename_row("Revenue", "")
        with pytest.raises(ValueError):
            sample.rename_column("North", "")

    def test_rename_to_existing_label_raises(self, sample):
        with pytest.raises(ValueError):
            sample.rename_row("Revenue", "Costs")
        with pytest.raises(ValueError):
            sample.rename_column("North", "South")

    def test_rename_to_same_name_is_allowed(self, sample):
        sample.rename_row("Revenue", "Revenue")
        assert sample.data.row_labels == ["Revenue", "Costs"]

    def test_rename_unknown_old_raises_keyerror_not_valueerror(self, sample):
        with pytest.raises(KeyError):
            sample.rename_row("Ghost", "Costs")


class TestMutationKeepsInvariants:
    def test_edits_survive_a_round_trip(self, book, sample):
        sample.create(book, sheet="Sheet")

        # strings only, so the round trip is exact (read_sheet coerces to str)
        sample.set_cell(row="Revenue", column="North", value="111")
        sample.add_row("Profit", ["1", "2", "3"])
        sample.drop_column("South")
        sample.rename_row("Costs", "Expenses")

        sample.write(book)
        assert Table.read(book, sample.name) == sample
