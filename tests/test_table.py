"""Tests for the Table class."""

from __future__ import annotations

import pytest

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
    )


class TestConstruction:
    def test_mismatched_row_labels_raise(self):
        with pytest.raises(ValueError):
            Table(data=[["1"]], column_headers=["a"], row_labels=["x", "y"])

    def test_mismatched_row_width_raises(self):
        with pytest.raises(ValueError):
            Table(data=[["1", "2"]], column_headers=["a"], row_labels=["x"])

    def test_grid_only_is_allowed(self):
        t = Table(data=[["1", "2"], ["3", "4"]])
        assert t.data.rows == [["1", "2"], ["3", "4"]]
        assert t.data.column_headers == []
        assert t.data.row_labels == []

    def test_non_string_row_label_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]], column_headers=["a"], row_labels=[1])

    def test_non_string_column_header_raises_typeerror(self):
        with pytest.raises(TypeError):
            Table(data=[["1"]], column_headers=[1], row_labels=["x"])


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
        sample.corner = "Quarter"
        assert sample.corner == "Quarter"
        assert sample.data.corner == "Quarter"


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


class TestDunders:
    def test_equality(self, sample):
        same = Table(
            data=[["100", "200", "150"], ["40", "60", "55"]],
            column_headers=["North", "South", "East"],
            row_labels=["Revenue", "Costs"],
            corner="Metric",
        )
        assert sample == same
        assert sample != Table(data=[["1"]])


class TestReadWriteRoundTrip:
    def test_write_then_read_reproduces_the_table(self, tmp_path, sample):
        path = tmp_path / "t.xlsx"
        sample.write(path)
        assert Table.read(path) == sample

    def test_assembled_grid_layout_on_disk(self, tmp_path, sample):
        path = tmp_path / "t.xlsx"
        sample.write(path)
        assert read_sheet(path) == [
            ["Metric", "North", "South", "East"],
            ["Revenue", "100", "200", "150"],
            ["Costs", "40", "60", "55"],
        ]

    def test_read_without_row_labels(self, tmp_path):
        path = tmp_path / "t.xlsx"
        Table(
            data=[["1", "2"]],
            column_headers=["a", "b"],
        ).write(path)
        t = Table.read(path, row_labels=False)
        assert t.data.column_headers == ["a", "b"]
        assert t.data.row_labels == []
        assert t.data.rows == [["1", "2"]]

    def test_read_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            Table.read(tmp_path / "nope.xlsx")


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

    def test_build_table_from_empty(self):
        t = Table(data=[], column_headers=["a", "b"])
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


class TestMutationKeepsInvariants:
    def test_edits_survive_a_round_trip(self, tmp_path, sample):
        # strings only, so the round trip is exact (read_sheet coerces to str)
        sample.set_cell(row="Revenue", column="North", value="111")
        sample.add_row("Profit", ["1", "2", "3"])
        sample.drop_column("South")
        sample.rename_row("Costs", "Expenses")

        path = tmp_path / "edited.xlsx"
        sample.write(path)
        assert Table.read(path) == sample
