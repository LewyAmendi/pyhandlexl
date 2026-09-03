"""Tests for the Table class."""

from __future__ import annotations

import pytest

from pyhandlexl.core import read_sheet
from pyhandlexl.table import Table


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
        assert t.data == [["1", "2"], ["3", "4"]]
        assert t.column_headers == []
        assert t.row_labels == []


class TestProperties:
    def test_parts(self, sample):
        assert sample.corner == "Metric"
        assert sample.column_headers == ["North", "South", "East"]
        assert sample.row_labels == ["Revenue", "Costs"]
        assert sample.data == [["100", "200", "150"], ["40", "60", "55"]]

    def test_properties_return_copies(self, sample):
        sample.data.append(["999"])
        sample.column_headers.append("West")
        assert len(sample.data) == 2
        assert sample.column_headers == ["North", "South", "East"]


class TestLabelAccess:
    def test_at(self, sample):
        assert sample.at("Revenue", "North") == "100"
        assert sample.at("Costs", "East") == "55"

    def test_row_and_column(self, sample):
        assert sample.row("Revenue") == ["100", "200", "150"]
        assert sample.column("South") == ["200", "60"]

    def test_unknown_label_raises_keyerror(self, sample):
        with pytest.raises(KeyError):
            sample.at("Profit", "North")
        with pytest.raises(KeyError):
            sample.column("West")


class TestPositionalAccess:
    def test_data_cell_by_ref_and_by_rowcol(self, sample):
        assert sample.cell("B2") == "100"
        assert sample.cell(row=2, column=2) == "100"
        assert sample.cell(row=3, column="D") == "55"

    def test_headers_and_labels_and_corner(self, sample):
        assert sample.cell(row=1, column=1) == "Metric"
        assert sample.cell(row=1, column=2) == "North"
        assert sample.cell(row=2, column=1) == "Revenue"

    def test_out_of_range_raises(self, sample):
        with pytest.raises(IndexError):
            sample.cell(row=99, column=1)

    def test_ref_and_rowcol_together_is_error(self, sample):
        with pytest.raises(TypeError):
            sample.cell("B2", row=2)


class TestDunders:
    def test_len_iter_getitem_contains(self, sample):
        assert len(sample) == 2
        assert list(sample) == [["100", "200", "150"], ["40", "60", "55"]]
        assert sample["Costs"] == ["40", "60", "55"]
        assert "Revenue" in sample
        assert "Profit" not in sample

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
        assert t.column_headers == ["a", "b"]
        assert t.row_labels == []
        assert t.data == [["1", "2"]]

    def test_read_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            Table.read(tmp_path / "nope.xlsx")
