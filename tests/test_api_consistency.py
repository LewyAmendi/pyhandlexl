"""Small consistency points from the pre-1.0 API audit."""

from __future__ import annotations

import inspect

import pytest

from pyhandlexl import (
    SheetNotFoundError,
    Table,
    TableData,
    TableNotFoundError,
    TableStyle,
    create_sheet,
    delete_sheet,
    read_sheet,
    rename_sheet,
)


@pytest.fixture
def t() -> Table:
    return Table(data=[[1, 2], [3, 4]], row_labels=["a", "b"], column_headers=["x", "y"], name="T")


class TestOneNameForALabelAndAHeader:
    def test_every_row_method_calls_the_label_label(self, t):
        assert t.read_row(label="a") == [1, 2]
        t.set_row(label="a", values=[5, 6])
        t.add_row(label="c", values=[7, 8])
        t.insert_row(position=1, label="z", values=[0, 0])
        t.drop_row(label="z")
        assert t.data.rows == [[5, 6], [3, 4], [7, 8]]

    def test_every_column_method_calls_the_header_header(self, t):
        assert t.read_column(header="x") == [1, 3]
        t.set_column(header="x", values=[9, 9])
        t.add_column(header="w", values=[5, 5])
        t.insert_column(position=1, header="v", values=[0, 0])
        t.drop_column(header="v")
        assert t.data.columns == [[9, 9], [2, 4], [5, 5]]

    def test_no_method_still_uses_the_longer_names(self):
        names = {
            parameter
            for member in vars(Table).values()
            if inspect.isfunction(member)
            for parameter in inspect.signature(member).parameters
        }
        assert "row_label" not in names and "column_header" not in names


class TestFromDictTakesAStyle:
    def test_the_style_is_used(self):
        t = Table.from_dict({"a": {"x": 1}}, name="T", style=TableStyle.MINIMAL)
        assert t.style == TableStyle.MINIMAL

    def test_the_default_is_the_default_style(self):
        assert Table.from_dict({"a": {"x": 1}}, name="T").style == TableStyle.DEFAULT


class TestErrorsReadAsSentences:
    def test_a_missing_table_is_not_wrapped_in_quotes(self, book):
        with pytest.raises(TableNotFoundError) as caught:
            Table.read(book, "Ghost")
        assert str(caught.value) == "no table named 'Ghost'"

    @pytest.mark.parametrize(
        "call",
        [
            lambda book: read_sheet(book, "Ghost"),
            lambda book: delete_sheet(book, "Ghost"),
            lambda book: rename_sheet(book, "Ghost", "Other"),
            lambda book: Table(column_headers=["a"], name="T").create(book, sheet="Ghost"),
        ],
        ids=["read_sheet", "delete_sheet", "rename_sheet", "Table.create"],
    )
    def test_a_missing_sheet_says_which(self, book, call):
        with pytest.raises(SheetNotFoundError) as caught:
            call(book)
        assert str(caught.value) == "no worksheet named 'Ghost'"

    def test_they_are_still_key_errors(self, book):
        with pytest.raises(KeyError):
            read_sheet(book, "Ghost")
        with pytest.raises(KeyError):
            Table.read(book, "Ghost")

    def test_other_sheet_functions_are_unaffected(self, book):
        create_sheet(book, "Real")
        assert read_sheet(book, "Real") == []


class TestTableDataRepr:
    def test_a_small_snapshot_shows_its_contents(self, t):
        text = repr(t.data)
        assert "rows=[[1, 2], [3, 4]]" in text and "row_labels=['a', 'b']" in text

    def test_a_large_snapshot_shows_only_its_shape(self):
        big = Table(
            data=[[i, i] for i in range(5000)],
            row_labels=[str(i) for i in range(5000)],
            column_headers=["a", "b"],
            name="T",
        )
        text = repr(big.data)
        assert len(text) < 300
        assert "5000 rows x 2 columns" in text

    def test_equality_still_compares_the_contents(self, t):
        assert t.data == t.data
        assert isinstance(t.data, TableData)
        other = Table(
            data=[[1, 2], [3, 5]], row_labels=["a", "b"], column_headers=["x", "y"], name="T"
        )
        assert t.data != other.data
