"""The SchemaRebuiltWarning: where it points, and that every operation keeps the rebuild."""

from __future__ import annotations

import warnings
from pathlib import Path

import pytest
from openpyxl import load_workbook

from pyhandlexl import (
    SchemaRebuiltWarning,
    Table,
    append_rows,
    clear_all_sheet_data,
    create_sheet,
    delete_sheet,
    delete_table,
    import_csv_to_xl,
    list_tables,
    rename_sheet,
    sheet_kind,
    table_info,
    write_sheet,
)
from pyhandlexl._multi_table import SCHEMA_SHEET


@pytest.fixture
def path(book):
    """A workbook with one table on 'Data' and plain grid data on 'Grid'."""
    create_sheet(book, "Data")
    create_sheet(book, "Grid")
    write_sheet(book, [["g"]], sheet="Grid")
    Table(data=[[1, 2]], row_labels=["r"], column_headers=["a", "b"], name="T").create(
        book, sheet="Data"
    )
    Table(data=[[5]], row_labels=["k"], column_headers=["v"], name="U").create(book, sheet="Data")
    return book


def _drop_schema(path) -> None:
    wb = load_workbook(path)
    del wb[SCHEMA_SHEET]
    wb.save(path)


def _has_schema(path) -> bool:
    return SCHEMA_SHEET in load_workbook(path).sheetnames


def _import_into_grid(path):
    csv_path = Path(path).with_suffix(".csv")
    csv_path.write_text("a,b\n", encoding="utf-8")
    import_csv_to_xl(csv_path, path, sheet="Grid")


# Each operation is written as a lambda HERE, in the test file, so "the caller's line"
# is a line in this file — which is exactly what the warning must point at.
OPERATIONS = {
    "Table.read": lambda p: Table.read(p, "T"),
    "list_tables": lambda p: list_tables(p),
    "table_info": lambda p: table_info(p, "T"),
    "delete_table": lambda p: delete_table(p, "U"),
    "sheet_kind": lambda p: sheet_kind(p, "Data"),
    "delete_sheet": lambda p: delete_sheet(p, "Grid"),
    "rename_sheet": lambda p: rename_sheet(p, "Grid", "Grid2"),
    "clear_all_sheet_data": lambda p: clear_all_sheet_data(p, "Grid"),
    "write_sheet": lambda p: write_sheet(p, [["z"]], sheet="Grid"),
    "import_csv_to_xl": lambda p: _import_into_grid(p),
    "append_rows": lambda p: append_rows(p, [["z"]], sheet="Grid"),
    "Table.create": lambda p: Table(column_headers=["a"], name="NEW").create(p, sheet="Data"),
}


class TestTheWarningPointsAtTheCallersLine:
    @pytest.mark.parametrize("name", list(OPERATIONS))
    def test_every_operation_reports_the_users_own_line(self, name, path):
        _drop_schema(path)
        with pytest.warns(SchemaRebuiltWarning) as caught:
            OPERATIONS[name](path)
        assert len(caught) == 1
        assert caught[0].filename == __file__

    def test_table_write_reports_the_users_own_line(self, path):
        table = Table.read(path, "T")
        table.add_row("r2", [3, 4])
        _drop_schema(path)
        with pytest.warns(SchemaRebuiltWarning) as caught:
            table.write(path)
        assert [w.filename for w in caught] == [__file__]

    def test_the_duplicate_name_warning_points_there_too(self, path):
        wb = load_workbook(path)
        ws = wb["Grid"]
        ws.delete_rows(1, ws.max_row)
        ws["A1"], ws["B1"] = "TABLE NAME", "T"  # a second marker claiming the name "T"
        ws["A2"], ws["B2"], ws["A3"], ws["B3"] = "corner", "y", "s", 2
        del wb[SCHEMA_SHEET]
        wb.save(path)
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            list_tables(path)
        rebuilt = [w for w in caught if issubclass(w.category, SchemaRebuiltWarning)]
        assert len(rebuilt) == 2  # the duplicate-name warning and the rebuild summary
        assert {w.filename for w in rebuilt} == {__file__}
        assert any("more than one location" in str(w.message) for w in rebuilt)


class TestEveryOperationKeepsTheRebuild:
    @pytest.mark.parametrize("name", list(OPERATIONS))
    def test_the_schema_sheet_is_back_in_the_file_afterwards(self, name, path):
        _drop_schema(path)
        with pytest.warns(SchemaRebuiltWarning):
            OPERATIONS[name](path)
        assert _has_schema(path)

    @pytest.mark.parametrize("name", ["sheet_kind", "write_sheet", "append_rows"])
    def test_a_second_call_neither_warns_nor_rescans(self, name, path):
        _drop_schema(path)
        with pytest.warns(SchemaRebuiltWarning):
            OPERATIONS[name](path)
        with warnings.catch_warnings():
            warnings.simplefilter("error")  # any SchemaRebuiltWarning would raise
            OPERATIONS[name](path)

    def test_write_sheet_still_replaces_the_grid_and_leaves_the_tables_intact(self, path):
        _drop_schema(path)
        with pytest.warns(SchemaRebuiltWarning):
            write_sheet(path, [["new", 1]], sheet="Grid")
        assert sorted(list_tables(path)) == ["T", "U"]
        assert Table.read(path, "T").read_row("r") == [1, 2]

    def test_append_rows_still_appends_and_leaves_the_tables_intact(self, path):
        _drop_schema(path)
        with pytest.warns(SchemaRebuiltWarning):
            append_rows(path, [["extra"]], sheet="Grid")
        assert sorted(list_tables(path)) == ["T", "U"]

    def test_sheet_kind_still_answers_correctly(self, path):
        _drop_schema(path)
        with pytest.warns(SchemaRebuiltWarning):
            assert sheet_kind(path, "Data") == "table"
        assert sheet_kind(path, "Grid") == "grid"


class TestNothingToRebuild:
    def test_a_workbook_without_tables_never_warns_or_gains_a_schema(self, book):
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            write_sheet(book, [["x"]])
            append_rows(book, [["y"]])
            sheet_kind(book, "Sheet")
            list_tables(book)
        assert not _has_schema(book)
