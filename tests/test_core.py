"""Tests for core.py: create_workbook, read_sheet, write_sheet, append_rows, sheet mgmt."""

from __future__ import annotations

import pytest
from openpyxl import load_workbook

from pyhandlexl.core import (
    append_rows,
    create_sheet,
    create_workbook,
    delete_sheet,
    delete_workbook,
    list_sheets,
    read_sheet,
    rename_sheet,
    sheet_exists,
    write_sheet,
)
from pyhandlexl.errors import DimensionError, SheetNameError, SheetNotFoundError


class TestCreateWorkbook:
    def test_creates_empty_workbook(self, tmp_path):
        path = tmp_path / "new.xlsx"
        create_workbook(path)
        assert list_sheets(path) == ["Sheet"]
        assert read_sheet(path) == []

    def test_custom_sheet_name(self, tmp_path):
        path = tmp_path / "new.xlsx"
        create_workbook(path, sheet="Data")
        assert list_sheets(path) == ["Data"]

    def test_existing_path_raises(self, book):
        with pytest.raises(FileExistsError):
            create_workbook(book)

    def test_invalid_sheet_name_raises(self, tmp_path):
        with pytest.raises(SheetNameError):
            create_workbook(tmp_path / "new.xlsx", sheet="bad/name")


class TestDeleteWorkbook:
    def test_deletes_the_file(self, book):
        delete_workbook(book)
        assert not book.exists()

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            delete_workbook(tmp_path / "nope.xlsx")

    def test_non_workbook_path_raises(self, tmp_path):
        txt = tmp_path / "notes.txt"
        txt.write_text("keep me")
        with pytest.raises(ValueError):
            delete_workbook(txt)
        assert txt.exists()


class TestFileMustExist:
    def test_write_sheet_raises_when_absent(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            write_sheet(tmp_path / "nope.xlsx", [["a"]])

    def test_append_rows_raises_when_absent(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            append_rows(tmp_path / "nope.xlsx", [["a"]])

    def test_create_sheet_raises_when_absent(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            create_sheet(tmp_path / "nope.xlsx", "Data")

    def test_read_sheet_raises_when_absent(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            read_sheet(tmp_path / "nope.xlsx")


class TestWriteThenRead:
    def test_round_trip(self, book):
        write_sheet(book, [["a", "b"], ["c", "d"]])
        assert read_sheet(book) == [["a", "b"], ["c", "d"]]

    def test_values_are_returned_as_strings(self, book):
        write_sheet(book, [[1, 2.5, True, None]])
        assert read_sheet(book) == [["1", "2.5", "True"]]  # trailing None trimmed

    def test_write_replaces_existing_content(self, book):
        write_sheet(book, [["old", "old", "old"], ["old", "old", "old"]])
        write_sheet(book, [["new"]])
        assert read_sheet(book) == [["new"]]

    def test_writing_a_new_sheet_keeps_the_default_one(self, book):
        write_sheet(book, [["x"]], sheet="Data")
        assert load_workbook(book).sheetnames == ["Sheet", "Data"]

    def test_other_sheets_are_preserved(self, book):
        wb = load_workbook(book)
        wb.active.title = "Keep"
        wb.active["A1"] = "untouched"
        wb.create_sheet("Target")
        wb.save(book)

        write_sheet(book, [["changed"]], sheet="Target")

        result = load_workbook(book)
        assert result["Keep"]["A1"].value == "untouched"
        assert read_sheet(book, "Target") == [["changed"]]


class TestReadSheet:
    def test_unknown_sheet_raises(self, book):
        write_sheet(book, [["a"]])
        with pytest.raises(SheetNotFoundError):
            read_sheet(book, "Ghost")

    def test_trailing_empty_cells_are_trimmed(self, book):
        write_sheet(book, [["a", "", ""], ["b", "c", ""]])
        assert read_sheet(book) == [["a"], ["b", "c"]]

    def test_pad_makes_result_rectangular(self, book):
        write_sheet(book, [["a"], ["b", "c", "d"]])
        assert read_sheet(book, pad=True) == [["a", "", ""], ["b", "c", "d"]]


class TestOrientation:
    def test_columns_orientation_transposes(self, book):
        write_sheet(book, [["h1", "h2"], ["v1", "v2"]], orientation="columns")
        assert read_sheet(book) == [["h1", "v1"], ["h2", "v2"]]

    def test_invalid_orientation_raises(self, book):
        with pytest.raises(ValueError):
            write_sheet(book, [["a"]], orientation="sideways")


class TestAppendRows:
    def test_appends_after_existing_rows(self, book):
        write_sheet(book, [["a"], ["b"]])
        append_rows(book, [["c"], ["d"]])
        assert read_sheet(book) == [["a"], ["b"], ["c"], ["d"]]

    def test_empty_rows_is_a_noop(self, book):
        write_sheet(book, [["a"]])
        append_rows(book, [])
        assert read_sheet(book) == [["a"]]

    def test_appends_to_an_empty_file(self, book):
        append_rows(book, [["a"], ["b"]])
        assert read_sheet(book) == [["a"], ["b"]]


class TestDimensionGuard:
    def test_too_many_columns_raises(self, book):
        with pytest.raises(DimensionError):
            write_sheet(book, [[""] * 16_385])


class TestSheetManagement:
    def test_list_sheets(self, book):
        wb = load_workbook(book)
        wb.active.title = "One"
        wb.create_sheet("Two")
        wb.save(book)
        assert list_sheets(book) == ["One", "Two"]

    def test_list_sheets_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            list_sheets(tmp_path / "nope.xlsx")

    def test_sheet_exists(self, book):
        create_sheet(book, "Data")
        assert sheet_exists(book, "Data") is True
        assert sheet_exists(book, "Missing") is False

    def test_create_sheet_adds_to_existing(self, book):
        create_sheet(book, "Extra")
        assert list_sheets(book) == ["Sheet", "Extra"]

    def test_create_sheet_duplicate_raises(self, book):
        create_sheet(book, "Data")
        with pytest.raises(ValueError):
            create_sheet(book, "Data")

    def test_create_sheet_invalid_name_raises(self, book):
        with pytest.raises(SheetNameError):
            create_sheet(book, "bad/name")

    def test_delete_sheet(self, book):
        create_sheet(book, "Gone")
        delete_sheet(book, "Gone")
        assert sheet_exists(book, "Gone") is False

    def test_delete_missing_sheet_raises(self, book):
        with pytest.raises(SheetNotFoundError):
            delete_sheet(book, "Ghost")

    def test_delete_last_sheet_raises(self, book):
        with pytest.raises(ValueError):
            delete_sheet(book, list_sheets(book)[0])

    def test_rename_sheet(self, book):
        write_sheet(book, [["a"]], sheet="Sheet")
        rename_sheet(book, "Sheet", "New")
        assert list_sheets(book) == ["New"]
        assert read_sheet(book, "New") == [["a"]]

    def test_rename_missing_sheet_raises(self, book):
        with pytest.raises(SheetNotFoundError):
            rename_sheet(book, "Nope", "New")

    def test_rename_to_existing_name_raises(self, book):
        create_sheet(book, "B")
        with pytest.raises(ValueError):
            rename_sheet(book, "Sheet", "B")

    def test_rename_invalid_name_raises(self, book):
        with pytest.raises(SheetNameError):
            rename_sheet(book, "Sheet", "x" * 32)
