"""Tests for validate.py."""

from __future__ import annotations

import pytest
from openpyxl import Workbook

from pyhandlexl.errors import DimensionError, SheetNameError
from pyhandlexl.validate import (
    MAX_COLUMNS,
    MAX_ROWS,
    check_dimensions,
    check_sheet_name,
    is_valid_xlsx,
)


class TestCheckDimensions:
    def test_at_the_limit_is_ok(self):
        check_dimensions(MAX_ROWS, MAX_COLUMNS)  # no raise

    def test_too_many_rows(self):
        with pytest.raises(DimensionError):
            check_dimensions(MAX_ROWS + 1, 1)

    def test_too_many_columns(self):
        with pytest.raises(DimensionError):
            check_dimensions(1, MAX_COLUMNS + 1)


class TestCheckSheetName:
    def test_normal_name_is_ok(self):
        check_sheet_name("Sales Q1")  # no raise

    def test_exactly_31_chars_is_ok(self):
        check_sheet_name("f" * 31)  # no raise

    def test_32_chars_raises(self):
        with pytest.raises(SheetNameError):
            check_sheet_name("f" * 32)

    def test_empty_raises(self):
        with pytest.raises(SheetNameError):
            check_sheet_name("")

    def test_non_string_raises(self):
        with pytest.raises(SheetNameError):
            check_sheet_name(None)

    @pytest.mark.parametrize("bad", list(r"\/?*[]:"))
    def test_illegal_character_raises(self, bad):
        with pytest.raises(SheetNameError):
            check_sheet_name(f"Sheet{bad}1")

    @pytest.mark.parametrize("name", ["History", "history", "HISTORY"])
    def test_reserved_history_raises(self, name):
        with pytest.raises(SheetNameError):
            check_sheet_name(name)


class TestIsValidXlsx:
    def test_real_workbook_is_valid(self, tmp_path):
        path = tmp_path / "good.xlsx"
        Workbook().save(path)
        assert is_valid_xlsx(path) is True

    def test_accepts_str_path(self, tmp_path):
        path = tmp_path / "good.xlsx"
        Workbook().save(path)
        assert is_valid_xlsx(str(path)) is True

    def test_missing_file_is_invalid(self, tmp_path):
        assert is_valid_xlsx(tmp_path / "nope.xlsx") is False

    def test_directory_is_invalid(self, tmp_path):
        assert is_valid_xlsx(tmp_path) is False

    def test_non_zip_content_is_invalid(self, tmp_path):
        path = tmp_path / "fake.xlsx"
        path.write_text("this is not an excel file")
        assert is_valid_xlsx(path) is False

    def test_zip_missing_required_parts_is_invalid(self, tmp_path):
        import zipfile

        path = tmp_path / "empty.xlsx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("hello.txt", "not a workbook")
        assert is_valid_xlsx(path) is False
