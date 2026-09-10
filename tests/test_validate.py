"""Tests for validate.py."""

from __future__ import annotations

import datetime as dt
from decimal import Decimal

import pytest
from openpyxl import Workbook

from pyhandlexl.errors import CellTypeError, DimensionError, SheetNameError
from pyhandlexl.validate import (
    MAX_COLUMNS,
    MAX_ROWS,
    check_cell_value,
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


class TestCheckCellValue:
    @pytest.mark.parametrize(
        "value",
        [
            "text",
            "",
            0,
            42,
            -1,
            3.14,
            5.0,
            True,
            False,
            None,
            dt.datetime(2026, 1, 1, 12, 0),
            dt.date(2026, 1, 1),
            dt.time(9, 30),
            dt.timedelta(hours=2),
        ],
    )
    def test_supported_values_pass(self, value):
        check_cell_value(value)  # no raise

    @pytest.mark.parametrize(
        "value",
        [[1, 2], {"a": 1}, {1, 2}, (1, 2), b"bytes", 1 + 2j, Decimal("1.5"), object()],
    )
    def test_unsupported_values_raise(self, value):
        with pytest.raises(CellTypeError):
            check_cell_value(value)

    def test_tz_aware_datetime_raises(self):
        with pytest.raises(CellTypeError):
            check_cell_value(dt.datetime(2026, 1, 1, tzinfo=dt.timezone.utc))


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
