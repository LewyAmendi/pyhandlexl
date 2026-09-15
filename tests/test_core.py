"""Tests for core.py: create_workbook, read_sheet, write_sheet, append_rows, sheet mgmt."""

from __future__ import annotations

import pytest
from openpyxl import load_workbook

from pyhandlexl import grid
from pyhandlexl.core import (
    append_rows,
    create_csv,
    create_sheet,
    create_workbook,
    delete_sheet,
    delete_workbook,
    export_xl_to_csv,
    import_csv_to_xl,
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

    def test_values_keep_their_type(self, book):
        write_sheet(book, [[1, 2.5, True, "text", None]])
        assert read_sheet(book) == [[1, 2.5, True, "text"]]  # trailing None trimmed

    def test_whole_number_float_reads_back_as_int(self, book):
        write_sheet(book, [[5.0]])
        assert read_sheet(book) == [[5]]

    def test_datetime_round_trips(self, book):
        import datetime as dt

        value = dt.datetime(2026, 9, 10, 14, 30)
        write_sheet(book, [[value]])
        assert read_sheet(book) == [[value]]

    def test_empty_cell_reads_as_none(self, book):
        write_sheet(book, [["a", None, "c"]])
        assert read_sheet(book) == [["a", None, "c"]]

    def test_unsupported_type_raises(self, book):
        from pyhandlexl import CellTypeError

        with pytest.raises(CellTypeError):
            write_sheet(book, [[[1, 2, 3]]])

    def test_tz_aware_datetime_raises(self, book):
        import datetime as dt

        from pyhandlexl import CellTypeError

        aware = dt.datetime(2026, 9, 10, tzinfo=dt.timezone.utc)
        with pytest.raises(CellTypeError):
            write_sheet(book, [[aware]])

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

    def test_trailing_none_is_trimmed(self, book):
        write_sheet(book, [["a", None, None], ["b", "c", None]])
        assert read_sheet(book) == [["a"], ["b", "c"]]

    def test_empty_string_becomes_an_empty_cell(self, book):
        # openpyxl / Excel do not distinguish "" from a blank cell
        write_sheet(book, [["a", "", "b"]])
        assert read_sheet(book) == [["a", None, "b"]]

    def test_pad_makes_result_rectangular(self, book):
        write_sheet(book, [["a"], ["b", "c", "d"]])
        assert read_sheet(book, pad=True) == [["a", None, None], ["b", "c", "d"]]


class TestOrientation:
    def test_columns_orientation_transposes(self, book):
        write_sheet(book, [["h1", "h2"], ["v1", "v2"]], orientation="columns")
        assert read_sheet(book) == [["h1", "v1"], ["h2", "v2"]]

    def test_invalid_orientation_raises(self, book):
        with pytest.raises(ValueError):
            write_sheet(book, [["a"]], orientation="sideways")

    def test_read_columns_orientation_transposes(self, book):
        write_sheet(book, [["h1", "h2"], ["v1", "v2"]])
        assert read_sheet(book, orientation="columns") == [["h1", "v1"], ["h2", "v2"]]

    def test_read_columns_orientation_trims_trailing_none_per_column(self, book):
        write_sheet(book, [["a", "b"], ["c"]])
        assert read_sheet(book, orientation="columns") == [["a", "c"], ["b"]]

    def test_read_columns_orientation_with_pad(self, book):
        write_sheet(book, [["a", "b"], ["c"]])
        assert read_sheet(book, orientation="columns", pad=True) == [
            ["a", "c"],
            ["b", None],
        ]

    def test_read_invalid_orientation_raises(self, book):
        with pytest.raises(ValueError):
            read_sheet(book, orientation="sideways")


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


class TestImportCsvToXl:
    def test_basic_import(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("a,b\n1,2\n3,4\n", encoding="utf-8")
        import_csv_to_xl(csv_path, book, sheet="Data")
        assert read_sheet(book, "Data") == [["a", "b"], ["1", "2"], ["3", "4"]]

    def test_values_are_always_str(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("1,2.5,true\n", encoding="utf-8")
        import_csv_to_xl(csv_path, book, sheet="Data")
        assert read_sheet(book, "Data") == [["1", "2.5", "true"]]

    def test_creates_sheet_if_missing(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("a\n1\n", encoding="utf-8")
        import_csv_to_xl(csv_path, book, sheet="New")
        assert "New" in list_sheets(book)

    def test_replaces_existing_sheet_without_disturbing_others(self, book, tmp_path):
        write_sheet(book, [["old"]], sheet="Data")
        create_sheet(book, "Other")
        write_sheet(book, [["untouched"]], sheet="Other")
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("new\n", encoding="utf-8")
        import_csv_to_xl(csv_path, book, sheet="Data")
        assert read_sheet(book, "Data") == [["new"]]
        assert read_sheet(book, "Other") == [["untouched"]]

    def test_default_sheet_is_active(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("a\n1\n", encoding="utf-8")
        import_csv_to_xl(csv_path, book)
        assert read_sheet(book) == [["a"], ["1"]]

    def test_missing_csv_raises(self, book, tmp_path):
        with pytest.raises(FileNotFoundError):
            import_csv_to_xl(tmp_path / "nope.csv", book)

    def test_missing_workbook_raises(self, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("a\n1\n", encoding="utf-8")
        with pytest.raises(FileNotFoundError):
            import_csv_to_xl(csv_path, tmp_path / "nope.xlsx")

    def test_invalid_sheet_name_raises(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("a\n1\n", encoding="utf-8")
        with pytest.raises(SheetNameError):
            import_csv_to_xl(csv_path, book, sheet="a/b")

    def test_dimension_guard(self, book, tmp_path):
        csv_path = tmp_path / "huge.csv"
        csv_path.write_text(",".join(["x"] * 16_385) + "\n", encoding="utf-8")
        with pytest.raises(DimensionError):
            import_csv_to_xl(csv_path, book, sheet="Huge")
        assert "Huge" not in list_sheets(book)

    def test_quoted_fields_with_embedded_commas_and_newlines(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        # newline="" when writing our own fixture, same as the csv module
        # itself requires — otherwise Windows' text-mode newline translation
        # mangles the \n embedded inside the quoted field.
        with csv_path.open("w", newline="", encoding="utf-8") as f:
            f.write('a,b\n"hello, world","line1\nline2"\n')
        import_csv_to_xl(csv_path, book, sheet="Data")
        assert read_sheet(book, "Data") == [["a", "b"], ["hello, world", "line1\nline2"]]

    def test_ragged_rows_preserved(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("a,b,c\n1\n1,2\n", encoding="utf-8")
        import_csv_to_xl(csv_path, book, sheet="Data")
        assert read_sheet(book, "Data") == [["a", "b", "c"], ["1"], ["1", "2"]]

    def test_utf8_bom_is_stripped_by_default(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_bytes(b"\xef\xbb\xbfname,value\nAlice,10\n")
        import_csv_to_xl(csv_path, book, sheet="Data")
        assert read_sheet(book, "Data") == [["name", "value"], ["Alice", "10"]]

    def test_custom_encoding(self, book, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_bytes("café\n".encode("latin-1"))
        import_csv_to_xl(csv_path, book, sheet="Data", encoding="latin-1")
        assert read_sheet(book, "Data") == [["café"]]

    def test_non_csv_source_raises(self, book, tmp_path):
        not_csv = tmp_path / "data.txt"
        not_csv.write_text("a\n1\n", encoding="utf-8")
        with pytest.raises(ValueError):
            import_csv_to_xl(not_csv, book)

    def test_non_workbook_target_raises(self, tmp_path):
        csv_path = tmp_path / "data.csv"
        csv_path.write_text("a\n1\n", encoding="utf-8")
        with pytest.raises(ValueError):
            import_csv_to_xl(csv_path, tmp_path / "data2.csv")


class TestExportXlToCsv:
    def test_basic_export(self, book, tmp_path):
        write_sheet(book, [["a", "b"], [1, 2.5]], sheet="Data")
        csv_path = tmp_path / "out.csv"
        export_xl_to_csv(book, csv_path, sheet="Data")
        # raw bytes, not read_text() — confirms the real CRLF line terminator
        # csv.writer uses, rather than whatever text-mode read normalizes to
        assert csv_path.read_bytes() == b"a,b\r\n1,2.5\r\n"

    def test_none_becomes_empty_field(self, book, tmp_path):
        write_sheet(book, [["a", None, "c"]], sheet="Data")
        csv_path = tmp_path / "out.csv"
        export_xl_to_csv(book, csv_path, sheet="Data")
        assert csv_path.read_bytes() == b"a,,c\r\n"

    def test_default_sheet_is_active(self, book, tmp_path):
        write_sheet(book, [["a"]])
        csv_path = tmp_path / "out.csv"
        export_xl_to_csv(book, csv_path)
        assert csv_path.read_bytes() == b"a\r\n"

    def test_missing_workbook_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            export_xl_to_csv(tmp_path / "nope.xlsx", tmp_path / "out.csv")

    def test_missing_sheet_raises(self, book, tmp_path):
        with pytest.raises(SheetNotFoundError):
            export_xl_to_csv(book, tmp_path / "out.csv", sheet="Ghost")

    def test_refuses_to_overwrite_existing_file(self, book, tmp_path):
        write_sheet(book, [["a"]])
        csv_path = tmp_path / "out.csv"
        csv_path.write_text("existing", encoding="utf-8")
        with pytest.raises(FileExistsError):
            export_xl_to_csv(book, csv_path)
        assert csv_path.read_text(encoding="utf-8") == "existing"

    def test_no_temp_file_left_on_success(self, book, tmp_path):
        write_sheet(book, [["a"]])
        csv_path = tmp_path / "out.csv"
        export_xl_to_csv(book, csv_path)
        assert {f.name for f in tmp_path.iterdir()} == {"book.xlsx", "out.csv"}

    def test_no_partial_file_left_on_failure(self, book, tmp_path, monkeypatch):
        write_sheet(book, [["a"]])
        csv_path = tmp_path / "out.csv"

        import csv as csv_module

        def broken_writer(*_args, **_kwargs):
            raise RuntimeError("boom")

        monkeypatch.setattr(csv_module, "writer", broken_writer)
        with pytest.raises(RuntimeError):
            export_xl_to_csv(book, csv_path)
        assert {f.name for f in tmp_path.iterdir()} == {"book.xlsx"}

    def test_round_trip_through_import(self, book, tmp_path):
        write_sheet(book, [["a", "b"], ["1", "2"]], sheet="Data")
        csv_path = tmp_path / "out.csv"
        export_xl_to_csv(book, csv_path, sheet="Data")
        import_csv_to_xl(csv_path, book, sheet="Data2")
        assert read_sheet(book, "Data2") == [["a", "b"], ["1", "2"]]

    def test_non_workbook_source_raises(self, tmp_path):
        not_xlsx = tmp_path / "data.csv"
        not_xlsx.write_text("a\n1\n", encoding="utf-8")
        with pytest.raises(ValueError):
            export_xl_to_csv(not_xlsx, tmp_path / "out.csv")

    def test_non_csv_target_raises(self, book, tmp_path):
        write_sheet(book, [["a"]])
        with pytest.raises(ValueError):
            export_xl_to_csv(book, tmp_path / "out.txt")


class TestCreateCsv:
    def test_creates_empty_file(self, tmp_path):
        path = tmp_path / "new.csv"
        create_csv(path)
        assert path.exists()
        assert path.read_bytes() == b""

    def test_already_exists_raises(self, tmp_path):
        path = tmp_path / "new.csv"
        create_csv(path)
        with pytest.raises(FileExistsError):
            create_csv(path)

    def test_wrong_extension_raises(self, tmp_path):
        with pytest.raises(ValueError):
            create_csv(tmp_path / "new.txt")


class TestReadSheetCsv:
    def test_basic_read(self, tmp_path):
        path = tmp_path / "data.csv"
        path.write_text("a,b\n1,2\n", encoding="utf-8")
        assert read_sheet(path) == [["a", "b"], ["1", "2"]]

    def test_no_trimming_of_trailing_empty_fields(self, tmp_path):
        path = tmp_path / "data.csv"
        path.write_text("a,b,,\n", encoding="utf-8")
        assert read_sheet(path) == [["a", "b", "", ""]]

    def test_ragged_rows_preserved(self, tmp_path):
        path = tmp_path / "data.csv"
        path.write_text("a,b,c\n1\n1,2\n", encoding="utf-8")
        assert read_sheet(path) == [["a", "b", "c"], ["1"], ["1", "2"]]

    def test_pad_fills_with_empty_string(self, tmp_path):
        path = tmp_path / "data.csv"
        path.write_text("a,b,c\n1\n", encoding="utf-8")
        assert read_sheet(path, pad=True) == [["a", "b", "c"], ["1", "", ""]]

    def test_sheet_given_raises(self, tmp_path):
        path = tmp_path / "data.csv"
        path.write_text("a\n1\n", encoding="utf-8")
        with pytest.raises(ValueError):
            read_sheet(path, sheet="Sheet1")

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            read_sheet(tmp_path / "nope.csv")


class TestWriteSheetCsv:
    def test_replaces_whole_file(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["old"]])
        write_sheet(path, [["new", "content"]])
        assert read_sheet(path) == [["new", "content"]]

    def test_values_are_stringified(self, tmp_path):
        import datetime as dt

        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [[1, 2.5, True, None, dt.date(2026, 1, 1)]])
        assert read_sheet(path) == [["1", "2.5", "True", "", "2026-01-01"]]

    def test_orientation_columns(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [[1, 2], [3, 4]], orientation="columns")
        assert read_sheet(path) == [["1", "3"], ["2", "4"]]

    def test_requires_file_to_exist(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            write_sheet(tmp_path / "nope.csv", [["a"]])

    def test_sheet_given_raises(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        with pytest.raises(ValueError):
            write_sheet(path, [["a"]], sheet="Sheet1")

    def test_no_dimension_limit(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["x"] * 20_000])
        assert len(read_sheet(path)[0]) == 20_000

    def test_no_temp_file_left_on_success(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["a"]])
        assert {f.name for f in tmp_path.iterdir()} == {"data.csv"}

    def test_no_partial_file_left_on_failure(self, tmp_path, monkeypatch):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["original"]])

        import csv as csv_module

        def broken_writer(*_args, **_kwargs):
            raise RuntimeError("boom")

        monkeypatch.setattr(csv_module, "writer", broken_writer)
        with pytest.raises(RuntimeError):
            write_sheet(path, [["new"]])
        assert read_sheet(path) == [["original"]]
        assert {f.name for f in tmp_path.iterdir()} == {"data.csv"}


class TestAppendRowsCsv:
    def test_appends_to_existing_content(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["a"]])
        append_rows(path, [["b"], ["c"]])
        assert read_sheet(path) == [["a"], ["b"], ["c"]]

    def test_empty_rows_is_a_noop(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["a"]])
        append_rows(path, [])
        assert read_sheet(path) == [["a"]]

    def test_appends_to_an_empty_file(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        append_rows(path, [["a"], ["b"]])
        assert read_sheet(path) == [["a"], ["b"]]

    def test_requires_file_to_exist(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            append_rows(tmp_path / "nope.csv", [["a"]])

    def test_sheet_given_raises(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        with pytest.raises(ValueError):
            append_rows(path, [["a"]], sheet="Sheet1")

    def test_values_are_stringified(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        append_rows(path, [[1, None, True]])
        assert read_sheet(path) == [["1", "", "True"]]

    def test_no_dimension_limit(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["a"]])
        append_rows(path, [["x"] * 20_000])
        assert len(read_sheet(path)[1]) == 20_000

    def test_does_not_read_existing_content(self, tmp_path, monkeypatch):
        # the whole point of append_rows existing separately from
        # write_sheet(path, grid.append_row(read_sheet(path), values)) is
        # that it never materialises the existing rows — this proves it
        # architecturally, not just by checking the final content
        from pyhandlexl import core as core_module

        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["a"], ["b"]])

        def boom(_path):
            raise AssertionError("append_rows must not read existing CSV content")

        monkeypatch.setattr(core_module, "_read_csv_rows", boom)
        append_rows(path, [["c"]])
        monkeypatch.undo()  # read_sheet also uses _read_csv_rows — only verify after
        assert read_sheet(path) == [["a"], ["b"], ["c"]]

    def test_appends_correctly_when_file_lacks_trailing_newline(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        with path.open("w", newline="", encoding="utf-8") as f:
            f.write("a,b")  # deliberately no trailing newline
        append_rows(path, [["c", "d"]])
        assert read_sheet(path) == [["a", "b"], ["c", "d"]]
        assert path.read_bytes() == b"a,b\r\nc,d\r\n"


class TestGridEditingOnCsv:
    def test_grid_functions_compose_with_csv_read_write(self, tmp_path):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [[1, 2], [3, 4]])

        g = read_sheet(path)
        g = grid.insert_row(g, 1, ["h1", "h2"])
        g = grid.set_value(g, 2, 1, "CHANGED")
        g = grid.append_column(g, ["h3", "x", "y"])
        write_sheet(path, g)

        assert read_sheet(path) == [
            ["h1", "h2", "h3"],
            ["CHANGED", "2", "x"],
            ["3", "4", "y"],
        ]

    def test_show_works_on_csv_sourced_grid(self, tmp_path, capsys):
        path = tmp_path / "data.csv"
        create_csv(path)
        write_sheet(path, [["a", "b"], ["1", "2"]])
        grid.show(read_sheet(path))
        out = capsys.readouterr().out
        assert "a" in out and "1" in out
