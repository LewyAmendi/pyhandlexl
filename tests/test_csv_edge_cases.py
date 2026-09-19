"""CSV files as raw grids: every awkward shape a real file can have."""

from __future__ import annotations

import random
import sys
from typing import ClassVar

import pytest

from pyhandlexl import (
    append_rows,
    create_csv,
    create_sheet,
    export_xl_to_csv,
    import_csv_to_xl,
    read_sheet,
    write_sheet,
)


def csv_file(tmp_path, content: bytes, name: str = "data.csv"):
    path = tmp_path / name
    path.write_bytes(content)
    return path


class TestReading:
    @pytest.mark.parametrize(
        ("content", "expected"),
        [
            (b"", []),
            (b"\n", [[]]),
            (b"\n\n", [[], []]),
            (b"a,b\n\nc,d\n", [["a", "b"], [], ["c", "d"]]),
            (b"a,b\r\nc,d\r\n", [["a", "b"], ["c", "d"]]),
            (b"a,b\nc,d", [["a", "b"], ["c", "d"]]),
            (b"a,b\rc,d\r", [["a", "b"], ["c", "d"]]),
            (b'"a,1","he said ""hi""","x\ny"\n', [["a,1", 'he said "hi"', "x\ny"]]),
            (b'"x\r\ny",z\n', [["x\r\ny", "z"]]),
            (b"\xef\xbb\xbfa,b\n", [["a", "b"]]),
            (b"a,b,c\nd\ne,f\n", [["a", "b", "c"], ["d"], ["e", "f"]]),
            (b"a;b;c\n", [["a;b;c"]]),
            (b"a\tb\n", [["a\tb"]]),
            (b",\n", [["", ""]]),
            (b",,\n,,\n", [["", "", ""], ["", "", ""]]),
            (b'"",""\n', [["", ""]]),
            (b" a , b \n", [[" a ", " b "]]),
            ("é,日本語,😀\n".encode(), [["é", "日本語", "😀"]]),
            (b'"unterminated\n', [["unterminated\n"]]),
        ],
        ids=lambda v: repr(v)[:40],
    )
    def test_read_sheet(self, tmp_path, content, expected):
        assert read_sheet(csv_file(tmp_path, content)) == expected

    def test_a_nul_byte_reads_on_python_311_and_is_a_clear_error_before_that(self, tmp_path):
        # the csv module itself refused NUL up to 3.10; either way, never a bare _csv.Error
        path = csv_file(tmp_path, b"a\x00b,c\n")
        if sys.version_info >= (3, 11):
            assert read_sheet(path) == [["a\x00b", "c"]]
        else:
            with pytest.raises(ValueError, match="line 1"):
                read_sheet(path)

    def test_pad_fills_ragged_rows_with_empty_strings(self, tmp_path):
        path = csv_file(tmp_path, b"a,b,c\nd\n")
        assert read_sheet(path, pad=True) == [["a", "b", "c"], ["d", "", ""]]

    def test_orientation_columns(self, tmp_path):
        path = csv_file(tmp_path, b"a,b\nc,d\n")
        assert read_sheet(path, orientation="columns") == [["a", "c"], ["b", "d"]]

    def test_a_field_larger_than_the_csv_modules_default_limit(self, tmp_path):
        # 131,072 characters is where the csv module gives up; a CSV has no size limit here
        big = "x" * 300_000
        path = csv_file(tmp_path, f"a,{big},b\n".encode())
        assert read_sheet(path) == [["a", big, "b"]]

    def test_reading_a_huge_field_does_not_change_the_csv_modules_global_limit(self, tmp_path):
        import csv

        before = csv.field_size_limit()
        read_sheet(csv_file(tmp_path, b"a,b\n"))
        assert csv.field_size_limit() == before

    def test_the_limit_is_restored_even_when_reading_fails(self, tmp_path):
        import csv

        before = csv.field_size_limit()
        with pytest.raises(UnicodeDecodeError):
            read_sheet(csv_file(tmp_path, b"caf\xe9,b\n"))  # latin-1: not valid UTF-8
        assert csv.field_size_limit() == before

    def test_undecodable_bytes_raise_a_unicode_error_not_garbage(self, tmp_path):
        with pytest.raises(UnicodeDecodeError):
            read_sheet(csv_file(tmp_path, b"\xff\xfe\x00bad"))


class TestWriting:
    def test_every_value_becomes_a_string_and_quoting_is_correct(self, tmp_path):
        path = csv_file(tmp_path, b"")
        write_sheet(path, [["a", None, 1, 2.5, True, "x,y", 'q"r', "l1\nl2", ""]])
        assert path.read_bytes() == b'a,,1,2.5,True,"x,y","q""r","l1\nl2",\r\n'
        assert read_sheet(path) == [["a", "", "1", "2.5", "True", "x,y", 'q"r', "l1\nl2", ""]]

    def test_a_bare_string_row_is_refused_not_split_into_characters(self, tmp_path):
        path = csv_file(tmp_path, b"keep\n")
        with pytest.raises(TypeError, match="split it into individual characters"):
            write_sheet(path, ["hello"])
        with pytest.raises(TypeError):
            append_rows(path, ["hello"])
        assert path.read_bytes() == b"keep\n"

    def test_write_sheet_with_no_rows_empties_the_file(self, tmp_path):
        path = csv_file(tmp_path, b"old,content\n")
        write_sheet(path, [])
        assert path.read_bytes() == b""

    def test_a_row_of_one_empty_string_is_told_apart_from_an_empty_row(self, tmp_path):
        # the writer quotes a lone empty field, so it can't be mistaken for a blank line
        path = csv_file(tmp_path, b"")
        write_sheet(path, [["a"], [""], [], ["b"]])
        assert path.read_bytes() == b'a\r\n""\r\n\r\nb\r\n'
        assert read_sheet(path) == [["a"], [""], [], ["b"]]

    def test_uppercase_extension_is_still_csv(self, tmp_path):
        path = csv_file(tmp_path, b"", "DATA.CSV")
        write_sheet(path, [["a", "b"]])
        assert path.read_bytes() == b"a,b\r\n"

    def test_sheet_is_not_meaningful(self, tmp_path):
        path = csv_file(tmp_path, b"")
        for call in (
            lambda: write_sheet(path, [["a"]], sheet="S"),
            lambda: append_rows(path, [["a"]], sheet="S"),
            lambda: read_sheet(path, sheet="S"),
        ):
            with pytest.raises(ValueError, match="sheet is not meaningful"):
                call()


class TestAppending:
    @pytest.mark.parametrize(
        ("existing", "expected"),
        [
            (b"", b"c,d\r\n"),
            (b"a,b", b"a,b\r\nc,d\r\n"),
            (b"a,b\n", b"a,b\nc,d\r\n"),
            (b"a,b\r\n", b"a,b\r\nc,d\r\n"),
        ],
        ids=["empty file", "no trailing newline", "LF ending", "CRLF ending"],
    )
    def test_append_starts_on_a_fresh_line(self, tmp_path, existing, expected):
        path = csv_file(tmp_path, existing)
        append_rows(path, [["c", "d"]])
        assert path.read_bytes() == expected

    def test_nothing_appended_means_nothing_changes(self, tmp_path):
        path = csv_file(tmp_path, b"a,b")
        append_rows(path, [])
        assert path.read_bytes() == b"a,b"  # not even the missing newline is added

    def test_an_empty_row_appends_a_blank_line(self, tmp_path):
        path = csv_file(tmp_path, b"a,b\n")
        append_rows(path, [[]])
        assert read_sheet(path) == [["a", "b"], []]

    def test_appending_many_rows_keeps_them_in_order(self, tmp_path):
        path = csv_file(tmp_path, b"")
        create_csv_rows = [[str(i), f"v{i}"] for i in range(200)]
        for chunk in (create_csv_rows[:70], create_csv_rows[70:71], create_csv_rows[71:]):
            append_rows(path, chunk)
        assert read_sheet(path) == create_csv_rows


class TestCreate:
    def test_create_csv_refuses_other_suffixes_and_existing_files(self, tmp_path):
        with pytest.raises(ValueError):
            create_csv(tmp_path / "data.txt")
        target = tmp_path / "data.csv"
        target.write_bytes(b"precious")
        with pytest.raises(FileExistsError):
            create_csv(target)
        assert target.read_bytes() == b"precious"

    def test_a_missing_csv_is_file_not_found_everywhere(self, tmp_path):
        missing = tmp_path / "missing.csv"
        for call in (
            lambda: read_sheet(missing),
            lambda: write_sheet(missing, [["a"]]),
            lambda: append_rows(missing, [["a"]]),
        ):
            with pytest.raises(FileNotFoundError):
                call()
        assert list(tmp_path.iterdir()) == []


class TestConvertingBetweenCsvAndXlsx:
    def test_export_stringifies_and_import_returns_text(self, book, tmp_path):
        write_sheet(book, [["007", "1e3", "TRUE", "2026-01-01", None, ""], [1, 2.5, True, None]])
        out = tmp_path / "out.csv"
        export_xl_to_csv(book, out)
        assert read_sheet(out) == [["007", "1e3", "TRUE", "2026-01-01"], ["1", "2.5", "True"]]
        target = tmp_path / "back.xlsx"
        from pyhandlexl import create_workbook

        create_workbook(target)
        import_csv_to_xl(out, target)
        assert read_sheet(target) == [["007", "1e3", "TRUE", "2026-01-01"], ["1", "2.5", "True"]]

    def test_unicode_survives_both_ways(self, book, tmp_path):
        write_sheet(book, [["é", "日本語", "😀", "a\nb"]])
        out = tmp_path / "out.csv"
        export_xl_to_csv(book, out)
        assert read_sheet(out) == [["é", "日本語", "😀", "a\nb"]]

    def test_export_refuses_to_overwrite(self, book, tmp_path):
        out = tmp_path / "out.csv"
        out.write_bytes(b"precious")
        with pytest.raises(FileExistsError):
            export_xl_to_csv(book, out)
        assert out.read_bytes() == b"precious"

    def test_windows_line_endings_inside_a_quoted_csv_field_are_stored_as_newlines(
        self, book, tmp_path
    ):
        source = csv_file(tmp_path, b'"line1\r\nline2",x\r\n')
        import_csv_to_xl(source, book, sheet="Imp")
        assert read_sheet(book, "Imp") == [["line1\nline2", "x"]]

    def test_import_into_a_lookalike_or_clashing_sheet_is_refused(self, book, tmp_path):
        create_sheet(book, "Log")
        source = csv_file(tmp_path, b"a\n")
        with pytest.raises(ValueError):
            import_csv_to_xl(source, book, sheet="LOG")


class TestRandomRoundTrips:
    """Whatever text goes into a CSV comes back out, byte for byte, for any awkward mix."""

    ALPHABET: ClassVar[list[str]] = [
        *list("ab ,\";\n\r\t'é日😀\\|"),
        "",
        " ",
        "  ",
        "TRUE",
        "007",
        "=x",
        "\x7f",
    ]

    @pytest.mark.parametrize("seed", range(40))
    def test_write_then_read(self, tmp_path, seed):
        rng = random.Random(seed)
        rows = []
        for _ in range(rng.randint(1, 8)):
            row = []
            for _ in range(rng.randint(1, 6)):
                row.append("".join(rng.choice(self.ALPHABET) for _ in range(rng.randint(0, 9))))
            rows.append(row)
        path = csv_file(tmp_path, b"")
        write_sheet(path, rows)
        assert read_sheet(path) == rows

    @pytest.mark.parametrize("seed", range(20))
    def test_append_chunks_then_read(self, tmp_path, seed):
        rng = random.Random(500 + seed)
        path = csv_file(tmp_path, b"")
        expected = []
        for _ in range(rng.randint(1, 6)):
            chunk = [
                ["".join(rng.choice(self.ALPHABET) for _ in range(rng.randint(0, 6))), "k"]
                for _ in range(rng.randint(1, 4))
            ]
            append_rows(path, chunk)
            expected += chunk
        assert read_sheet(path) == expected
