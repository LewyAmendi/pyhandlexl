"""Cell-level formatting: CellFormat on Table cells, and Styled values in grids."""

from __future__ import annotations

import warnings

import pytest
from openpyxl import load_workbook
from openpyxl.styles import Font

from pyhandlexl import (
    CellFormat,
    MergeConflictWarning,
    Styled,
    Table,
    append_rows,
    create_csv,
    create_sheet,
    grid,
    read_sheet,
    write_sheet,
)
from pyhandlexl._cell_format import formats_from_json, formats_to_json

BOLD = CellFormat(bold=True)
YELLOW = CellFormat(fill="FFFF00")


def sheet_cell(path, sheet, ref):
    return load_workbook(path)[sheet][ref]


@pytest.fixture
def path(book):
    create_sheet(book, "S")
    Table([[1, 2], [3, 4]], ["r1", "r2"], "id", column_headers=["x", "y"], name="A").create(
        book, "S"
    )
    Table([[5], [6]], ["r1", "r2"], column_headers=["z"], name="B").create(book, "S")
    return book


# A is at A1:C4 (marker row 1, header row 2, data from row 3); B starts at E1.


class TestCellFormatValue:
    def test_defaults_change_nothing(self):
        assert CellFormat().is_empty()
        assert not BOLD.is_empty()

    @pytest.mark.parametrize(
        "kwargs",
        [
            {"bold": 1},
            {"italic": "yes"},
            {"font_color": "red"},
            {"font_color": ""},
            {"fill": "12345"},
            {"number_format": ""},
            {"align": "middle"},
            {"wrap": "yes"},
        ],
    )
    def test_bad_fields_are_refused(self, kwargs):
        with pytest.raises((TypeError, ValueError)):
            CellFormat(**kwargs)

    def test_empty_fill_means_no_fill(self):
        assert CellFormat(fill="").fill == ""

    def test_styled_needs_a_format(self):
        with pytest.raises(TypeError):
            Styled(1, "bold")  # type: ignore[arg-type]
        with pytest.raises(TypeError):
            Styled(Styled(1, BOLD), BOLD)


class TestTableFormats:
    def test_set_get_and_the_three_targets(self, path):
        t = Table.read(path, "A")
        t.set_format(BOLD, row="r1", column="x")
        t.set_format(YELLOW, row="r2")
        t.set_format(CellFormat(align="right"), column="y")
        assert t.get_format("r1", "x") == BOLD
        assert t.get_format("r2", "x") == YELLOW
        assert t.get_format("r1", "y") == CellFormat(align="right")
        assert t.get_format("r2", "y") == CellFormat(fill="FFFF00", align="right")
        assert Table.read(path, "A").get_format("r1", "x") is None

    def test_formats_combine_field_by_field(self, path):
        t = Table.read(path, "A")
        t.set_format(BOLD, row="r1", column="x")
        t.set_format(YELLOW, row="r1", column="x")
        assert t.get_format("r1", "x") == CellFormat(bold=True, fill="FFFF00")
        t.set_format(CellFormat(bold=False), row="r1", column="x")
        assert t.get_format("r1", "x") == CellFormat(bold=False, fill="FFFF00")

    def test_an_empty_format_removes_nothing_new(self, path):
        t = Table.read(path, "A")
        t.set_format(CellFormat(), row="r1", column="x")
        assert t.get_format("r1", "x") is None

    def test_clear(self, path):
        t = Table.read(path, "A")
        t.set_format(BOLD, row="r1")
        t.set_format(BOLD, row="r2")
        t.clear_format(row="r1", column="x")
        assert t.get_format("r1", "x") is None
        assert t.get_format("r1", "y") == BOLD
        t.clear_format(column="y")
        assert t.get_format("r1", "y") is None and t.get_format("r2", "y") is None
        t.clear_format()
        assert t.get_format("r2", "x") is None

    def test_misuse(self, path):
        t = Table.read(path, "A")
        with pytest.raises(ValueError):
            t.set_format(BOLD)
        with pytest.raises(TypeError):
            t.set_format("bold", row="r1")  # type: ignore[arg-type]
        with pytest.raises(KeyError):
            t.set_format(BOLD, row="nope")
        with pytest.raises(KeyError):
            t.get_format("r1", "nope")

    def test_applied_in_the_file_over_the_table_style(self, path):
        t = Table.read(path, "A")
        t.set_format(
            CellFormat(bold=True, fill="FFFF00", number_format="0.00"), row="r1", column="y"
        )
        t.write(path)
        cell = sheet_cell(path, "S", "C3")
        assert cell.font.b and cell.fill.fgColor.rgb.endswith("FFFF00")
        assert cell.number_format == "0.00"
        assert cell.value == 2  # the value is untouched
        other = sheet_cell(path, "S", "B3")  # a neighbour keeps the table's own look
        assert not other.font.b and other.number_format == "General"

    def test_wrap_and_align_share_one_alignment_without_clobbering(self, path):
        t = Table.read(path, "A")
        t.set_format(CellFormat(align="center"), row="r1", column="x")
        t.set_format(CellFormat(wrap=True), row="r1", column="x")
        t.write(path)
        cell = sheet_cell(path, "S", "B3")
        assert cell.alignment.wrap_text and cell.alignment.horizontal == "center"
        t = Table.read(path, "A")
        assert t.get_format("r1", "x") == CellFormat(align="center", wrap=True)
        t.set_format(CellFormat(wrap=False), row="r1", column="x")
        t.write(path)
        cell = sheet_cell(path, "S", "B3")
        assert not cell.alignment.wrap_text and cell.alignment.horizontal == "center"

    def test_a_format_is_stored_and_read_back(self, path):
        t = Table.read(path, "A")
        t.set_format(CellFormat(bold=True, font_color="C00000", align="center"), row="r2")
        t.write(path)
        back = Table.read(path, "A")
        assert back.get_format("r2", "y") == CellFormat(
            bold=True, font_color="C00000", align="center"
        )
        assert back.get_format("r1", "y") is None

    def test_cleared_formats_leave_the_file(self, path):
        t = Table.read(path, "A")
        t.set_format(BOLD, row="r1", column="y")
        t.write(path)
        t = Table.read(path, "A")
        t.clear_format()
        t.write(path)
        assert not sheet_cell(path, "S", "C3").font.b
        assert Table.read(path, "A").get_format("r1", "y") is None

    def test_create_stores_formats_too(self, book):
        create_sheet(book, "N")
        t = Table([[1]], ["r"], column_headers=["c"], name="N1")
        t.set_format(BOLD, row="r", column="c")
        t.create(book, "N")
        assert Table.read(book, "N1").get_format("r", "c") == BOLD


class TestFormatsFollowTheirCell:
    def test_insert_drop_and_rename(self, path):
        t = Table.read(path, "A")
        t.set_format(BOLD, row="r2", column="y")
        t.insert_row(1, "top", [0, 0])
        t.insert_column(1, "first", [0, 0, 0])
        t.rename_row("r2", "second")
        t.rename_column("y", "why")
        t.write(path)
        back = Table.read(path, "A")
        assert back.row_labels == ["top", "r1", "second"]
        assert back.get_format("second", "why") == BOLD
        # row 'second' is the 5th sheet row, column 'why' the 4th column (D)
        assert sheet_cell(path, "S", "D5").font.b
        assert not sheet_cell(path, "S", "D3").font.b  # where the old position was

    def test_dropping_removes_the_format(self, path):
        t = Table.read(path, "A")
        t.set_format(BOLD, row="r1")
        t.drop_row("r1")
        t.write(path)
        back = Table.read(path, "A")
        assert back.get_format("r2", "x") is None
        t.add_row("r1", [1, 2])
        t.write(path)
        assert Table.read(path, "A").get_format("r1", "x") is None  # not resurrected

    def test_a_neighbour_growing_moves_formats_with_the_table(self, path):
        b = Table.read(path, "B")
        b.set_format(CellFormat(number_format="0.00", fill="00FF00"), column="z")
        b.write(path)
        a = Table.read(path, "A")
        a.add_column("w", [7, 8])
        a.write(path)
        moved = sheet_cell(path, "S", "G3")  # B was at F:F-ish: it moved one column right
        assert moved.value == 5
        assert moved.number_format == "0.00" and moved.fill.fgColor.rgb.endswith("00FF00")
        assert Table.read(path, "B").get_format("r1", "z") is not None

    def test_hand_formatting_survives_a_neighbour_growing(self, path):
        wb = load_workbook(path)
        target = wb["S"]["F3"]
        target.font = Font(bold=True, italic=True)
        target.number_format = "0.0%"
        wb.save(path)
        a = Table.read(path, "A")
        a.add_column("w", [7, 8])
        a.write(path)
        moved = sheet_cell(path, "S", "G3")
        assert moved.value == 5
        assert moved.font.b and moved.font.i and moved.number_format == "0.0%"
        assert sheet_cell(path, "S", "E3").value is None  # the old label column is now the gap


class TestFormatsMerge:
    def test_another_writers_format_survives_our_unrelated_edit(self, path):
        alice, bob = Table.read(path, "A"), Table.read(path, "A")
        bob.set_format(BOLD, row="r1", column="x")
        bob.write(path)
        alice.set_cell(row="r2", column="y", value=40)
        with warnings.catch_warnings():
            warnings.simplefilter("error", MergeConflictWarning)
            alice.write(path)
        assert alice.get_format("r1", "x") == BOLD
        assert Table.read(path, "A").get_format("r1", "x") == BOLD

    def test_both_format_the_same_cell_ours_wins(self, path):
        alice, bob = Table.read(path, "A"), Table.read(path, "A")
        bob.set_format(BOLD, row="r1", column="x")
        bob.write(path)
        alice.set_format(YELLOW, row="r1", column="x")
        with pytest.warns(MergeConflictWarning, match="format of"):
            alice.write(path)
        assert Table.read(path, "A").get_format("r1", "x") == YELLOW


class TestSchemaStorage:
    def test_round_trip_and_chunking(self):
        formats = {(i, 0): CellFormat(number_format=f'0.{"0" * 5}"{i}"') for i in range(3000)}
        formats[(5000, 7)] = BOLD
        chunks = formats_to_json(formats)
        assert len(chunks) > 1
        assert all(len(c) <= 32767 for c in chunks)
        assert formats_from_json(chunks) == formats

    def test_nothing_to_store(self):
        assert formats_to_json({}) == []
        assert formats_from_json([]) == {}

    def test_many_formats_survive_a_real_file(self, book):
        create_sheet(book, "Big")
        rows = [f"r{i}" for i in range(1500)]
        t = Table([[i, i] for i in range(1500)], rows, column_headers=["a", "b"], name="Big")
        for i, label in enumerate(rows):
            t.set_format(CellFormat(number_format=f'0.00" unit {i}"'), row=label, column="a")
        t.create(book, "Big")
        back = Table.read(book, "Big")
        assert back.get_format("r1499", "a") == CellFormat(number_format='0.00" unit 1499"')
        assert back.get_format("r0", "b") is None
        # and the schema stays usable by the other calls
        assert Table.read(book, "Big").row_labels[-1] == "r1499"


class TestGridFormatting:
    def test_write_and_read_back(self, book):
        write_sheet(
            book, [["h", Styled("n", YELLOW)], [Styled(1.5, CellFormat(number_format="0.00")), 2]]
        )
        assert read_sheet(book) == [["h", "n"], [1.5, 2]]  # plain by default
        got = read_sheet(book, formats=True)
        assert got[0][1] == Styled("n", YELLOW)
        assert got[1][0] == Styled(1.5, CellFormat(number_format="0.00"))
        assert got[0][0] == "h" and got[1][1] == 2

    def test_wrap_round_trips_in_a_grid(self, book):
        write_sheet(book, [[Styled("a long sentence", CellFormat(wrap=True, align="right"))]])
        got = read_sheet(book, formats=True)[0][0]
        assert got == Styled("a long sentence", CellFormat(wrap=True, align="right"))

    def test_helpers_move_the_format_with_the_value(self, book):
        g = [["a", Styled("b", BOLD)], ["c", "d"]]
        g = grid.insert_column(g, 1, ["x", "y"])
        g = grid.insert_row(g, 1, ["top"] * 3)
        write_sheet(book, g)
        assert sheet_cell(book, "Sheet", "C2").font.b
        assert not sheet_cell(book, "Sheet", "B2").font.b
        t = grid.transpose(g)
        assert t[2][1] == Styled("b", BOLD)

    def test_orientation_columns(self, book):
        write_sheet(book, [[Styled("a", BOLD), "b"]], orientation="columns")
        assert sheet_cell(book, "Sheet", "A1").font.b
        assert not sheet_cell(book, "Sheet", "A2").font.b

    def test_append_rows_formats_land_on_the_appended_rows(self, book):
        write_sheet(book, [[1], [2]])
        append_rows(book, [[Styled(3, BOLD)], [4]])
        assert sheet_cell(book, "Sheet", "A3").font.b
        assert not sheet_cell(book, "Sheet", "A2").font.b

    def test_styled_in_a_csv_is_just_its_value(self, tmp_path):
        csv = tmp_path / "a.csv"
        create_csv(csv)
        write_sheet(csv, [[Styled(1, BOLD), "x"]])
        assert read_sheet(csv) == [["1", "x"]]
        with pytest.raises(ValueError):
            read_sheet(csv, formats=True)

    def test_display_and_dataframes_see_plain_values(self, capsys):
        g = [["h"], [Styled(1, BOLD)]]
        assert "Styled" not in grid.show(g)
        assert "Styled" not in capsys.readouterr().out
        pytest.importorskip("pandas")
        assert grid.to_dataframe(g)["h"].tolist() == [1]

    def test_a_styled_value_is_still_checked(self, book):
        from pyhandlexl import CellTypeError

        with pytest.raises(CellTypeError):
            write_sheet(book, [[Styled(object(), BOLD)]])
