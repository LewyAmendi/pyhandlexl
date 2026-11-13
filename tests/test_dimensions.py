"""Column widths, row heights, and opt-in text fitting (TableStyle.fit_text / fit_text=True)."""

from __future__ import annotations

import re
import warnings
import zipfile

import pytest
from openpyxl import load_workbook
from openpyxl.worksheet.dimensions import ColumnDimension

from pyhandlexl import (
    CellFormat,
    MergeConflictWarning,
    Styled,
    Table,
    TableStyle,
    append_rows,
    create_csv,
    create_sheet,
    delete_table,
    read_dimensions,
    write_sheet,
)

LONG = "word " * 30  # 150 characters


def widths(path, sheet="S"):
    return read_dimensions(path, sheet).column_widths


def heights(path, sheet="S"):
    return read_dimensions(path, sheet).row_heights


@pytest.fixture
def path(book):
    create_sheet(book, "S")
    Table([[1, 2], [3, 4]], ["r1", "r2"], "id", column_headers=["x", "y"], name="A").create(
        book, "S"
    )
    Table([[5], [6]], ["r1", "r2"], column_headers=["z"], name="B").create(book, "S")
    return book


# A: label A, data x=B, y=C (rows 3 and 4). B: label E, data z=F.


class TestNothingChangesByDefault:
    def test_no_dimensions_are_written(self, path):
        t = Table.read(path, "A")
        t.add_row("r3", [7, LONG])
        t.write(path)
        assert widths(path) == {} and heights(path) == {}
        assert not load_workbook(path)["S"]["C5"].alignment.wrap_text

    def test_the_stored_style_is_unchanged_for_old_readers(self, path):
        sheet = load_workbook(path)["_pyhandlexl_tables"]
        assert "fit_text" not in sheet["G2"].value


class TestTableSizes:
    def test_set_get_clear(self, path):
        t = Table.read(path, "A")
        t.set_column_width("x", 20)
        t.set_row_height("r2", 40)
        assert t.get_column_width("x") == 20 and t.get_column_width("y") is None
        assert t.get_row_height("r2") == 40 and t.get_row_height("r1") is None
        t.clear_column_width("x")
        t.clear_row_height()
        assert t.get_column_width("x") is None and t.get_row_height("r2") is None

    @pytest.mark.parametrize("bad", [0, -3, 256, True, "wide", None])
    def test_bad_width(self, path, bad):
        with pytest.raises((TypeError, ValueError)):
            Table.read(path, "A").set_column_width("x", bad)

    @pytest.mark.parametrize("bad", [0, 410, False, "tall"])
    def test_bad_height(self, path, bad):
        with pytest.raises((TypeError, ValueError)):
            Table.read(path, "A").set_row_height("r1", bad)

    def test_unknown_labels(self, path):
        t = Table.read(path, "A")
        for call in (
            lambda: t.set_column_width("nope", 10),
            lambda: t.get_column_width("nope"),
            lambda: t.set_row_height("nope", 10),
            lambda: t.clear_row_height("nope"),
        ):
            with pytest.raises(KeyError):
                call()

    def test_written_read_back_and_applied(self, path):
        t = Table.read(path, "A")
        t.set_column_width("y", 25.5)
        t.set_row_height("r2", 33)
        t.write(path)
        assert widths(path) == {3: 25.5}
        assert heights(path) == {4: 33}
        back = Table.read(path, "A")
        assert back.get_column_width("y") == 25.5 and back.get_row_height("r2") == 33

    def test_they_follow_their_column_and_row(self, path):
        t = Table.read(path, "A")
        t.set_column_width("y", 25)
        t.set_row_height("r2", 33)
        t.write(path)
        t = Table.read(path, "A")
        t.insert_column(1, "new", [0, 0])
        t.insert_row(1, "top", [0, 0, 0])
        t.rename_column("y", "why")
        t.write(path)
        assert widths(path) == {4: 25}  # 'why' is now the 4th column
        assert heights(path) == {5: 33}  # 'r2' is now the 3rd data row
        t.drop_column("why")
        t.drop_row("r2")
        t.write(path)
        assert widths(path) == {} and heights(path) == {}
        assert "why" not in Table.read(path, "A").column_headers

    def test_a_cleared_size_leaves_the_file(self, path):
        t = Table.read(path, "A")
        t.set_column_width("x", 20)
        t.write(path)
        t = Table.read(path, "A")
        t.clear_column_width()
        t.write(path)
        assert widths(path) == {}

    def test_deleting_a_table_takes_its_sizes_back(self, path):
        t = Table.read(path, "A")
        t.set_column_width("x", 20)
        t.set_row_height("r1", 30)
        t.style = TableStyle(fit_text=True)
        t.write(path)
        assert widths(path) and heights(path)
        delete_table(path, "A")
        assert 2 not in widths(path) and heights(path) == {}

    def test_a_width_changed_by_hand_afterwards_is_left_alone(self, path):
        t = Table.read(path, "A")
        t.set_column_width("x", 20)
        t.write(path)
        wb = load_workbook(path)
        wb["S"].column_dimensions["B"].width = 44
        wb.save(path)
        t = Table.read(path, "A")
        t.clear_column_width()
        t.write(path)
        assert widths(path) == {2: 44}  # the owner's 44, not ours, so it stays

    def test_hand_widths_on_other_columns_are_untouched(self, path):
        wb = load_workbook(path)
        wb["S"].column_dimensions["C"].width = 31
        wb.save(path)
        t = Table.read(path, "A")
        t.set_column_width("x", 20)
        t.write(path)
        assert widths(path) == {2: 20, 3: 31}

    def test_a_neighbour_growing_carries_the_widths_along(self, path):
        b = Table.read(path, "B")
        b.set_column_width("z", 18)
        b.write(path)
        wb = load_workbook(path)
        wb["S"].column_dimensions["E"].width = 15  # B's label column, by hand
        wb.save(path)
        assert widths(path) == {5: 15, 6: 18}
        a = Table.read(path, "A")
        a.add_column("w", [7, 8])
        a.write(path)
        assert widths(path) == {6: 15, 7: 18}  # moved one column right, nothing left behind

    def test_setting_a_width_inside_a_grouped_range_keeps_the_ranges_valid(self, path):
        wb = load_workbook(path)
        ws = wb["S"]
        ws.column_dimensions["B"] = ColumnDimension(ws, index="B", min=2, max=4, width=12)
        wb.save(path)
        t = Table.read(path, "A")
        t.set_column_width("x", 30)  # column B, the start of the group B:D
        t.set_column_width("y", 40)  # column C, inside it
        t.write(path)
        assert widths(path) == {2: 30, 3: 40, 4: 12}
        with zipfile.ZipFile(path) as z:
            sheet_xml = next(
                z.read(n).decode()
                for n in z.namelist()
                if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n) and b"<cols>" in z.read(n)
            )
        spans = sorted(
            (int(a), int(b)) for a, b in re.findall(r'<col [^>]*min="(\d+)" max="(\d+)"', sheet_xml)
        )
        assert all(spans[i][1] < spans[i + 1][0] for i in range(len(spans) - 1))  # no overlaps

    def test_sizes_merge_like_cells(self, path):
        alice, bob = Table.read(path, "A"), Table.read(path, "A")
        bob.set_column_width("x", 20)
        bob.write(path)
        alice.set_cell(row="r1", column="y", value=99)
        with warnings.catch_warnings():
            warnings.simplefilter("error", MergeConflictWarning)
            alice.write(path)
        assert alice.get_column_width("x") == 20
        assert widths(path) == {2: 20}
        carol, dave = Table.read(path, "A"), Table.read(path, "A")
        dave.set_row_height("r1", 20)
        dave.write(path)
        carol.set_row_height("r1", 50)
        with pytest.warns(MergeConflictWarning, match="height"):
            carol.write(path)
        assert heights(path) == {3: 50}


class TestFitTextOnTables:
    def fitted(self, path, **kwargs):
        t = Table.read(path, "A")
        t.style = TableStyle(fit_text=True, **kwargs)
        return t

    def test_sizes_columns_to_content_and_wraps_only_long_text(self, path):
        t = self.fitted(path)
        t.add_row("a-rather-long-row-label", [LONG, "short"])
        t.write(path)
        w = widths(path)
        assert w[1] == 26  # 23 characters + padding
        assert w[2] == 63  # the cap of 60 + padding: the 150-character text is wrapped, not widened
        assert w[3] == 9  # nothing wider than a default column
        ws = load_workbook(path)["S"]
        assert ws["B5"].alignment.wrap_text
        assert not ws["C5"].alignment.wrap_text and not ws["B3"].alignment.wrap_text

    def test_max_width_is_the_cap(self, path):
        t = self.fitted(path, max_width=20)
        t.add_row("r3", ["x" * 30, 1])
        t.write(path)
        assert widths(path)[2] == 23
        assert load_workbook(path)["S"]["B5"].alignment.wrap_text

    def test_a_line_break_wraps(self, path):
        t = self.fitted(path)
        t.set_cell(row="r1", column="x", value="one\ntwo")
        t.write(path)
        assert load_workbook(path)["S"]["B3"].alignment.wrap_text

    def test_it_refits_when_the_content_changes(self, path):
        t = self.fitted(path)
        t.set_cell(row="r1", column="x", value="y" * 40)
        t.write(path)
        assert widths(path)[2] == 43
        t.set_cell(row="r1", column="x", value=1)
        t.write(path)
        assert widths(path)[2] == 9

    def test_a_given_width_wins_over_fitting(self, path):
        t = self.fitted(path)
        t.set_column_width("x", 12)
        t.set_cell(row="r1", column="x", value="y" * 40)
        t.write(path)
        assert widths(path)[2] == 12

    def test_a_width_set_by_hand_is_left_alone(self, path):
        wb = load_workbook(path)
        wb["S"].column_dimensions["B"].width = 50
        wb.save(path)
        t = self.fitted(path)
        t.set_cell(row="r1", column="x", value="y" * 10)
        t.write(path)
        assert widths(path)[2] == 50

    def test_an_explicit_wrap_setting_wins(self, path):
        t = self.fitted(path)
        t.add_row("r3", [LONG, 1])
        t.set_format(CellFormat(wrap=False), row="r3", column="x")
        t.write(path)
        assert not load_workbook(path)["S"]["B5"].alignment.wrap_text

    def test_turning_it_off_releases_what_it_set(self, path):
        t = self.fitted(path)
        t.set_cell(row="r1", column="x", value="y" * 40)
        t.write(path)
        assert widths(path)
        t = Table.read(path, "A")
        assert t.style.fit_text  # stored with the table
        t.style = TableStyle()
        t.write(path)
        assert widths(path) == {}

    def test_it_is_remembered_with_its_cap(self, path):
        self.fitted(path, max_width=33).write(path)
        style = Table.read(path, "A").style
        assert style.fit_text and style.max_width == 33

    def test_a_neighbour_growing_keeps_a_fitted_table_fitted(self, path):
        b = Table.read(path, "B")
        b.style = TableStyle(fit_text=True)
        b.set_cell(row="r1", column="z", value="q" * 30)
        b.write(path)
        before = widths(path)
        a = Table.read(path, "A")
        a.add_column("w", [7, 8])
        a.write(path)
        after = widths(path)
        assert after[7] == before[6] and after[6] == before[5]  # moved one column right

    def test_style_validation(self):
        with pytest.raises(TypeError):
            TableStyle(fit_text=1)  # type: ignore[arg-type]
        for bad in (0, 256, True, 1.5):
            with pytest.raises(ValueError):
                TableStyle(max_width=bad)  # type: ignore[arg-type]


class TestGridSizes:
    def test_explicit_sizes(self, book):
        write_sheet(book, [["a", "b"], [1, 2]], column_widths={1: 30, 2: 12.5}, row_heights={2: 45})
        d = read_dimensions(book)
        assert d.column_widths == {1: 30, 2: 12.5} and d.row_heights == {2: 45}

    def test_columns_orientation_sizes_are_sheet_positions(self, book):
        write_sheet(book, [["a", "b"]], orientation="columns", row_heights={2: 20})
        assert read_dimensions(book).row_heights == {2: 20}

    @pytest.mark.parametrize(
        "kwargs",
        [
            {"column_widths": {0: 10}},
            {"column_widths": {1: 0}},
            {"column_widths": {"A": 10}},
            {"row_heights": {1: 500}},
            {"row_heights": {2_000_000: 10}},
            {"max_width": 0},
            {"column_widths": "wide"},
            {"row_heights": [1, 2]},
        ],
    )
    def test_bad_arguments(self, book, kwargs):
        with pytest.raises((TypeError, ValueError)):
            write_sheet(book, [[1]], **kwargs)
        with pytest.raises((TypeError, ValueError)):
            append_rows(book, [[1]], **kwargs)

    def test_fit_text(self, book):
        write_sheet(
            book,
            [["name", "notes"], ["x" * 12, LONG], ["yy", "ok"]],
            fit_text=True,
        )
        d = read_dimensions(book)
        assert d.column_widths == {1: 15, 2: 63}
        ws = load_workbook(book).active
        assert ws["B2"].alignment.wrap_text and not ws["B3"].alignment.wrap_text

    def test_fit_text_respects_given_widths_and_explicit_wrap(self, book):
        write_sheet(
            book,
            [[LONG, Styled(LONG, CellFormat(wrap=False))]],
            column_widths={1: 20},
            fit_text=True,
        )
        assert read_dimensions(book).column_widths[1] == 20
        ws = load_workbook(book).active
        assert ws["A1"].alignment.wrap_text and not ws["B1"].alignment.wrap_text

    def test_off_by_default(self, book):
        write_sheet(book, [[LONG]])
        assert read_dimensions(book).column_widths == {}

    def test_append_only_ever_grows(self, book):
        write_sheet(book, [["a"]], column_widths={1: 40})
        append_rows(book, [["short"]], fit_text=True)
        assert read_dimensions(book).column_widths == {1: 40}  # not narrowed
        append_rows(book, [["z" * 50]], fit_text=True)
        assert read_dimensions(book).column_widths == {1: 53}
        assert not load_workbook(book).active["A3"].alignment.wrap_text  # fits under the cap
        append_rows(book, [[LONG]], fit_text=True)
        assert load_workbook(book).active["A4"].alignment.wrap_text

    def test_csv_has_none_of_it(self, tmp_path):
        csv = tmp_path / "a.csv"
        create_csv(csv)
        for kwargs in ({"column_widths": {1: 10}}, {"row_heights": {1: 10}}, {"fit_text": True}):
            with pytest.raises(ValueError):
                write_sheet(csv, [[1]], **kwargs)
            with pytest.raises(ValueError):
                append_rows(csv, [[1]], **kwargs)
        with pytest.raises(ValueError):
            read_dimensions(csv)

    def test_read_dimensions_expands_a_grouped_range(self, book):
        wb = load_workbook(book)
        ws = wb.active
        ws.column_dimensions["A"] = ColumnDimension(ws, index="A", min=1, max=3, width=9)
        wb.save(book)
        assert read_dimensions(book).column_widths == {1: 9, 2: 9, 3: 9}

    def test_read_dimensions_unknown_sheet(self, book):
        from pyhandlexl import SheetNotFoundError

        with pytest.raises(SheetNotFoundError):
            read_dimensions(book, "nope")
