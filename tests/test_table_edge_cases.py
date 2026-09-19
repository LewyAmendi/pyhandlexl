"""Corners of the ``Table`` API that ordinary use doesn't reach."""

from __future__ import annotations

import copy
import datetime as dt
import pickle

import pytest

from pyhandlexl import (
    Table,
    TableExistsError,
    TableNotFoundError,
    TableStyle,
    create_sheet,
    delete_table,
    list_tables,
    table_info,
    write_sheet,
)


@pytest.fixture
def path(book):
    create_sheet(book, "D")
    create_sheet(book, "E")
    return book


def small(name: str = "T") -> Table:
    return Table(
        data=[[1, 2], [3, 4]], row_labels=["r1", "r2"], column_headers=["a", "b"], name=name
    )


class TestABareStringIsNotAListOfValues:
    """``add_row("x", "ab")`` used to be accepted as two cells, 'a' and 'b'."""

    @pytest.mark.parametrize(
        "call",
        [
            lambda t: t.add_row("new", "ab"),
            lambda t: t.insert_row(1, "new", "ab"),
            lambda t: t.set_row("r1", "ab"),
            lambda t: t.set_column("a", "ab"),
            lambda t: t.add_column("new", "ab"),
            lambda t: t.insert_column(1, "new", "ab"),
            lambda t: t.add_row("new", b"ab"),
        ],
        ids=[
            "add_row",
            "insert_row",
            "set_row",
            "set_column",
            "add_column",
            "insert_column",
            "bytes",
        ],
    )
    def test_it_is_refused_and_the_table_is_unchanged(self, call):
        table = small()
        before = (table.data.row_labels, table.data.column_headers, table.data.rows)
        with pytest.raises(TypeError, match="split it into individual characters"):
            call(table)
        assert (table.data.row_labels, table.data.column_headers, table.data.rows) == before

    def test_a_tuple_or_generator_is_fine(self):
        table = small()
        table.add_row("t", (5, 6))
        table.add_row("g", (v for v in (7, 8)))
        table.add_column("c", [0, 0, 0, 0])
        assert table.data.row_labels == ["r1", "r2", "t", "g"]
        assert table.read_row("t") == [5, 6, 0]
        assert table.read_row("g") == [7, 8, 0]


class TestDegenerateShapes:
    def test_one_by_one(self, path):
        Table(data=[[None]], row_labels=["r"], column_headers=["c"], name="One").create(
            path, sheet="D"
        )
        assert Table.read(path, "One").data.rows == [[None]]

    def test_a_header_only_table_round_trips_and_grows(self, path):
        empty = Table(column_headers=["a", "b"], name="Z", corner="k")
        empty.create(path, sheet="D")
        back = Table.read(path, "Z")
        assert back.data.rows == [] and back.data.column_headers == ["a", "b"]
        assert back.data.corner == "k"
        back.add_row("first", [1, 2])
        back.write(path)
        assert Table.read(path, "Z").data.rows == [[1, 2]]

    def test_dropping_every_row_leaves_a_valid_header_only_table(self, path):
        table = small()
        table.create(path, sheet="D")
        for label in ("r1", "r2"):
            table.drop_row(label)
        table.write(path)
        back = Table.read(path, "T")
        assert back.data.rows == [] and back.data.column_headers == ["a", "b"]
        assert table_info(path, "T").n_rows == 0

    def test_the_only_column_cannot_be_dropped(self):
        table = Table(data=[[1]], row_labels=["r"], column_headers=["only"], name="T")
        with pytest.raises(ValueError, match="only remaining column"):
            table.drop_column("only")
        assert table.data.column_headers == ["only"]

    def test_renaming_to_the_same_name_is_a_no_op(self):
        table = small()
        table.rename_row("r1", "r1")
        table.rename_column("a", "a")
        assert table.data.row_labels == ["r1", "r2"] and table.data.column_headers == ["a", "b"]

    def test_renaming_onto_an_existing_name_is_refused(self):
        table = small()
        with pytest.raises(ValueError, match="already exists"):
            table.rename_row("r1", "r2")
        with pytest.raises(ValueError, match="already exists"):
            table.rename_column("a", "b")

    def test_swapping_two_names_through_a_temporary_works(self):
        table = small()
        table.rename_row("r1", "tmp")
        table.rename_row("r2", "r1")
        table.rename_row("tmp", "r2")
        assert table.data.row_labels == ["r2", "r1"]
        assert table.read_row("r1") == [3, 4]

    @pytest.mark.parametrize("position", [0, -1, 4, 99])
    def test_insert_positions_out_of_range(self, position):
        table = small()
        with pytest.raises(IndexError):
            table.insert_row(position, "x", [0, 0])
        with pytest.raises(IndexError):
            table.insert_column(position, "x", [0, 0])

    @pytest.mark.parametrize("position", [1.0, "1", None])
    def test_non_integer_insert_positions(self, position):
        with pytest.raises(TypeError):
            small().insert_row(position, "x", [0, 0])


class TestTableNames:
    @pytest.mark.parametrize(
        "name",
        [
            "TABLE NAME",
            "two\nlines",
            "tab\there",
            " spaced ",
            "é 日本語 😀",
            "n" * 5_000,
            "D",
            "a/b\\c",
            "'quoted'",
            "=formula-like",
            "0",
        ],
        ids=lambda n: repr(n)[:30],
    )
    def test_any_reasonable_name_round_trips(self, path, name):
        Table(data=[[1]], row_labels=["r"], column_headers=["c"], name=name).create(path, sheet="D")
        assert Table.read(path, name).name == name
        assert name in list_tables(path)

    def test_names_that_differ_only_by_case_are_different_tables(self, path):
        Table(data=[[1]], row_labels=["r"], column_headers=["c"], name="sales").create(
            path, sheet="D"
        )
        Table(data=[[2]], row_labels=["r"], column_headers=["c"], name="SALES").create(
            path, sheet="D"
        )
        assert Table.read(path, "sales").read_row("r") == [1]
        assert Table.read(path, "SALES").read_row("r") == [2]

    def test_a_name_is_unique_across_sheets(self, path):
        small().create(path, sheet="D")
        with pytest.raises(TableExistsError):
            small().create(path, sheet="E")

    def test_an_empty_or_non_string_name_is_refused(self):
        with pytest.raises(ValueError):
            Table(column_headers=["a"], name="")
        with pytest.raises(TypeError):
            Table(column_headers=["a"], name=5)  # type: ignore[arg-type]

    def test_a_name_reused_after_delete_is_a_fresh_table(self, path):
        small().create(path, sheet="D")
        delete_table(path, "T")
        Table(data=[[9]], row_labels=["z"], column_headers=["q"], name="T").create(path, sheet="D")
        assert Table.read(path, "T").data.column_headers == ["q"]


class TestPlacementAfterDeletes:
    def test_a_new_table_lands_beyond_the_tables_that_remain(self, path):
        for name in ("A", "B", "C"):
            small(name).create(path, sheet="D")
        delete_table(path, "B")  # leaves a gap; the others don't move
        small("N").create(path, sheet="D")
        assert list_tables(path) == ["A", "C", "N"]
        for name in ("A", "C", "N"):
            assert Table.read(path, name).data.rows == [[1, 2], [3, 4]]
        starts = sorted(table_info(path, n).sheet for n in ("A", "C", "N"))
        assert starts == ["D", "D", "D"]

    def test_deleting_the_rightmost_table_lets_the_next_one_reuse_its_place(self, path):
        small("A").create(path, sheet="D")
        small("B").create(path, sheet="D")
        delete_table(path, "B")
        small("B2").create(path, sheet="D")
        assert Table.read(path, "A").data.rows == [[1, 2], [3, 4]]
        assert Table.read(path, "B2").data.rows == [[1, 2], [3, 4]]

    def test_growing_a_left_table_carries_the_style_of_the_ones_it_pushes(self, path):
        small("L").create(path, sheet="D")
        right = Table(
            data=[[1]],
            row_labels=["r"],
            column_headers=["c"],
            name="R",
            style=TableStyle.NONE,
        )
        right.create(path, sheet="D")
        left = Table.read(path, "L")
        for i in range(3):
            left.add_column(f"extra{i}", [0, 0])
        left.write(path)
        assert Table.read(path, "R").style == TableStyle.NONE
        assert Table.read(path, "R").read_row("r") == [1]

    def test_a_grid_sheet_and_a_table_sheet_never_mix(self, path):
        write_sheet(path, [["g"]], sheet="E")
        from pyhandlexl import SheetKindError

        with pytest.raises(SheetKindError):
            small().create(path, sheet="E")


class TestCopiesAreIndependent:
    def test_two_reads_are_independent_objects(self, path):
        small().create(path, sheet="D")
        a, b = Table.read(path, "T"), Table.read(path, "T")
        a.set_cell(row="r1", column="a", value="changed")
        assert b.read_cell(row="r1", column="a") == 1

    def test_the_data_snapshot_is_detached(self):
        table = small()
        snapshot = table.data
        snapshot.rows[0][0] = "hacked"
        snapshot.row_labels.append("x")
        snapshot.column_headers.clear()
        assert table.data.rows[0][0] == 1
        assert table.data.row_labels == ["r1", "r2"] and table.data.column_headers == ["a", "b"]

    def test_to_dict_is_detached(self):
        table = small()
        out = table.to_dict()
        out["r1"]["a"] = "hacked"
        assert table.read_cell(row="r1", column="a") == 1

    def test_deepcopy_and_copy_and_pickle(self, path):
        small().create(path, sheet="D")
        original = Table.read(path, "T")
        for clone in (copy.deepcopy(original), pickle.loads(pickle.dumps(original))):
            assert clone == original
            clone.set_cell(row="r1", column="a", value="changed")
            assert original.read_cell(row="r1", column="a") == 1
            clone.write(path)  # a clone still knows how to merge and write
        assert Table.read(path, "T").read_cell(row="r1", column="a") == "changed"

    def test_the_constructor_copies_its_arguments(self):
        rows, labels, headers = [[1, 2]], ["r"], ["a", "b"]
        table = Table(rows, labels, column_headers=headers, name="T")
        rows[0][0] = "x"
        labels.append("y")
        headers.append("z")
        assert table.data.rows == [[1, 2]] and table.data.row_labels == ["r"]
        assert table.data.column_headers == ["a", "b"]


class TestEquality:
    def test_compares_content_not_identity_and_ignores_style_and_name(self):
        a = small("one")
        b = small("two")
        b.style = TableStyle.NONE
        assert a == b
        b.set_cell(row="r1", column="a", value=99)
        assert a != b

    def test_compares_column_types_and_corner(self):
        from pyhandlexl import ColumnType

        a, b = small(), small()
        b.set_column_type("a", ColumnType.NUMBER)
        assert a != b
        b.set_column_type("a", ColumnType.ANY)
        assert a == b
        b.set_corner("k")
        assert a != b

    def test_a_table_is_not_equal_to_other_types(self):
        assert small() != "table"
        assert small() != [[1, 2], [3, 4]]
        assert (small() == None) is False  # noqa: E711


class TestMetadata:
    def test_times_are_utc_stable_and_monotonic(self, path):
        table = small()
        assert table.info.created_at is None and table.info.sheet is None
        table.create(path, sheet="D")
        created = table.info.created_at
        assert created.tzinfo is not None and created.utcoffset() == dt.timedelta(0)
        stamps = [table.info.modified_at]
        for i in range(3):
            table.set_cell(row="r1", column="a", value=i)
            table.write(path)
            stamps.append(table.info.modified_at)
        assert stamps == sorted(stamps)
        assert table.info.created_at == created  # never moves
        on_disk = table_info(path, "T")
        assert on_disk.created_at == created
        assert on_disk.modified_at == stamps[-1]

    def test_info_matches_the_table_and_the_sheet(self, path):
        small().create(path, sheet="E")
        info = Table.read(path, "T").info
        assert (info.n_rows, info.n_cols, info.sheet) == (2, 2, "E")
        assert table_info(path, "T") == info

    def test_writing_twice_with_no_change_keeps_the_data(self, path):
        table = small()
        table.create(path, sheet="D")
        table.write(path)
        table.write(path)
        assert Table.read(path, "T") == small()


class TestNotFound:
    def test_reading_a_name_that_is_not_a_table(self, path):
        write_sheet(path, [["g"]], sheet="E")
        for name in ("Nope", "E", "D", "Sheet", ""):
            with pytest.raises(TableNotFoundError):
                Table.read(path, name)

    def test_writing_a_table_that_was_never_created(self, path):
        with pytest.raises(TableNotFoundError):
            small().write(path)

    def test_the_error_names_the_table(self, path):
        with pytest.raises(TableNotFoundError, match="Ghost"):
            Table.read(path, "Ghost")


class TestShowAndRepr:
    def test_show_survives_odd_content(self, capsys):
        table = Table(
            data=[
                [None, "x" * 60, "日本語", "multi\nline"],
                [1.5, True, dt.datetime(2026, 1, 1), ""],
            ],
            row_labels=["😀", "a\tb"],
            column_headers=["h0", "h" * 40, "é", "d"],
            name="Odd",
        )
        table.show()
        table.show(rows=1)
        table.show(head=0, tail=0)
        table.show(head=None, tail=None)
        assert "日本語" in capsys.readouterr().out

    def test_repr_never_raises(self):
        for name in ("plain", "two\nlines", "é", "n" * 200):
            assert isinstance(repr(Table(column_headers=["a"], name=name)), str)
